import sys
import re
from PyQt5.QtWidgets import QApplication, QPushButton, QMessageBox, QWidget, QLabel, QSizePolicy, QToolTip
from PyQt5.QtCore import Qt, QTimer, QAbstractNativeEventFilter, QVariantAnimation, QMimeData, QPoint
from PyQt5.QtGui import QScreen, QPixmap, QPainter, QImage, QColor, QIcon, QFont, QDrag
from PyQt5.QtWinExtras import QtWin
import ctypes
from ctypes import wintypes, windll, byref
import psutil  # to list the currently opened windows
import win32gui  # to interact with windows
import win32process  # to get process info of windows
import win32api
import win32con

TASKBAR_SIZE = 96
BUTTON_HEIGHT = 32

ASFW_ANY = -1

class APPBARDATA(ctypes.Structure):
    _fields_ = [
        ("cbSize", wintypes.DWORD),
        ("hWnd", wintypes.HWND),
        ("uCallbackMessage", wintypes.UINT),
        ("uEdge", wintypes.UINT),
        ("rc", wintypes.RECT),
        ("lParam", wintypes.LPARAM),
    ]

def get_primary_screen_geometry(app):
    primary_screen = app.primaryScreen()
    return primary_screen.availableGeometry()

class DraggableButton(QPushButton):
    def __init__(self, title, parent):
        super().__init__(title, parent)
        self.setAcceptDrops(True)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.drag_start_position = event.pos()
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if not (event.buttons() & Qt.LeftButton):
            return
        if (event.pos() - self.drag_start_position).manhattanLength() < QApplication.startDragDistance():
            return

        drag = QDrag(self)
        mime_data = QMimeData()
        mime_data.setText(self.text())  # Store button text to help identify during drop
        drag.setMimeData(mime_data)
        drag.setHotSpot(event.pos() - self.rect().topLeft())

        drop_action = drag.exec_(Qt.MoveAction)

    def dragEnterEvent(self, event):
        event.acceptProposedAction()

    def dropEvent(self, event):
        # Notify the main application to handle the button swap
        self.parent().swap_buttons(self, event.source())
        event.acceptProposedAction()

class ShellHookListener(QAbstractNativeEventFilter):
    def __init__(self, main_window):
        super().__init__()
        self.main_window = main_window

    def nativeEventFilter(self, event_type, message):
        if event_type == "windows_generic_MSG":
            msg = ctypes.wintypes.MSG.from_address(int(message))
            if msg.message == self.main_window.WM_SHELLHOOKMESSAGE:
                if msg.wParam in [self.main_window.HSHELL_WINDOWCREATED, self.main_window.HSHELL_WINDOWDESTROYED, self.main_window.HSHELL_WINDOWTITLECHANGE]:
                    # print(f"Shell message received: wParam={msg.wParam}")  # Debugging output
                    self.main_window.update_taskbar_buttons()
        return False, 0

class FixedWindowApp(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.register_app_bar()

        self.pre_top_process_id = -1

        # Use QTimer to periodically update taskbar buttons to ensure consistency
        self.update_timer = QTimer(self)
        self.update_timer.timeout.connect(self.update_taskbar_buttons)
        self.update_timer.start(1000)  # Update every 1 second to ensure buttons reflect current state

        # Use QTimer to delay the move operation slightly
        QTimer.singleShot(100, self.move_to_left)

        # Register the shell hook to listen to window events
        self.setup_shell_hook()

    def initUI(self):
        # Set window title (optional, as window doesn't have a title bar)
        self.setWindowTitle('固定左側窗口')
        
        # Remove window frame, making it impossible to resize or move
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint | Qt.Tool)
        
        self.setFixedSize(TASKBAR_SIZE, get_primary_screen_geometry(QApplication.instance()).height())

        # Set a darkened background
        self.set_darkened_background()

        # Create a button to simulate pressing the Windows key
        self.windows_key_button = DraggableButton('Windows鍵', self)
        self.windows_key_button.setGeometry(0, 0, TASKBAR_SIZE, BUTTON_HEIGHT)
        self.windows_key_button.clicked.connect(self.press_windows_key)
        self.add_hover_animation(self.windows_key_button)

        # Create a button to close the application
        self.close_button = DraggableButton('Close', self)
        self.close_button.setToolTip("Your tooltip text")
        self.close_button.setGeometry(0, BUTTON_HEIGHT * 1, TASKBAR_SIZE, BUTTON_HEIGHT)
        self.close_button.clicked.connect(self.close_app)
        self.add_hover_animation(self.close_button)

        # Dictionary to keep track of dynamically created buttons
        self.taskbar_buttons = {}

        # Add buttons for each window in the taskbar
        self.add_taskbar_buttons()

    def swap_buttons(self, target_button, source_button):
        # Get the geometry of both buttons
        target_geometry = target_button.geometry()
        source_geometry = source_button.geometry()

        # Swap the positions of the target and source buttons
        target_button.setGeometry(source_geometry)
        source_button.setGeometry(target_geometry)

        # Update button order in the taskbar_buttons dictionary
        hwnd_source = next((hwnd for hwnd, button in self.taskbar_buttons.items() if button == source_button), None)
        hwnd_target = next((hwnd for hwnd, button in self.taskbar_buttons.items() if button == target_button), None)

        if hwnd_source and hwnd_target:
            # Swap the dictionary values
            self.taskbar_buttons[hwnd_source], self.taskbar_buttons[hwnd_target] = self.taskbar_buttons[hwnd_target], self.taskbar_buttons[hwnd_source]

    def add_hover_animation(self, button):
        # Adjust text to show custom ellipsis (~) if too long, considering icon size
        font_metrics = button.fontMetrics()
        icon_width = button.iconSize().width() if not button.icon().isNull() else 0
        padding = 15  # Include some padding for better visual spacing
        available_width = button.width() - icon_width - padding
        button.setToolTip(button.text())
        if font_metrics.width(button.text()) > available_width:
            elided_text = button.text()
            while font_metrics.width(elided_text + "...") > available_width and len(elided_text) > 0:
                elided_text = elided_text[:-1]
            elided_text += "..."
            button.setText(elided_text)
        # Adjust text to show ellipsis if too long
        font_metrics = button.fontMetrics()
        elided_text = font_metrics.elidedText(button.text(), Qt.ElideRight, button.width() - 10)  # 10 for padding
        button.setText(elided_text)
        button.setStyleSheet("QPushButton{ background-color: navy; color: white; border: none; border-top: 1px solid gray; border-bottom: 1px solid gray; padding-left: 5px; text-align: left;} QToolTip{background-color: white; color: black; border: 1px solid black;}")
        button.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        button.setStyleSheet(button.styleSheet() + " text-align: left; padding-left: 5px; white-space: nowrap; ")
        
        original_color = QColor("navy")
        hover_color = QColor("red")

        def enter_event(event):
            animation = QVariantAnimation()
            animation.setDuration(300)
            animation.setStartValue(original_color)
            animation.setEndValue(hover_color)
            animation.valueChanged.connect(lambda color: button.setStyleSheet("QPushButton { background-color: "+color.name()+"; color: white; border: none; border-top: 1px solid gray; border-bottom: 1px solid gray; padding-left: 5px; text-align: left;} QToolTip {background-color: white; color: black; border: 1px solid black; }"))
            animation.start()
            QToolTip.showText(button.mapToGlobal(button.rect().center()), button.toolTip(), button)
            button.animation = animation  # Keep a reference to avoid garbage collection

        def leave_event(event):
            animation = QVariantAnimation()
            animation.setDuration(300)
            animation.setStartValue(hover_color)
            animation.setEndValue(original_color)
            animation.valueChanged.connect(lambda color: button.setStyleSheet("QPushButton { background-color: "+color.name()+"; color: white; border: none; border-top: 1px solid gray; border-bottom: 1px solid gray; padding-left: 5px; text-align: left;} QToolTip {background-color: white; color: black; border: 1px solid black;}"))
            animation.start()
            button.animation = animation  # Keep a reference to avoid garbage collection

        button.enterEvent = enter_event
        button.leaveEvent = leave_event

    def set_darkened_background(self):
        # Capture the current screen
        self.setAttribute(Qt.WA_TranslucentBackground)
        screen = QApplication.primaryScreen()
        screenshot = screen.grabWindow(0)

        # Crop the screenshot to match the window's geometry
        screen_geometry = get_primary_screen_geometry(QApplication.instance())
        screenshot = screenshot.copy(0, screen_geometry.top(), TASKBAR_SIZE, screen_geometry.height())

        # Convert QPixmap to QImage for processing
        image = screenshot.toImage()

        # Create a darkened version of the image (reduce brightness by 40%)
        darkened_image = QImage(image.size(), QImage.Format_ARGB32)
        painter = QPainter(darkened_image)
        painter.setCompositionMode(QPainter.CompositionMode_Multiply)
        painter.fillRect(darkened_image.rect(), QColor(0, 0, 0, 196))  # transparency black fill
        painter.end()

        # Set the darkened image as the background
        darkened_label = QLabel(self)
        darkened_label.setGeometry(0, 0, TASKBAR_SIZE, screen_geometry.height())
        darkened_pixmap = QPixmap.fromImage(darkened_image)
        darkened_label.setPixmap(darkened_pixmap)
        darkened_label.lower()  # Make sure the darkened background is behind other widgets

    def get_window_icon(self, hwnd):
        # Try to get the icon of the window using WM_GETICON
        icon_handle = win32gui.SendMessage(hwnd, win32con.WM_GETICON, win32con.ICON_SMALL, 0)
        if icon_handle == 0:
            # If no icon found, try to get the large icon
            icon_handle = win32gui.SendMessage(hwnd, win32con.WM_GETICON, win32con.ICON_BIG, 0)
        if icon_handle == 0:
            # If still no icon, get the class icon
            icon_handle = ctypes.windll.user32.GetClassLongPtrW(hwnd, -14)

        if icon_handle != 0:
            # Convert the icon handle to QPixmap using QtWin
            icon_pixmap = QtWin.fromHICON(icon_handle)
            ctypes.windll.user32.DestroyIcon(icon_handle)
            return icon_pixmap
        return None
    def add_taskbar_buttons(self):
        current_y = BUTTON_HEIGHT * 3  # Start below the existing buttons
        hwnd_list = self.get_taskbar_windows()
        for hwnd, title in hwnd_list:
            ex_style = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
            if not (ex_style & win32con.WS_EX_TOOLWINDOW):
                button = QPushButton(title, self)
                # Get the window icon and set it to the button
                icon_pixmap = self.get_window_icon(hwnd)
                if icon_pixmap:
                    button.setIcon(QIcon(icon_pixmap))
                button.setGeometry(0, current_y, TASKBAR_SIZE, BUTTON_HEIGHT)
                button.clicked.connect(lambda checked, hwnd=hwnd: self.toggle_window(hwnd))
                self.add_hover_animation(button)
                self.taskbar_buttons[hwnd] = button
                current_y += BUTTON_HEIGHT

    def get_taskbar_windows(self):
        def enum_windows_callback(hwnd, hwnd_list):
            # Filter only normal, visible windows with titles
            if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd):
                # print(f"Found: {win32gui.GetWindowText(hwnd)}")
                hwnd_list.append(hwnd)
            return True

        hwnd_list = []
        win32gui.EnumWindows(enum_windows_callback, hwnd_list)
        # print("-------------------------------------------------");
        return [(hwnd, win32gui.GetWindowText(hwnd)) for hwnd in hwnd_list]

    def update_taskbar_buttons(self):
        hwnd_list = [hwnd for hwnd, _ in self.get_taskbar_windows()]

        # Update titles for existing windows if they have changed
        for hwnd in self.taskbar_buttons:
            new_title = win32gui.GetWindowText(hwnd)
            button = self.taskbar_buttons[hwnd]
            old_title = button.text()
            old_title = re.sub(r"\.\.\.$", "", old_title)
            if not new_title.startswith(old_title):
                button.setToolTip(new_title)
                font_metrics = button.fontMetrics()
                icon_width = button.iconSize().width() if not button.icon().isNull() else 0
                padding = 15  # Include some padding for better visual spacing
                available_width = button.width() - icon_width - padding

                if font_metrics.width(new_title) > available_width:
                    elided_text = new_title
                    while font_metrics.width(elided_text + "...") > available_width and len(elided_text) > 0:
                        elided_text = elided_text[:-1]
                    elided_text += "..."
                    button.setText(elided_text)

        # Add buttons for newly opened windows
        for hwnd, title in self.get_taskbar_windows():
            ex_style = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
            if hwnd not in self.taskbar_buttons and win32gui.IsWindowVisible(hwnd) and not (ex_style & win32con.WS_EX_TOOLWINDOW):
                button = QPushButton(title, self)
                # Get the window icon and set it to the button
                icon_pixmap = self.get_window_icon(hwnd)
                if icon_pixmap:
                    button.setIcon(QIcon(icon_pixmap))
                button.setGeometry(0, 0, TASKBAR_SIZE, BUTTON_HEIGHT)  # 初始位置隨意設置，稍後重新排列
                button.clicked.connect(lambda checked, hwnd=hwnd: self.toggle_window(hwnd))
                button.show()
                self.add_hover_animation(button)
                self.taskbar_buttons[hwnd] = button

        # Rearrange all taskbar buttons to ensure they are in the correct order
        current_y = BUTTON_HEIGHT * 3  # Start below the existing static buttons
        for hwnd in self.taskbar_buttons:
            button = self.taskbar_buttons[hwnd]
            button.setGeometry(0, current_y, TASKBAR_SIZE, BUTTON_HEIGHT)
            current_y += BUTTON_HEIGHT

    def toggle_window(self, hwnd):
        # Toggle the specified window between minimized and foreground
        try:
            process_id = win32process.GetWindowThreadProcessId(hwnd)
            if process_id == self.pre_top_process_id:
                # If the window is already in the foreground or is the second window, minimize it
                win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
                self.pre_top_process_id = -1
            else:
                self.pre_top_process_id = process_id
                # If the window is not in the foreground, bring it to the front
                if win32gui.IsIconic(hwnd):
                    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                win32gui.SetForegroundWindow(hwnd)
                win32gui.BringWindowToTop(hwnd)
        except Exception as e:
            print(f"Failed to toggle window {hwnd}: {e}")

    def close_app(self):
        self.unregister_app_bar()
        QApplication.instance().quit()

    def move_to_left(self):
        # Move the window to the left edge of the screen
        screen_geometry = get_primary_screen_geometry(QApplication.instance())
        self.move(0, screen_geometry.top())

    def press_windows_key(self):
        # Simulate pressing the Windows key
        ctypes.windll.user32.keybd_event(0x5B, 0, 0, 0)  # Press the Windows key (0x5B)
        ctypes.windll.user32.keybd_event(0x5B, 0, 2, 0)  # Release the Windows key (0x5B)

    def register_app_bar(self):
        # Register the app as an AppBar to reserve screen space
        self.appbar_data = APPBARDATA()
        self.appbar_data.cbSize = ctypes.sizeof(APPBARDATA)
        self.appbar_data.hWnd = int(self.winId())
        self.appbar_data.uEdge = 0  # Left side of the screen

        screen_geometry = get_primary_screen_geometry(QApplication.instance())
        self.appbar_data.rc.top = screen_geometry.top()
        self.appbar_data.rc.left = 0
        self.appbar_data.rc.right = self.appbar_data.rc.left + TASKBAR_SIZE
        self.appbar_data.rc.bottom = screen_geometry.bottom()

        # Register the AppBar with ABM_NEW
        ctypes.windll.shell32.SHAppBarMessage(0x00000000, ctypes.byref(self.appbar_data))

        # Modify the AppBar position and settings with ABM_QUERYPOS and ABM_SETPOS
        ctypes.windll.shell32.SHAppBarMessage(0x00000002, ctypes.byref(self.appbar_data))  # ABM_QUERYPOS
        ctypes.windll.shell32.SHAppBarMessage(0x00000003, ctypes.byref(self.appbar_data))  # ABM_SETPOS

    def unregister_app_bar(self):
        # Remove the AppBar reservation when closing the app
        if hasattr(self, 'appbar_data'):
            # Reset the reserved area to allow windows to occupy the space again
            self.appbar_data.rc.top = 0
            self.appbar_data.rc.left = 0
            self.appbar_data.rc.right = 0
            self.appbar_data.rc.bottom = 0
            ctypes.windll.shell32.SHAppBarMessage(0x00000003, ctypes.byref(self.appbar_data))  # ABM_SETPOS to reset
            ctypes.windll.shell32.SHAppBarMessage(0x00000001, ctypes.byref(self.appbar_data))  # ABM_REMOVE

    def setup_shell_hook(self):
        user32 = ctypes.windll.user32
        self.hWnd = int(self.winId())

        # Register to receive shell hook messages
        self.WM_SHELLHOOKMESSAGE = user32.RegisterWindowMessageW("SHELLHOOK")
        self.HSHELL_WINDOWCREATED = 0x0001
        self.HSHELL_WINDOWDESTROYED = 0x0002
        self.HSHELL_WINDOWTITLECHANGE = 0x000C  # Message ID for window title change
        if not user32.RegisterShellHookWindow(self.hWnd):
            print("Failed to register shell hook window.")  # Debugging output

        # Create a shell hook listener and install it
        self.shell_hook_listener = ShellHookListener(self)
        QApplication.instance().installNativeEventFilter(self.shell_hook_listener)
        # print(f"Shell hook registered with message ID: {self.WM_SHELLHOOKMESSAGE}")  # Debugging output

if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyleSheet("""
        QToolTip { 
            background-color: white; 
            color: black; 
            border: 1px solid black; 
        }
    """)
    
    mainWin = FixedWindowApp()
    mainWin.show()
    sys.exit(app.exec_())

