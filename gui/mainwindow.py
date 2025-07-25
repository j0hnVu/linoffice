# This Python file uses the following encoding: utf-8
import sys
from PySide6.QtWidgets import QApplication, QWidget, QMainWindow, QMessageBox
from PySide6.QtUiTools import QUiLoader
from PySide6.QtCore import QFile, QTimer
import subprocess
import os

LINOFFICE_SCRIPT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'linoffice.sh'))
SETUP_SCRIPT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'setup.sh'))
UNINSTALL_SCRIPT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'uninstall.sh'))

class MainWindow(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.load_ui('main.ui')
        self.setWindowTitle(self.ui.windowTitle())
        self.connect_buttons()
        self.update_container_status()
        # Set up a timer to update container status every 30 seconds
        self.status_timer = QTimer(self)
        self.status_timer.timeout.connect(self.update_container_status)
        self.status_timer.start(30000)  # 30,000 ms = 30 seconds

    def load_ui(self, ui_file):
        loader = QUiLoader()
        file = QFile(ui_file)
        file.open(QFile.ReadOnly)
        self.ui = loader.load(file, self)
        file.close()

    # Connect the buttons to functions
    def connect_buttons(self):
        self.ui.pushButton_settings.clicked.connect(self.open_settings_window)
        self.ui.pushButton_tools.clicked.connect(self.open_tools_window)
        self.ui.pushButton_troubleshooting.clicked.connect(self.open_troubleshooting_window)
        # Connect app launch buttons
        self.ui.pushButton_word.clicked.connect(lambda: self.launch_linoffice_app('word'))
        self.ui.pushButton_excel.clicked.connect(lambda: self.launch_linoffice_app('excel'))
        self.ui.pushButton_powerpoint.clicked.connect(lambda: self.launch_linoffice_app('powerpoint'))
        self.ui.pushButton_outlook.clicked.connect(lambda: self.launch_linoffice_app('outlook'))
        self.ui.pushButton_onenote.clicked.connect(lambda: self.launch_linoffice_app('onenote'))

    # Functions to open secondary windows
    def open_settings_window(self):
        self.settings_window = SettingsWindow()
        self.settings_window.show()

    def open_tools_window(self):
        self.tools_window = ToolsWindow(main_window=self)
        self.tools_window.show()

    def open_troubleshooting_window(self):
        self.troubleshooting_window = TroubleshootingWindow()
        self.troubleshooting_window.show()

    def launch_linoffice_app(self, *args):
        subprocess.Popen([LINOFFICE_SCRIPT, *args])

    def update_container_status(self):
        try:
            result = subprocess.run(['podman', 'ps', '--filter', 'name=LinOffice', '--format', '{{.Status}}'], capture_output=True, text=True, check=True)
            status = result.stdout.strip()
            if status:
                status_text = f"Container: running ({status})"
            else:
                status_text = "Container: not running"
        except Exception as e:
            status_text = f"Container: error"
        self.ui.label.setText(status_text)

# Defining secondary windows
class SettingsWindow(QMainWindow):
    def __init__(self, parent=None):
        super(SettingsWindow, self).__init__(parent)
        self.load_ui('settings.ui')
        self.setWindowTitle(self.ui.windowTitle())

    def load_ui(self, ui_file):
        loader = QUiLoader()
        file = QFile(ui_file)
        file.open(QFile.ReadOnly)
        self.ui = loader.load(file, self)
        file.close()

class ToolsWindow(QMainWindow):
    def __init__(self, parent=None, main_window=None):
        super(ToolsWindow, self).__init__(parent)
        self.main_window = main_window
        self.load_ui('tools.ui')
        self.setWindowTitle(self.ui.windowTitle())
        self.connect_tools_buttons()

    def load_ui(self, ui_file):
        loader = QUiLoader()
        file = QFile(ui_file)
        file.open(QFile.ReadOnly)
        self.ui = loader.load(file, self)
        file.close()

    # Connect buttons in tools window with LinOffice script
    def connect_tools_buttons(self):
        if self.main_window:
            self.ui.pushButton_update.clicked.connect(lambda: self.main_window.launch_linoffice_app('update'))
            self.ui.pushButton_powershell.clicked.connect(lambda: self.main_window.launch_linoffice_app('manual', 'powershell.exe'))
            self.ui.pushButton_regedit.clicked.connect(lambda: self.main_window.launch_linoffice_app('manual', 'regedit.exe'))
            self.ui.pushButton_cmd.clicked.connect(lambda: self.main_window.launch_linoffice_app('manual', 'cmd.exe'))
            self.ui.pushButton_explorer.clicked.connect(lambda: self.main_window.launch_linoffice_app('manual', 'explorer.exe'))

class TroubleshootingWindow(QMainWindow):
    def __init__(self, parent=None):
        super(TroubleshootingWindow, self).__init__(parent)
        self.load_ui('troubleshooting.ui')
        self.setWindowTitle(self.ui.windowTitle())
        self.connect_troubleshooting_buttons()

    def load_ui(self, ui_file):
        loader = QUiLoader()
        file = QFile(ui_file)
        file.open(QFile.ReadOnly)
        self.ui = loader.load(file, self)
        file.close()

    def connect_troubleshooting_buttons(self):
        self.ui.pushButton_lockfiles.clicked.connect(self.run_cleanup_full)
        self.ui.pushButton_desktopfiles.clicked.connect(self.run_setup_desktop)
        self.ui.pushButton_reset.clicked.connect(self.run_reset)
        self.ui.pushButton_logfile.clicked.connect(self.open_logfile)
        self.ui.pushButton_website.clicked.connect(self.open_website)
        self.ui.pushButton_uninstall.clicked.connect(self.run_uninstall)

    def run_cleanup_full(self):
        subprocess.Popen([LINOFFICE_SCRIPT, 'cleanup', '--full'])

    def run_setup_desktop(self):
        subprocess.Popen([SETUP_SCRIPT, '--desktop'])

    def run_reset(self):
        subprocess.Popen([LINOFFICE_SCRIPT, 'reset'])

    def open_logfile(self):
        logfile = os.path.expanduser('~/.local/share/linoffice/linoffice.log')
        # Try to open with xdg-open (Linux default)
        subprocess.Popen(['xdg-open', logfile])

    def open_website(self):
        import webbrowser
        webbrowser.open('https://github.com/eylenburg/linoffice')

    def run_uninstall(self):
        reply = QMessageBox.question(self, 'Confirm Uninstall',
                                     'Are you sure you want to uninstall LinOffice?',
                                     QMessageBox.Yes | QMessageBox.Cancel, QMessageBox.Cancel)
        if reply == QMessageBox.Yes:
            # Try to open in a new terminal window (x-terminal-emulator, gnome-terminal, konsole)
            terminal_cmds = [
                ['x-terminal-emulator', '-e', UNINSTALL_SCRIPT],
                ['konsole', '-e', UNINSTALL_SCRIPT],
                ['gnome-terminal', '--', UNINSTALL_SCRIPT],
                ['xfce4-terminal', '-e', UNINSTALL_SCRIPT],
                ['lxterminal', '-e', UNINSTALL_SCRIPT],
                ['mate-terminal', '-e', UNINSTALL_SCRIPT],
                ['xterm', '-e', UNINSTALL_SCRIPT],
            ]
            for cmd in terminal_cmds:
                try:
                    subprocess.Popen(cmd)
                    break
                except FileNotFoundError:
                    continue

if __name__ == "__main__":
    app = QApplication(sys.argv)
    widget = MainWindow()
    widget.show()
    sys.exit(app.exec())
