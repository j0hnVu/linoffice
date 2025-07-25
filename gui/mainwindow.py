# This Python file uses the following encoding: utf-8
import sys
from PySide6.QtWidgets import QApplication, QWidget, QMainWindow
from PySide6.QtUiTools import QUiLoader
from PySide6.QtCore import QFile
import subprocess

class MainWindow(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.load_ui('main.ui')
        self.setWindowTitle(self.ui.windowTitle())
        self.connect_buttons()
        self.update_container_status()

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
        self.tools_window = ToolsWindow()
        self.tools_window.show()

    def open_troubleshooting_window(self):
        self.troubleshooting_window = TroubleshootingWindow()
        self.troubleshooting_window.show()

    def launch_linoffice_app(self, app_name):
        subprocess.Popen(['../linoffice.sh', app_name])

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
    def __init__(self, parent=None):
        super(ToolsWindow, self).__init__(parent)
        self.load_ui('tools.ui')
        self.setWindowTitle(self.ui.windowTitle())

    def load_ui(self, ui_file):
        loader = QUiLoader()
        file = QFile(ui_file)
        file.open(QFile.ReadOnly)
        self.ui = loader.load(file, self)
        file.close()

class TroubleshootingWindow(QMainWindow):
    def __init__(self, parent=None):
        super(TroubleshootingWindow, self).__init__(parent)
        self.load_ui('troubleshooting.ui')
        self.setWindowTitle(self.ui.windowTitle())

    def load_ui(self, ui_file):
        loader = QUiLoader()
        file = QFile(ui_file)
        file.open(QFile.ReadOnly)
        self.ui = loader.load(file, self)
        file.close()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    widget = MainWindow()
    widget.show()
    sys.exit(app.exec())
