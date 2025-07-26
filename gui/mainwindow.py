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
        self.settings_changed = False
        self.load_current_settings()
        self.connect_settings_buttons()

    def load_ui(self, ui_file):
        loader = QUiLoader()
        file = QFile(ui_file)
        file.open(QFile.ReadOnly)
        self.ui = loader.load(file, self)
        file.close()

    def connect_settings_buttons(self):
        # Connect the set language button
        self.ui.pushButton_setlang.clicked.connect(self.run_setlang)
        
        # Connect OK and Cancel buttons
        self.ui.pushButton_ok.clicked.connect(self.save_settings)
        self.ui.pushButton_cancel.clicked.connect(self.close)
        
        # Connect settings change signals to track modifications
        self.ui.checkBox_suspend.toggled.connect(self.mark_settings_changed)
        self.ui.comboBox_scaling.currentTextChanged.connect(self.mark_settings_changed)
        self.ui.comboBox_date.currentTextChanged.connect(self.mark_settings_changed)
        self.ui.comboBox_decimalseparator.currentTextChanged.connect(self.mark_settings_changed)
        self.ui.lineEdit_currency.textChanged.connect(self.mark_settings_changed)

    def mark_settings_changed(self):
        self.settings_changed = True

    def load_current_settings(self):
        """Load current settings from config files"""
        try:
            import re
            # Load linoffice.conf settings
            linoffice_conf_path = os.path.join(os.path.dirname(LINOFFICE_SCRIPT), 'config', 'linoffice.conf')
            if os.path.exists(linoffice_conf_path):
                with open(linoffice_conf_path, 'r') as f:
                    content = f.read()
                    # Set autopause checkbox
                    if 'AUTOPAUSE="on"' in content:
                        self.ui.checkBox_suspend.setChecked(True)
                    elif 'AUTOPAUSE="off"' in content:
                        self.ui.checkBox_suspend.setChecked(False)
                    
                    # Set scaling combobox
                    if 'RDP_SCALE="100"' in content:
                        self.ui.comboBox_scaling.setCurrentText("100%")
                    elif 'RDP_SCALE="140"' in content:
                        self.ui.comboBox_scaling.setCurrentText("140%")
                    elif 'RDP_SCALE="180"' in content:
                        self.ui.comboBox_scaling.setCurrentText("180%")

            # Load registry_override.conf settings
            registry_conf_path = os.path.join(os.path.dirname(LINOFFICE_SCRIPT), 'config', 'registry_override.conf')
            if os.path.exists(registry_conf_path):
                with open(registry_conf_path, 'r') as f:
                    content = f.read()
                    # Set date format
                    date_match = re.search(r'DATE_FORMAT="([^"]*)"', content)
                    if date_match and date_match.group(1):
                        self.ui.comboBox_date.setCurrentText(date_match.group(1))
                    
                    # Set decimal separator
                    decimal_match = re.search(r'DECIMAL_SEPARATOR="([^"]*)"', content)
                    if decimal_match and decimal_match.group(1):
                        self.ui.comboBox_decimalseparator.setCurrentText(decimal_match.group(1))
                    
                    # Set currency symbol
                    currency_match = re.search(r'CURRENCY_SYMBOL="([^"]*)"', content)
                    if currency_match and currency_match.group(1):
                        self.ui.lineEdit_currency.setText(currency_match.group(1))
        except Exception as e:
            print(f"Error loading settings: {e}")

    def run_setlang(self):
        """Run the set language command"""
        subprocess.Popen([LINOFFICE_SCRIPT, 'manual', 'C:\\Program Files\\Microsoft Office\\root\\Office16\\SETLANG.EXE'])

    def save_settings(self):
        """Save all settings to config files"""
        try:
            import re
            registry_settings_changed = False
            
            # Save linoffice.conf settings
            linoffice_conf_path = os.path.join(os.path.dirname(LINOFFICE_SCRIPT), 'config', 'linoffice.conf')
            if os.path.exists(linoffice_conf_path):
                with open(linoffice_conf_path, 'r') as f:
                    content = f.read()
                
                # Update AUTOPAUSE setting
                autopause_value = "on" if self.ui.checkBox_suspend.isChecked() else "off"
                content = content.replace('AUTOPAUSE="on"', f'AUTOPAUSE="{autopause_value}"')
                content = content.replace('AUTOPAUSE="off"', f'AUTOPAUSE="{autopause_value}"')
                
                # Update RDP_SCALE setting
                scaling_text = self.ui.comboBox_scaling.currentText()
                if scaling_text.endswith('%'):
                    scaling_value = scaling_text[:-1]  # Remove % character
                else:
                    scaling_value = scaling_text
                
                content = content.replace('RDP_SCALE="100"', f'RDP_SCALE="{scaling_value}"')
                content = content.replace('RDP_SCALE="140"', f'RDP_SCALE="{scaling_value}"')
                content = content.replace('RDP_SCALE="180"', f'RDP_SCALE="{scaling_value}"')
                
                with open(linoffice_conf_path, 'w') as f:
                    f.write(content)

            # Save registry_override.conf settings
            registry_conf_path = os.path.join(os.path.dirname(LINOFFICE_SCRIPT), 'config', 'registry_override.conf')
            if os.path.exists(registry_conf_path):
                with open(registry_conf_path, 'r') as f:
                    original_content = f.read()
                
                content = original_content
                
                # Update DATE_FORMAT
                date_value = self.ui.comboBox_date.currentText()
                if date_value != "(no change)":
                    content = re.sub(r'DATE_FORMAT="[^"]*"', f'DATE_FORMAT="{date_value}"', content)
                
                # Update DECIMAL_SEPARATOR
                decimal_value = self.ui.comboBox_decimalseparator.currentText()
                if decimal_value != "(no change)":
                    content = re.sub(r'DECIMAL_SEPARATOR="[^"]*"', f'DECIMAL_SEPARATOR="{decimal_value}"', content)
                
                # Update CURRENCY_SYMBOL
                currency_value = self.ui.lineEdit_currency.text()
                content = re.sub(r'CURRENCY_SYMBOL="[^"]*"', f'CURRENCY_SYMBOL="{currency_value}"', content)
                
                # Check if registry settings actually changed
                if content != original_content:
                    registry_settings_changed = True
                
                with open(registry_conf_path, 'w') as f:
                    f.write(content)
            
            # Run registry_override.sh if registry settings were changed
            if registry_settings_changed:
                registry_script_path = os.path.join(os.path.dirname(LINOFFICE_SCRIPT), 'config', 'registry_override.sh')
                if os.path.exists(registry_script_path):
                    try:
                        subprocess.run([registry_script_path], check=True)
                    except subprocess.CalledProcessError as e:
                        QMessageBox.warning(self, 'Warning', f'Registry override script failed to run: {e}')
                    except FileNotFoundError:
                        QMessageBox.warning(self, 'Warning', 'Registry override script not found or not executable')
            
            self.settings_changed = False
            self.close()
            
        except Exception as e:
            QMessageBox.critical(self, 'Error', f'Failed to save settings: {e}')

    def closeEvent(self, event):
        """Handle window close event with confirmation dialog if settings changed"""
        if self.settings_changed:
            reply = QMessageBox.question(self, 'Save Changes?',
                                       'You have unsaved changes. Do you want to save them before closing?',
                                       QMessageBox.Save | QMessageBox.Discard | QMessageBox.Cancel,
                                       QMessageBox.Save)
            
            if reply == QMessageBox.Save:
                self.save_settings()
                event.accept()
            elif reply == QMessageBox.Discard:
                event.accept()
            else:  # Cancel
                event.ignore()
        else:
            event.accept()

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

    # Connect buttons in troubleshooting window with LinOffice script
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
