# This Python file uses the following encoding: utf-8
import sys
from PySide6.QtWidgets import QApplication, QWidget, QMainWindow, QMessageBox, QDialog, QVBoxLayout, QLabel, QPushButton
from PySide6.QtUiTools import QUiLoader
from PySide6.QtCore import QFile, QTimer
import subprocess
import os
import csv
import threading

LINOFFICE_SCRIPT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'linoffice.sh'))
SETUP_SCRIPT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'setup.sh'))
UNINSTALL_SCRIPT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'uninstall.sh'))

# Define the user's local registry override config path
USER_REGISTRY_CONFIG = os.path.expanduser('~/.local/share/linoffice/registry_override.conf')

# Define the languages CSV file path
LANGUAGES_CSV = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'config', 'languages.csv'))

# Define the internet state file path
INTERNET_STATE_FILE = os.path.expanduser('~/.local/share/linoffice/internet')

def ensure_internet_state_file():
        """Ensure the internet state file exists with default 'on' value"""
        state_dir = os.path.dirname(INTERNET_STATE_FILE)
        if not os.path.exists(state_dir):
            os.makedirs(state_dir, exist_ok=True)
        
        if not os.path.exists(INTERNET_STATE_FILE):
            # Create the file with default 'on' state
            with open(INTERNET_STATE_FILE, 'w') as f:
                f.write('on\n')

def load_internet_state():
    """Load the current internet state from file"""
    ensure_internet_state_file()
    try:
        with open(INTERNET_STATE_FILE, 'r') as f:
            state = f.read().strip()
            return state == 'on'
    except Exception as e:
        print(f"Error loading internet state: {e}")
        return True  # Default to 'on' if error

def save_internet_state(state_on):
    """Save the internet state to file"""
    ensure_internet_state_file()
    try:
        with open(INTERNET_STATE_FILE, 'w') as f:
            f.write('on\n' if state_on else 'off\n')
    except Exception as e:
        print(f"Error saving internet state: {e}")

def ensure_registry_config_exists():
    """Ensure the registry_override.conf file exists in user's local directory"""
    config_dir = os.path.dirname(USER_REGISTRY_CONFIG)
    if not os.path.exists(config_dir):
        os.makedirs(config_dir, exist_ok=True)
    
    if not os.path.exists(USER_REGISTRY_CONFIG):
        # Create the file with default empty values
        with open(USER_REGISTRY_CONFIG, 'w') as f:
            f.write('DATE_FORMAT=""\n')
            f.write('DECIMAL_SEPARATOR=""\n')
            f.write('CURRENCY_SYMBOL=""\n')

def load_languages_from_csv():
    """Load languages from the CSV file and return a list of formatted options"""
    languages = []
    try:
        with open(LANGUAGES_CSV, 'r', encoding='utf-8') as f:
            reader = csv.reader(f)
            next(reader)  # Skip header row
            for row in reader:
                if len(row) >= 3:
                    kbd_code, lang_code, lang_name = row[0], row[1], row[2]
                    # Format: "af_ZA [Afrikaans (South Africa)]"
                    formatted_option = f"{lang_code} [{lang_name}]"
                    languages.append((formatted_option, kbd_code))
    except Exception as e:
        print(f"Error loading languages from CSV: {e}")
    return languages

class MainWindow(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        # Ensure registry config exists before loading UI
        ensure_registry_config_exists()
        self.load_ui('main.ui')
        self.setWindowTitle(self.ui.windowTitle())
        self.connect_buttons()
        self.update_container_status()
        # Set up a timer to update container status every 30 seconds
        self.status_timer = QTimer(self)
        self.status_timer.timeout.connect(self.update_container_status)
        self.status_timer.start(30000)  # 30,000 ms = 30 seconds
        # Check container status and prompt user if not running
        self.check_and_prompt_container()

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

    def check_and_prompt_container(self):
        try:
            result = subprocess.run(['podman', 'ps', '--filter', 'name=LinOffice', '--format', '{{.Status}}'], capture_output=True, text=True, check=True)
            if not result.stdout.strip():
                # Container is not running, show dialog
                dialog = QMessageBox(self)
                dialog.setWindowTitle("Container Not Running")
                dialog.setText("The LinOffice container is not running. When you start a Windows app for the first time, it might take a bit longer as the container needs to be started first.")
                dialog.setStandardButtons(QMessageBox.Ok)
                response = dialog.exec()
        except subprocess.CalledProcessError as e:
            print(f"DEBUG: podman ps error: {e.stderr}")
            QMessageBox.critical(self, "Error", "Could not check LinOffice container status.")

# Defining secondary windows
class SettingsWindow(QMainWindow):
    def __init__(self, parent=None):
        super(SettingsWindow, self).__init__(parent)
        self.load_ui('settings.ui')
        self.setWindowTitle(self.ui.windowTitle())
        self.settings_changed = False
        self._initial_network_checked = None  # Track initial state
        self.load_current_settings()
        # Store the initial state after loading
        self._initial_network_checked = self.ui.checkBox_network.isChecked()
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
        self.ui.checkBox_network.toggled.connect(self.mark_settings_changed)  # Track network changes
        self.ui.comboBox_scaling.currentTextChanged.connect(self.mark_settings_changed)
        self.ui.comboBox_date.currentTextChanged.connect(self.mark_settings_changed)
        self.ui.comboBox_decimalseparator.currentTextChanged.connect(self.mark_settings_changed)
        self.ui.lineEdit_currency.textChanged.connect(self.mark_settings_changed)
        self.ui.comboBox_keyboard.currentTextChanged.connect(self.mark_settings_changed)

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
                    
                    # Load and populate keyboard comboBox
                    self.populate_keyboard_combo()
                    
                    # Set current keyboard selection
                    kbd_match = re.search(r'RDP_KBD="([^"]*)"', content)
                    if kbd_match and kbd_match.group(1):
                        current_kbd = kbd_match.group(1)
                        # Find the corresponding language option
                        for i in range(self.ui.comboBox_keyboard.count()):
                            item_data = self.ui.comboBox_keyboard.itemData(i)
                            if item_data == current_kbd:
                                self.ui.comboBox_keyboard.setCurrentIndex(i)
                                break

            # Load network state from file
            network_state = load_internet_state()
            self.ui.checkBox_network.setChecked(network_state)

            # Load registry_override.conf settings
            registry_conf_path = USER_REGISTRY_CONFIG
            # Ensure the file exists (recreate if deleted)
            ensure_registry_config_exists()
            
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

    def populate_keyboard_combo(self):
        """Populate the keyboard comboBox with language options from CSV"""
        # Clear existing items except the first "(no change)" option
        self.ui.comboBox_keyboard.clear()
        self.ui.comboBox_keyboard.addItem("(no change)", "")
        
        # Load languages from CSV
        languages = load_languages_from_csv()
        for formatted_option, kbd_code in languages:
            self.ui.comboBox_keyboard.addItem(formatted_option, kbd_code)

    def run_setlang(self):
        """Run the set language command"""
        subprocess.Popen([LINOFFICE_SCRIPT, 'manual', 'C:\\Program Files\\Microsoft Office\\root\\Office16\\SETLANG.EXE'])

    def save_settings(self):
        """Save all settings to config files"""
        try:
            import re
            registry_settings_changed = False
            # --- Network checkbox logic ---
            network_checked = self.ui.checkBox_network.isChecked()
            if self._initial_network_checked is not None and network_checked != self._initial_network_checked:
                if network_checked:
                    threading.Thread(target=lambda: subprocess.Popen([LINOFFICE_SCRIPT, 'internet_on'])).start()
                else:
                    threading.Thread(target=lambda: subprocess.Popen([LINOFFICE_SCRIPT, 'internet_off'])).start()
                # Save the new state to file
                save_internet_state(network_checked)
                # Update the initial state for next time
                self._initial_network_checked = network_checked
            # --- End network checkbox logic ---
            
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
                
                # Update RDP_KBD setting
                keyboard_data = self.ui.comboBox_keyboard.currentData()
                if keyboard_data:  # Only update if a keyboard is selected (not "(no change)")
                    kbd_string = f'/kbd:layout:{keyboard_data}'
                    if re.search(r'RDP_KBD="[^"]*"', content):
                        content = re.sub(r'RDP_KBD="[^"]*"', f'RDP_KBD="{kbd_string}"', content)
                    else:
                        # If RDP_KBD is missing, add it (optional, but not required by your spec)
                        content += f'\nRDP_KBD="{kbd_string}"\n'
                
                with open(linoffice_conf_path, 'w') as f:
                    f.write(content)

            # Save registry_override.conf settings
            registry_conf_path = USER_REGISTRY_CONFIG
            # Ensure the file exists before saving
            ensure_registry_config_exists()
            
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
            
            # Run linoffice.sh registry_override if registry settings were changed
            if registry_settings_changed:
                try:
                    subprocess.Popen([LINOFFICE_SCRIPT, 'registry_override'])
                except FileNotFoundError:
                    QMessageBox.warning(self, 'Warning', 'LinOffice script not found or not executable')
            
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

    def show_warning_dialog_rdp(self, action):
        dialog = QMessageBox(self)
        dialog.setWindowTitle("Warning")
        dialog.setText("Warning: Be careful when changing any Windows settings as it might break something.")
        dialog.setStandardButtons(QMessageBox.Ok)
        dialog.setIcon(QMessageBox.Information)
        
        # Execute the action only if OK is clicked
        if dialog.exec() == QMessageBox.Ok:
            action()

    def show_warning_dialog_vnc(self, action):
        dialog = QMessageBox(self)
        dialog.setWindowTitle("Warning")
        dialog.setText("Warning: Be careful when changing any Windows settings as it might break something.\n\nNote: The VNC connection does not have clipboard sharing or /home folder sharing.\n\nThe password to log in is 'MyWindowsPassword'. Please sign out at the end (Start > MyWindowsUser > Sign out) as RDP connections might be blocked otherwise.")
        dialog.setStandardButtons(QMessageBox.Ok)
        dialog.setIcon(QMessageBox.Information)
        
        # Execute the action only if OK is clicked
        if dialog.exec() == QMessageBox.Ok:
            action()

    # Connect buttons in tools window with LinOffice script
    def connect_tools_buttons(self):
        if self.main_window:
            self.ui.pushButton_update.clicked.connect(lambda: self.main_window.launch_linoffice_app('update'))
            self.ui.pushButton_powershell.clicked.connect(lambda: self.main_window.launch_linoffice_app('manual', 'powershell.exe'))
            self.ui.pushButton_regedit.clicked.connect(lambda: self.main_window.launch_linoffice_app('manual', 'regedit.exe'))
            self.ui.pushButton_cmd.clicked.connect(lambda: self.main_window.launch_linoffice_app('manual', 'cmd.exe'))
            self.ui.pushButton_explorer.clicked.connect(lambda: self.main_window.launch_linoffice_app('manual', 'explorer.exe'))
            self.ui.pushButton_access.clicked.connect(lambda: self.main_window.launch_linoffice_app('manual', 'msaccess.exe'))
            self.ui.pushButton_publisher.clicked.connect(lambda: self.main_window.launch_linoffice_app('manual', 'mspub.exe'))
            self.ui.pushButton_windows_rdp.clicked.connect(lambda: self.show_warning_dialog_rdp(lambda: self.main_window.launch_linoffice_app('windows')))
            self.ui.pushButton_windows_vnc.clicked.connect(lambda: self.show_warning_dialog_vnc(self.open_vnc_in_browser))

    def open_vnc_in_browser(self):
        import webbrowser
        webbrowser.open('http://127.0.0.1:8006')

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
