#!/bin/bash

# LinOffice Setup Script

CONTAINER_NAME="LinOffice" # should match the name in the compose.yaml
CONTAINER_EXISTS=0  # 0 = Does not exist (default), 1 = exists

# Absolute filepaths
USER_APPLICATIONS_DIR="${HOME}/.local/share/applications"
APPDATA_PATH="${HOME}/.local/share/linoffice"
# Ensure APPDATA_PATH exists before using it
mkdir -p "$APPDATA_PATH"
SUCCESS_FILE="${APPDATA_PATH}/success"
PROGRESS_FILE="${APPDATA_PATH}/setup_progress.log"

# Relative filepaths
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
LINOFFICE_DIR="$SCRIPT_DIR"
LINOFFICE="$(realpath "${SCRIPT_DIR}/linoffice.sh")"
COMPOSE_FILE="$(realpath "${SCRIPT_DIR}/config/compose.yaml")"
LINOFFICE_CONF="$(realpath "${SCRIPT_DIR}/config/linoffice.conf")"
OEM_DIR="$(realpath "${SCRIPT_DIR}/config/oem")"
LOCALE_REG_SCRIPT="$(realpath "${SCRIPT_DIR}/config/locale_reg.sh")"
LOCALE_LANG_SCRIPT="$(realpath "${SCRIPT_DIR}/config/locale_lang.sh")"
REGIONAL_REG="$(realpath "${SCRIPT_DIR}/config/oem/registry/regional_settings.reg")"
LOGFILE="${APPDATA_PATH}/windows_install.log"
APPS_DIR="$(realpath "${SCRIPT_DIR}/apps")"
DESKTOP_DIR="$(realpath "${APPS_DIR}/desktop")"
FREERDP_COMMAND="" # will be checked in the script whether it's xfreerdp, xfreerdp3, or the Flatpak version

# Progress tracking states
PROGRESS_REQUIREMENTS="requirements_completed"
PROGRESS_CONTAINER="container_created"
PROGRESS_OFFICE="office_installed"
PROGRESS_DESKTOP="desktop_files_installed"

# Command line arguments
DESKTOP_ONLY=false
FIRSTRUN=false
INSTALL_OFFICE_ONLY=false
HEALTHCHECK=false
INSTALL_EVKEY_ONLY=false

# Colors for output
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

USE_VENV=0

COMPOSE_COMMAND="podman-compose"

# Functions to print colored output
print_error() {
    echo -e "${RED}ERROR:${NC} $1" >&2
}

print_success() {
    echo -e "${GREEN}SUCCESS:${NC} $1"
}

print_info() {
    echo -e "${YELLOW}INFO:${NC} $1"
}

print_step() {
    echo -e "\n${GREEN}Step $1:${NC} $2"
}

print_progress() {
    echo -e "${GREEN}Progress:${NC} $1"
}

# Name: 'use_venv'
# Role: Activate virtual environment if available
use_venv() {
  local venv_dir="$HOME/.local/bin/linoffice/venv"
  local activate_script="$venv_dir/bin/activate"

  print_info "Checking for virtual environment at: $venv_dir"
  
  if [[ -f "$activate_script" ]]; then
    print_info "Virtual environment found at $venv_dir"
    source "$activate_script"
    VENV_PATH="$venv_dir"
    USE_VENV=1

    PYTHON_PATH="$venv_dir/bin/python3"
    print_info "Virtual environment Python: $PYTHON_PATH"

    USER_SITE_PATH=$($PYTHON_PATH -m site | grep USER_SITE | awk -F"'" '{print $2}')
    PODMAN_COMPOSE_BIN=$USER_SITE_PATH/podman_compose.py

    if [[ -f "$PODMAN_COMPOSE_BIN" ]]; then
        print_info "Using podman-compose from virtual environment: $PODMAN_COMPOSE_BIN"
        COMPOSE_COMMAND="$PYTHON_PATH $PODMAN_COMPOSE_BIN"
        return 0
    else
        print_info "podman-compose not found in virtual environment user packages at: $PODMAN_COMPOSE_BIN"
        print_info "Checking if virtual environment can access system podman-compose..."
        
        # Check if venv can access system podman-compose (due to --system-site-packages)
        if command -v podman-compose &> /dev/null; then
            print_info "Virtual environment can access system podman-compose"
            COMPOSE_COMMAND="podman-compose"
            return 0
        else
            print_info "podman-compose not available in virtual environment or system"
            print_info "Will check for system podman-compose instead"
            # Don't return here - let the system check handle it
            USE_VENV=0  # Reset to system mode since venv doesn't have podman-compose
            return 1
        fi
    fi
  else
    print_info "Virtual environment not found at $venv_dir, using system Python"
    return 1
  fi
}

use_venv

# Function to display usage information
print_usage() {
    print_info "Usage: $0 [--desktop] [--firstrun] [--installoffice] [--healthcheck]"
    print_info "Options:"
    print_info " (no flag)     Run the installation script from the beginning"
    print_info "  --desktop    Only recreate the desktop files (.desktop launchers)"
    print_info "  --firstrun   Force RDP and Office installation checks"
    print_info "  --installoffice   Only run the Office installation script script (in case the Windows installation has finished but Office is not installed)"
    print_info "  --healthcheck   Check that the system requirements are met and dependencies are installed and the container is healthy"
    exit 1
}

# Parse command line arguments
while [[ $# -gt 0 ]]; do
    case $1 in
        --desktop)
            DESKTOP_ONLY=true
            shift
            ;;
        --firstrun)
            FIRSTRUN=true
            shift
            ;;
        --installoffice)
            INSTALL_OFFICE_ONLY=true
            shift
            ;;
        --healthcheck)
            HEALTHCHECK=true
            shift
            ;;
        --installevkey)
            INSTALL_EVKEY_ONLY=true
            ;;            
        --help)
            print_usage
            ;;
        *)
            print_error "Unknown option: $1"
            print_usage
            ;;
    esac
done

# Function to exit with error
exit_with_error() {
    print_error "$1"
    exit 1
}

# Progress tracking functions
function init_progress_file() {
    mkdir -p "$(dirname "$PROGRESS_FILE")"
    touch "$PROGRESS_FILE"
}

function mark_progress() {
    local step=$1
    echo "$step" >> "$PROGRESS_FILE"
}

function check_progress() {
    local step=$1
    if [ -f "$PROGRESS_FILE" ] && grep -q "^$step$" "$PROGRESS_FILE"; then
        return 0
    else
        return 1
    fi
}

function clear_progress() {
    if [ -f "$PROGRESS_FILE" ]; then
        rm "$PROGRESS_FILE"
    fi
}

function check_linoffice_container() {
    print_info "Checking if LinOffice container exists already"
    if podman container exists "$CONTAINER_NAME"; then
        print_info "Container exists already."
        CONTAINER_EXISTS=1
    else
        print_info "Container does not yet exist."
        CONTAINER_EXISTS=0
    fi
}

# Function to detect and set the FreeRDP command
function detect_freerdp_command() {
    # Set FREERDP_COMMAND to the first available FreeRDP command (xfreerdp, xfreerdp3, or flatpak)
    # check xfreerdp3 before xfreerdp, some systems have both installed. e.g. SteamOS has outdated xfreerdp, but also has xfreerdp3 installed
    if command -v xfreerdp3 &>/dev/null; then
        FREERDP_COMMAND="xfreerdp3"
        return
    fi
    if command -v flatpak &>/dev/null; then
        if flatpak list --columns=application | grep -q "^com.freerdp.FreeRDP$"; then
            FREERDP_COMMAND="flatpak run --command=xfreerdp com.freerdp.FreeRDP"
            return
        fi
    fi
    if command -v xfreerdp &>/dev/null; then
        FREERDP_COMMAND="xfreerdp"
        return
    fi
    FREERDP_COMMAND="" # Not found
}

# Function to check if all requirements are met to run the Windows VM in Podman
function check_requirements() {

    # Exit on any error
    set -e
    print_info "Starting LinOffice setup script..."
    print_step "1" "Checking requirements"

    # Check minimum RAM (8 GB)
    print_info "Checking minimum RAM"
    REQUIRED_RAM=7 # 8 GB shows up as 7.6 GB so best to just set the threshold to 7 in this script
    AVAILABLE_RAM="$(free -b | awk '/^Mem:/{print int($2/1024/1024/1024)}')"
    if [ "$AVAILABLE_RAM" -lt "$REQUIRED_RAM" ]; then
        exit_with_error "Insufficient RAM. Required: ${REQUIRED_RAM}GB, Available: ${AVAILABLE_RAM}GB. \
    Please upgrade your system memory to at least ${REQUIRED_RAM}GB."
    fi
    print_success "Sufficient RAM detected: ${AVAILABLE_RAM}GB"

    # Check minimum free storage (64 GB)
    check_linoffice_container
    if [ "$CONTAINER_EXISTS" -eq 1 ]; then
        print_info "Container exists already. Skipping check if sufficient free storage is available."
    else        
        print_info "Checking minimum free storage"
        REQUIRED_STORAGE=64
        AVAILABLE_STORAGE=$(df -B1G --output=avail /home | tail -n 1 | awk '{print $1}')
        if [ "$AVAILABLE_STORAGE" -lt "$REQUIRED_STORAGE" ]; then
            exit_with_error "Insufficient free storage. Required: ${REQUIRED_STORAGE}GB, Available: ${AVAILABLE_STORAGE}GB \
        Please free up disk space or use a different storage device."
        fi
        print_success "Sufficient free storage detected: ${AVAILABLE_STORAGE}GB"
    fi

    # Check if computer supports virtualization
    print_info "Checking virtualization support"

    if ! command -v lscpu &> /dev/null; then
        exit_with_error "lscpu command not found. Please install util-linux package."
    fi

    # Check for virtualization support
    if lscpu | grep -qiE 'virtualization|vmx|svm'; then
        echo "Virtualization is supported."
    else
        exit_with_error "CPU virtualization not supported or not enabled.
        
        HOW TO FIX:
        1. Reboot your computer and enter BIOS/UEFI settings (usually F2, F12, Del, or Esc during boot)
        2. Look for virtualization settings:
        - Intel: Enable 'Intel VT-x' or 'Intel Virtualization Technology'
        - AMD: Enable 'AMD-V' or 'SVM Mode'
        3. Save settings and reboot
        4. If you can't find these options, your CPU may not support virtualization"
    fi

    # Additional check for KVM support
    if [ ! -e /dev/kvm ]; then
        exit_with_error "KVM device not available. Virtualization may not be enabled in BIOS.
        
    HOW TO FIX:
    1. Ensure virtualization is enabled in BIOS (see previous instructions)
    2. Install KVM kernel modules: sudo modprobe kvm
    3. For Intel CPUs: sudo modprobe kvm_intel
    4. For AMD CPUs: sudo modprobe kvm_amd
    5. Reboot if necessary"
    fi

    print_success "Virtualization support detected: $VIRT_SUPPORT"

    # Check if podman is installed
    print_info "Checking if podman is installed"

    if ! command -v podman &> /dev/null; then
        exit_with_error "podman is not installed.
        
    HOW TO FIX:
    Ubuntu/Debian: sudo apt update && sudo apt install podman
    Fedora/RHEL: sudo dnf install podman
    OpenSUSE: sudo zypper install podman    
    Arch Linux: sudo pacman -S podman
    openSUSE: sudo zypper install podman

    Or visit: https://podman.io/getting-started/installation"
    fi
    
    if ! podman info >/dev/null 2>&1; then
        exit_with_error "Podman is not configured correctly or you lack sufficient permissions. Run 'podman info' to diagnose the issue."
    fi

    PODMAN_VERSION=$(podman --version)
    print_success "podman is installed: $PODMAN_VERSION"

    # Check if podman-compose is installed
    print_info "Checking if podman-compose is installed"
    print_info "Python environment: $(if [[ "$USE_VENV" -eq 1 ]]; then echo "Virtual environment at $VENV_PATH"; else echo "System Python"; fi)"

    # Determine which Python to use for dependency checks
    if [[ "$USE_VENV" -eq 1 ]]; then
        PYTHON_CMD="$PYTHON_PATH"
        PYTHON_ENV="virtual environment"
    else
        PYTHON_CMD="python3"
        PYTHON_ENV="system"
    fi

    if [[ "$USE_VENV" -eq 0 ]]; then
        # Use system podman-compose, not the one in ~/.local/bin which might be broken
        if [[ -x "/usr/bin/podman-compose" ]]; then
            COMPOSE_COMMAND="/usr/bin/podman-compose"
            print_success "Using system podman-compose: /usr/bin/podman-compose"
        elif command -v podman-compose &> /dev/null; then
            COMPOSE_COMMAND="podman-compose"
            print_success "Using podman-compose from PATH: $(command -v podman-compose)"
        else
            exit_with_error "podman-compose is not installed.

        HOW TO FIX:
        Option 1 - Using pip: pip3 install podman-compose
        Option 2 - Using package manager:
        Ubuntu/Debian: sudo apt install podman-compose
        Fedora: sudo dnf install podman-compose
        OpenSUSE: sudo zypper install podman-compose
        Arch Linux: sudo pacman -S podman-compose

        Or visit: https://github.com/containers/podman-compose"
        fi
        # Check if python-dotenv is installed (dependency of podman-compose)
        if ! command -v $PYTHON_CMD &> /dev/null; then
            exit_with_error "$PYTHON_CMD command not found. Please install Python 3."
        fi
        
        # Show which Python is being used for debugging
        PYTHON_FULL_PATH=$(command -v $PYTHON_CMD)
        print_info "Using $PYTHON_ENV Python: $PYTHON_FULL_PATH"
        
        if $PYTHON_CMD -c "import dotenv" >/dev/null 2>&1; then
            print_success "python-dotenv is installed in $PYTHON_ENV environment."
        else
            # Provide more detailed error information
            print_error "python-dotenv is not installed or not accessible to $PYTHON_FULL_PATH"
            print_info "Python version: $($PYTHON_CMD --version 2>&1)"
            print_info "Python path: $PYTHON_FULL_PATH"
            print_info "Available packages: $($PYTHON_CMD -m pip list 2>/dev/null | grep -i dotenv || echo 'No dotenv packages found')"
            
            exit_with_error "python-dotenv is not installed in $PYTHON_ENV environment.

        HOW TO FIX:
        Using pip: $PYTHON_CMD -m pip install python-dotenv
        If you don't have pip, you can install it with your package manager.
        Ubuntu/Debian: sudo apt install python-dotenv
        Fedora: sudo dnf install python-dotenv
        OpenSUSE: sudo zypper install python-python-dotenv
        Arch Linux: sudo pacman -S python-dotenv
        
        If the package is installed but not detected, try:
        - Check if you have multiple Python versions: which python3
        - Install for the specific Python: $PYTHON_FULL_PATH -m pip install python-dotenv"

        fi
    else
        # When using virtual environment, check if podman-compose is installed in venv
        if [[ -f "$PODMAN_COMPOSE_BIN" ]]; then
            print_success "podman-compose is installed in virtual environment."
        else
            exit_with_error "podman-compose is not installed in virtual environment.

        HOW TO FIX:
        The virtual environment needs podman-compose installed.
        Run: $PYTHON_PATH -m pip install podman-compose"
        fi
        
        # Check if python-dotenv is installed in virtual environment (dependency of podman-compose)
        if $PYTHON_CMD -c "import dotenv" >/dev/null 2>&1; then
            print_success "python-dotenv is installed in virtual environment."
        else
            print_error "python-dotenv is not installed in virtual environment."
            print_info "Python version: $($PYTHON_CMD --version 2>&1)"
            print_info "Python path: $PYTHON_CMD"
            print_info "Available packages: $($PYTHON_CMD -m pip list 2>/dev/null | grep -i dotenv || echo 'No dotenv packages found')"
            print_info "User site packages: $($PYTHON_CMD -m site --user-site 2>/dev/null || echo 'Unknown')"
            
            # Check if it's available in user packages (installed with --user flag)
            USER_SITE_PACKAGES=$($PYTHON_CMD -m site --user-site 2>/dev/null)
            if [[ -n "$USER_SITE_PACKAGES" ]] && [[ -d "$USER_SITE_PACKAGES" ]] && find "$USER_SITE_PACKAGES" -name "*dotenv*" -type d 2>/dev/null | grep -q .; then
                print_info "python-dotenv found in user site packages but not importable"
                print_info "This might be due to virtual environment configuration issues"
            fi
            
            exit_with_error "python-dotenv is not installed in virtual environment.

        HOW TO FIX:
        The virtual environment needs python-dotenv installed.
        Run: $PYTHON_CMD -m pip install python-dotenv
        
        If packages were installed with --user flag by quickstart.sh, try:
        $PYTHON_CMD -m pip install --user python-dotenv"
        fi
    fi

    COMPOSE_VERSION=$($COMPOSE_COMMAND --version)
    print_success "podman-compose is installed: $COMPOSE_VERSION"

    # Check if FreeRDP is available
    print_info "Checking if FreeRDP is available"

    detect_freerdp_command
    local FREERDP_MAJOR_VERSION=""
    if [ -n "$FREERDP_COMMAND" ]; then
        if [ "$FREERDP_COMMAND" = "xfreerdp" ]; then
            FREERDP_MAJOR_VERSION=$(xfreerdp --version | head -n 1 | grep -o -m 1 '\b[0-9]\S*' | head -n 1 | cut -d'.' -f1)
        elif [ "$FREERDP_COMMAND" = "xfreerdp3" ]; then
            FREERDP_MAJOR_VERSION=$(xfreerdp3 --version | head -n 1 | grep -o -m 1 '\b[0-9]\S*' | head -n 1 | cut -d'.' -f1)
        elif [ "$FREERDP_COMMAND" = "flatpak run --command=xfreerdp com.freerdp.FreeRDP" ]; then
            FREERDP_MAJOR_VERSION=$(flatpak list --columns=application,version | grep "^com.freerdp.FreeRDP" | awk '{print $2}' | cut -d'.' -f1)
            # Check if Flatpak has access to /home
            if ! flatpak info --show-permissions com.freerdp.FreeRDP | grep -q "filesystems=.*home"; then
                exit_with_error "Flatpak FreeRDP does not have access to /home directory.
                
                HOW TO FIX:
                1. Close any running FreeRDP instances
                2. Run this command to grant access:
                   flatpak override --user --filesystem=home com.freerdp.FreeRDP
                3. Run this setup script again"
            fi
        fi
        if [[ ! $FREERDP_MAJOR_VERSION =~ ^[0-9]+$ ]] || ((FREERDP_MAJOR_VERSION < 3)); then
            exit_with_error "FreeRDP version 3 or greater is required. Detected version: $FREERDP_MAJOR_VERSION"
        fi
    else
        exit_with_error "FreeRDP is not installed
        
    HOW TO FIX:
    Option 1 - Using Flatpak and Flathub: flatpak install com.freerdp.FreeRDP
    Option 2 - Using package manager:
    Ubuntu/Debian: sudo apt install freerdp3-x11
    Fedora: sudo dnf install freerdp
    OpenSUSE: sudo zypper install freerdp
    Arch Linux: sudo pacman -S freerdp"
    fi

    if ! $FREERDP_COMMAND --version >/dev/null 2>&1; then
        exit_with_error "FreeRDP command '$FREERDP_COMMAND' is not functional. Please ensure FreeRDP is correctly installed and configured."
    fi

    print_success "FreeRDP found. Using FreeRDP command '${FREERDP_COMMAND}'."

    # Check if iptables modules are loaded
    print_info "Checking iptables kernel modules"
    if ! lsmod | grep -q ip_tables || ! lsmod | grep -q iptable_nat; then
        print_info "WARNING: iptables kernel modules not loaded. Sharing the /home folder with the Windows VM will not work unless connected via RDP. HOW TO FIX:
        
    Run the following command:
    echo -e 'ip_tables\niptable_nat' | sudo tee /etc/modules-load.d/iptables.conf
    Then reboot your system."
    fi
    print_success "iptables modules are loaded"

    # Check if most important LinOffice files exist
    print_info "Checking for essential setup files"

    if [ ! -d "$OEM_DIR" ]; then
        exit_with_error "OEM files not found
    Please ensure the config/oem directory exists"
    fi

    # Check OEM directory permissions
    if [ ! -r "$OEM_DIR" ] || [ ! -x "$OEM_DIR" ] || ! find "$OEM_DIR" -type f -readable | head -1 >/dev/null 2>&1; then
        exit_with_error "Insufficient permissions to access OEM directory: $OEM_DIR
        
        HOW TO FIX:
        1. Check directory permissions: ls -ld $OEM_DIR
        2. Fix permissions: chmod -R u+rwX $OEM_DIR
        3. If using SELinux/AppArmor, you may need to adjust security contexts"
    fi

    # Check if compose.yaml exists
    if [ ! -f "$COMPOSE_FILE.default" ]; then
        exit_with_error "Compose file not found: $COMPOSE_FILE.default
    Please ensure the file exists in the config directory."
    fi

        # Check if LinOffice script exists
    if [ ! -f "$LINOFFICE_CONF.default" ]; then
        exit_with_error "LinOffice configuration file not found: $LINOFFICE_CONF.default
    Please ensure the file exists in the config directory."
    fi
    
    if [ ! -f "$LINOFFICE" ]; then
        exit_with_error "File not found: $LINOFFICE"
    fi
    
    print_success "Files found."

    # Make scripts executable
    print_info "Making scripts executable"

    if [ ! -f "$LINOFFICE" ]; then
        exit_with_error "File not found: $LINOFFICE
    Please ensure the script is in the same directory as this setup script."
    fi

    if [ ! -f "$LOCALE_REG_SCRIPT" ]; then
        exit_with_error "File not found: $LOCALE_REG_SCRIPT
    Please ensure the config directory and locale_reg.sh script exist."
    fi

    if [ ! -f "$LOCALE_LANG_SCRIPT" ]; then
        exit_with_error "File not found: $LOCALE_LANG_SCRIPT
    Please ensure the config directory and local_compose.sh script exist."
    fi

    chmod +x "$LINOFFICE" || exit_with_error "Failed to make $LINOFFICE executable"
    chmod +x "$LOCALE_REG_SCRIPT" || exit_with_error "Failed to make $LOCALE_REG_SCRIPT executable"
    chmod +x "$LOCALE_LANG_SCRIPT" || exit_with_error "Failed to make $LOCALE_LANG_SCRIPT executable"

    print_success "Made scripts executable"

    # Check for various potential Podman problems
    # Check subUID/subGID mappings as some users had problems here
    print_info "Checking subUID/subGID mappings"
    if ! grep -q "^$(whoami):" /etc/subuid || ! grep -q "^$(whoami):" /etc/subgid; then
        exit_with_error "Missing subUID/subGID mappings for the user.
        HOW TO FIX:
        1. Run: sudo usermod --add-subuids 100000-165535 --add-subgids 100000-165535 $(whoami)
        2. Refresh Podman configuration: podman system migrate
        3. Verify mappings in /etc/subuid and /etc/subgid"
    fi
    print_success "subUID/subGID mappings verified."
    
    # Check Podman storage configuration and whether overlay storage driver is used, as some users had problems here
    print_info "Checking Podman storage configuration"
    if ! podman info --format '{{.Store.GraphDriverName}}' 2>/dev/null | grep -q "overlay"; then
        exit_with_error "Podman is not using overlay storage driver or there's a configuration issue.
        
        HOW TO FIX:
        1. Check Podman storage configuration: podman info
        2. Try resetting Podman: podman system reset (WARNING: This removes all containers and images)
        3. Check if /run/containers directory exists and is writable"
    fi
    
    # Determine if running rootless or rootful
    if podman info --format '{{.Host.Security.Rootless}}' | grep -q true; then
        IS_ROOTLESS=true
        STORAGE_DIR="$HOME/.local/share/containers/storage"
    else
        IS_ROOTLESS=false
        STORAGE_DIR="/var/lib/containers/storage"
    fi
    print_info "Podman running in $( $IS_ROOTLESS && echo 'rootless' || echo 'rootful' ) mode"

    # Set network directory paths based on rootless status first
    if [ "$IS_ROOTLESS" = true ]; then
        NETAVARK_DIR="$HOME/.local/share/containers/networks"
        CNI_DIR="$HOME/.config/cni/net.d"
    else
        NETAVARK_DIR="/var/lib/containers/networks"
        CNI_DIR="/etc/cni/net.d"
    fi

    # Now select the correct directory based on the network backend
    if [ "$NETWORK_BACKEND" = "netavark" ]; then
        NETWORK_DIR="$NETAVARK_DIR"
    else
        # Default to CNI backend
        NETWORK_DIR="$CNI_DIR"
    fi
    print_info "Podman network directory: $NETWORK_DIR"

    # Check if storage directory is accessible
    if [ ! -d "$STORAGE_DIR" ] || [ ! -w "$STORAGE_DIR" ]; then
        exit_with_error "Podman storage directory inaccessible: $STORAGE_DIR
        
        1. Check if directory exists: ls -ld \"$STORAGE_DIR\"
        2. If it exists, fix permissions: $( $IS_ROOTLESS && echo "chmod -R u+rwX \"$STORAGE_DIR\"" || echo "sudo chmod -R u+rwX \"$STORAGE_DIR\"" )
        3. If it does not exist, initialize Podman: podman info"
    fi
    print_success "Podman storage directory verified: $STORAGE_DIR"
    
    # Check which networking backend is in use
    print_info "Checking Podman networking is working"
    NETWORK_BACKEND=$(podman info --format '{{.Host.NetworkBackend}}' 2>/dev/null)
    if [ -z "$NETWORK_BACKEND" ]; then
        exit_with_error "Failed to detect Podman's network backend. Make sure Podman is correctly installed and accessible to your user. Run 'podman info' to diagnose."
    fi
    print_info "Podman is using network backend: $NETWORK_BACKEND"

    # Test network creation for all backends
    TEST_NET_NAME="linoffice_net_test_$(date +%s)"
    print_info "Testing network creation with backend: $NETWORK_BACKEND"
    if ! podman network create "$TEST_NET_NAME" >/dev/null 2>&1; then
        exit_with_error "Failed to create test network '$TEST_NET_NAME'.
        
        HOW TO FIX:
        1. Check Podman logs: journalctl -u podman
        2. $( $IS_ROOTLESS && echo 'Ensure user has sufficient permissions.' || echo 'Run as root or check sudo permissions.' )
        3. Reinstall network backend:
           - For netavark: $( $IS_ROOTLESS && echo 'podman system reset && podman info' || echo 'sudo dnf reinstall netavark || sudo apt install netavark' )
           - For CNI: Ensure CNI plugins are installed (e.g., sudo dnf install containernetworking-plugins)
        4. Verify SELinux/AppArmor settings if enabled."
    fi
    print_success "Test network '$TEST_NET_NAME' created successfully."

    # Check that network directory exists
    if [ ! -d "$NETWORK_DIR" ]; then
        print_info "WARNING: Network directory does not exist: $NETWORK_DIR
        This might lead to errors."
    fi
    if [ ! -w "$NETWORK_DIR" ]; then
        print_info "WARNING: Network directory not writable: $NETWORK_DIR
        This might lead to errors."
    fi
    
    # Clean up test network
    if podman network exists "$TEST_NET_NAME" >/dev/null 2>&1; then
        podman network rm "$TEST_NET_NAME" >/dev/null 2>&1 || print_info "Note: Failed to remove test network '$TEST_NET_NAME', you may remove it manually."
    fi
    print_success "Podman networking check completed."

    # Test basic container creation to catch storage issues early
    print_info "Testing basic container functionality..."
    if ! timeout 60 podman run --rm alpine:latest echo "test" >/dev/null 2>&1; then
        exit_with_error "Basic container test failed. This could indicate storage driver issues.
        
        HOW TO FIX:
        1. Check Podman logs: journalctl --user -u podman
        2. Try: podman system reset (WARNING: removes all containers/images)
        3. Ensure /run/containers and storage directories have correct permissions
        4. Check if your filesystem supports overlay mounts"
    fi
    print_success "Podman test container created and removed successfully."

    # Check connectivity to microsoft.com
    print_info "Checking connectivity to Microsoft"

    if ! curl -s --head --request GET --max-time 10 -L https://www.microsoft.com | grep -q "200"; then
        # Alternative method: curl to a reliable fallback endpoint
        if ! curl -s --head --request GET --max-time 10 -L https://www.office.com | grep -q "200"; then
            exit_with_error "Unable to connect to microsoft.com.
            HOW TO FIX:
            1. Check your internet connection
            2. Verify DNS settings: Ensure you can resolve microsoft.com (try: nslookup microsoft.com)
            3. Check firewall settings: Ensure outbound connections to microsoft.com are allowed
            4. Try again or contact your network administrator"
        else
            print_success "Successfully connected to Microsoft"
        fi
    else
        print_success "Successfully connected to Microsoft"
    fi

    # Run locale scripts
    print_step "2" "Detecting region and language settings"
    print_info "Running locale configuration scripts"

    print_info "Executing: $LOCALE_REG_SCRIPT"
    if ! "$LOCALE_REG_SCRIPT"; then
        exit_with_error "Failed to execute $LOCALE_REG_SCRIPT (exit code: $?)"
    fi

    print_info "Executing: $LOCALE_LANG_SCRIPT"
    if ! "$LOCALE_LANG_SCRIPT"; then
        exit_with_error "Failed to execute $LOCALE_LANG_SCRIPT (exit code: $?)"
    fi

    print_success "Locale script executed successfully"

    # Check if newly created regional.reg exists
    print_info "Checking for regional_settings.reg file"

    if [ ! -f "$REGIONAL_REG" ]; then
        exit_with_error "Required file not found: $REGIONAL_REG
    Please ensure the config/oem/registry directory exists and contains regional_settings.reg"
    fi

    print_success "Found regional_settings.reg file"
}

function setup_logfile() {
    # Check if the logfile already exists, if yes rename old one with its last modified date and start with a fresh logfile
    mkdir -p "$(dirname "$LOGFILE")"
    echo "Logfile: $LOGFILE"
    if [ -e "$LOGFILE" ]; then
        MODIFIED_DATE=$(stat -c %y "$LOGFILE" | sed 's/[: ]/_/g' | cut -d '.' -f 1)
        mv "$LOGFILE" "${LOGFILE%.log}_$MODIFIED_DATE.log"
    fi
}

function create_container() {
    print_step "3" "Setting up the LinOffice container"
    local bootcount=0
    local required_boots=4  # Accept 4 as the minimum, but allow for 5 if it happens (see comment below)
    # The Windows/Office install process may show 4 or 5 reboots depending on version and script details.
    # Sometimes the reboot between install.bat and InstallOffice.ps1 is not a full UEFI reboot, so only 4 are seen.
    # We proceed if we see at least 4 reboots, and print a note if more are detected.
        # this is how many times the Windows VM needs to boot to be ready
        # the string to look for is "BdsDxe: starting Boot0004"
        # 3 reboots will be logged during initial Windows until you can see the desktop for the first time
        # 1 reboot at the end of install.bat (this is the one that is not always logged for some reason)
        # 1 reboot at the end of the InstallOffice.ps1
    local result=1  # 0 = success, 1 = failure (assume failure by default)
    local download_started=false
    local download_finished=false
    local install_started=false
    local timeout_counter=0
    local max_timeout=3600  # 60 minutes maximum wait time between podman-compose log output
    local last_activity_time=$(date +%s)
    local windows_version=""

    # Start podman-compose in the background with unbuffered output and strip ANSI codes
    print_info "Starting podman-compose in detached mode..."
	# If the compose file doesn't exist yet, initialize it from the default template
	if [ ! -f "$COMPOSE_FILE" ]; then
		if [ -f "$COMPOSE_FILE.default" ]; then
			print_info "Creating $COMPOSE_FILE from default template"
			cp "$COMPOSE_FILE.default" "$COMPOSE_FILE" || exit_with_error "Failed to copy $COMPOSE_FILE.default to $COMPOSE_FILE"
		else
			exit_with_error "Compose file missing: $COMPOSE_FILE and $COMPOSE_FILE.default not found"
		fi
	fi
    if ! $COMPOSE_COMMAND --file "$COMPOSE_FILE" up -d >>"$LOGFILE" 2>&1; then
        exit_with_error "Failed to start containers. Check $LOGFILE for details."
    fi

    # Check if container was actually created
    sleep 5
    if ! podman ps -a --filter "name=$CONTAINER_NAME" --format "{{.Names}}" | grep -q "^$CONTAINER_NAME$"; then
        exit_with_error "Container $CONTAINER_NAME was not created successfully. Check $LOGFILE for detailed error messages."
    fi

    print_info "Tailing logs from container: $CONTAINER_NAME"
    podman logs -f --timestamps "$CONTAINER_NAME" 2>&1 | \
        stdbuf -oL -eL sed -u 's/\x1b\[[0-9;]*m//g' >> "$LOGFILE" &
    log_pid=$!

    print_info "Monitoring container setup progress..."
    
    # Monitor the logfile for progress
    while true; do
        local current_time=$(date +%s)
        if [ $((current_time - last_activity_time)) -gt $max_timeout ]; then
            result=1
            exit_with_error "Container setup timed out after $((max_timeout/60)) minutes."
        fi
        
        # Read the logfile if it exists
        if [ -f "$LOGFILE" ]; then
            # Get file size to detect new activity
            local current_size=$(stat -c%s "$LOGFILE" 2>/dev/null || echo "0")
            local previous_size=${previous_size:-0}
            
            if [ "$current_size" -gt "$previous_size" ]; then
                last_activity_time=$current_time
                previous_size=$current_size
            fi

            # Check for download progress
            if ! $download_started && grep -q "Downloading Windows" "$LOGFILE"; then
                print_step "4" "Starting Windows download (about 5 GB). This will take a while depending on your Internet speed."
                download_started=true
                last_activity_time=$current_time
                windows_version=$(grep "Downloading Windows" "$LOGFILE" | tail -1 | grep -oE '10|11')
            fi

            # Output download progress at each percent
            if $download_started && ! $download_finished; then
                # Parse wget progress line: percent and speed
                last_percent=${last_percent:--1}
                progress_line=$(grep -E "%" "$LOGFILE" | grep -E "[0-9.]+[MK]" | tail -1)
                pct=$(echo "$progress_line" | grep -oE "[ ]{1,3}[0-9]{1,3}%" | tail -1 | tr -d ' %')
                speed=$(echo "$progress_line" | grep -oE "[0-9.]+[MK]" | tail -1)
                if [[ "$pct" =~ ^[0-9]+$ ]] && [ "$pct" -gt "$last_percent" ] && [ "$pct" -le 100 ]; then
                    if [ -n "$windows_version" ]; then
                        print_progress "Downloading Windows ${windows_version}: ${pct}% | Speed: ${speed}B/s"
                    else
                        print_progress "Downloading Windows: ${pct}% | Speed: ${speed}B/s"
                    fi
                    last_percent=$pct
                fi
            fi

            # Check for download completion
            if $download_started && ! $download_finished && grep -q "100%" "$LOGFILE"; then
                print_success "Windows download finished"
                download_finished=true
                last_activity_time=$current_time

                # Monitor for either "Windows started" or "Shutdown completed" after download
                print_info "Waiting for Windows to start after download..."
                local monitor_timeout=300  # 5 minutes
                local monitor_elapsed=0
                local monitor_interval=5
                while [ $monitor_elapsed -lt $monitor_timeout ]; do
                    if grep -q "Windows started" "$LOGFILE"; then
                        print_step "5" "Installing Windows. This will take a while."
                        install_started=true
                        last_activity_time=$(date +%s)
                        break
                    fi
                    if grep -q "Shutdown completed" "$LOGFILE"; then
                        exit_with_error "Windows installation failed: Detected shutdown before Windows started. Check $LOGFILE for details and see https://github.com/dockur/windows/issues for troubleshooting."
                    fi
                    sleep $monitor_interval
                    monitor_elapsed=$((monitor_elapsed + monitor_interval))
                done
                if [ $monitor_elapsed -ge $monitor_timeout ]; then
                    exit_with_error "Timeout waiting for Windows to start after download. Check $LOGFILE for details."
                fi
            fi

            # Check for error conditions
            if grep -iq "error\|failed\|cannot\|timeout" "$LOGFILE" | tail -10 | grep -q "FATAL\|ERROR"; then
                print_error "Error detected in container logs. Check $LOGFILE for details."
                # Don't exit immediately, but log the concern
            fi

            # Check for boot progress
            local current_boots=0
            current_boots=$(grep -c "BdsDxe: starting Boot0004" "$LOGFILE" 2>/dev/null) || current_boots=0
            if [ "$current_boots" -gt "$bootcount" ]; then
                bootcount=$current_boots
                print_success "Reboot $bootcount of $required_boots completed"
                if [ "$bootcount" -eq 3 ]; then
                    print_success "Windows installation finished"
                    print_step "6" "Downloading and installing Office (about 3 GB). This will take a while."
                fi
                if [ "$bootcount" -gt 4 ]; then
                    print_info "More than 4 reboots detected ($bootcount). This is expected in some cases. Proceeding."
                fi
                last_activity_time=$current_time
                if [ "$bootcount" -ge "$required_boots" ]; then
                    result=0
                    break
                fi
            fi
        fi

        # Sleep briefly to avoid high CPU usage
        sleep 5
    done

    # Stop the background log tailing process
    if kill -0 "$log_pid" 2>/dev/null; then
        kill "$log_pid" 2>/dev/null || true
    fi

    # Then check success/failure
    if [ "$result" -eq 0 ]; then
        sleep 5
        if ! podman ps -q --filter "name=$CONTAINER_NAME" | grep -q .; then
            exit_with_error "Container setup completed but container is not running. Check $LOGFILE for details."
        else
            print_success "Container setup completed successfully"
            return 0
        fi
    else
        exit_with_error "Container setup failed. Check $LOGFILE for details or visit 127.0.0.1:8006 in your web browser."
    fi
}

function verify_container_health() {
    print_info "Verifying container health..."
    
    # Ensure container exists, otherwise create it
    if ! podman container exists "$CONTAINER_NAME" 2>/dev/null; then
        print_info "Container does not exist. Creating it now with podman-compose up -d..."
        if ! $COMPOSE_COMMAND --file "$COMPOSE_FILE" up -d; then
            print_error "Failed to create container via compose up -d"
            return 1
        fi
        print_info "Waiting for container to boot..."
        sleep 20
    fi

    # Check if container is running, otherwise start it
    if ! podman ps -q --filter "name=$CONTAINER_NAME" | grep -q .; then
        print_info "Container is not running. Attempting to start it..."
        if ! $COMPOSE_COMMAND --file "$COMPOSE_FILE" start; then
            print_error "Failed to start container"
            print_info "Container may be in an improper state. Try these commands to fix it:
            1. podman rm -f LinOffice
            2. $COMPOSE_COMMAND --file config/compose.yaml up -d"
            return 1
        fi
        print_info "Waiting for container to boot..."
        sleep 20
    fi
    
    # Check container logs for any obvious errors
    local container_logs=$(podman logs --tail 50 "$CONTAINER_NAME" 2>/dev/null || echo "")
    if echo "$container_logs" | grep -iq "error\|failed\|fatal"; then
        print_error "Container logs show potential issues"
        print_info "If the container is in an improper state, try these commands to fix it:
        1. podman rm -f LinOffice
        2. $COMPOSE_COMMAND --file config/compose.yaml up -d"
        return 1
    fi
    
    return 0
}

function check_available() {
    if [ -z "$FREERDP_COMMAND" ]; then
        detect_freerdp_command
    fi
    print_step "7" "Checking if everything is set up correctly"
    print_info "Checking if RDP server is available"
    local max_attempts=10 
    local reboot_threshold=7
    local attempt=0
    local success=0
    local vm_rebooted=false
    
    if [ ! -e "$SUCCESS_FILE" ]; then
        while true; do
            attempt=0
            success=0
            vm_rebooted=false
            
            while [ $attempt -lt $max_attempts ]; do
            attempt=$((attempt + 1))

            # First verify container is healthy
            if ! verify_container_health; then
                print_error "Container health check failed on attempt $attempt"
                if [ $attempt -ge $max_attempts ]; then
                    break
                fi
                print_info "Waiting 10 seconds before next attempt..."
                sleep 10
                continue
            fi

            # Try to check if RDP is ready
            print_info "Testing RDP connection (attempt $attempt of $max_attempts)..."
            
            # First, try the original minimal command (as it works on many systems) before applying variants
            echo "DEBUG: Using FreeRDP command (minimal): $FREERDP_COMMAND" >> "$LOGFILE"
            local minimal_output
            if [[ "$FREERDP_COMMAND" == flatpak* ]]; then
                minimal_output=$(timeout 30 bash -c "$FREERDP_COMMAND /cert:ignore /u:MyWindowsUser /p:MyWindowsPassword /v:127.0.0.1 /port:3388 /app:program:cmd.exe,cmd:'/c tsdiscon'" 2>&1)
            else
                minimal_output=$(timeout 30 "$FREERDP_COMMAND" \
                    /cert:ignore \
                    /u:MyWindowsUser \
                    /p:MyWindowsPassword \
                    /v:127.0.0.1 \
                    /port:3388 \
                    /app:program:cmd.exe,cmd:'/c tsdiscon' \
                    2>&1)
            fi
            local minimal_exit=$?
            echo "DEBUG: FreeRDP (minimal) output was:" >> "$LOGFILE"
            echo "$minimal_output" >> "$LOGFILE"
            echo "DEBUG: FreeRDP (minimal) exit code was: $minimal_exit" >> "$LOGFILE"

            if echo "$minimal_output" | grep -q "ERRINFO_LOGOFF_BY_USER"; then
                print_success "RDP server is available (user logoff detected)"
                success=1
                break
            fi

            # Minimal probe: try all three binaries in fixed order regardless of detected default
            local candidates_minimal=("xfreerdp3" "flatpak run --command=xfreerdp com.freerdp.FreeRDP" "xfreerdp")

            # Function to run a single probe variant (conservative flags) against a specific candidate
            run_probe_variant() {
                local candidate="$1"
                local use_xvfb="$2" # "yes" or "no"
                local display_backup="$DISPLAY"
                local xvfb_pid=""
                local env_prefix="XLIB_SKIP_ARGB_VISUALS=1 LIBGL_ALWAYS_SOFTWARE=1 GDK_BACKEND=x11 QT_QPA_PLATFORM=xcb"
                local base_args="/cert:ignore /sec:rdp /gdi:sw /bpp:16 /u:MyWindowsUser /p:MyWindowsPassword /v:127.0.0.1 /port:3388 /app:program:cmd.exe,cmd:'/c tsdiscon'"

                # Optional headless X fallback via Xvfb
                if [ "$use_xvfb" = "yes" ]; then
                    if command -v Xvfb >/dev/null 2>&1; then
                        print_info "Using headless X fallback (Xvfb) for probe"
                        Xvfb :99 -screen 0 1024x768x16 >/dev/null 2>&1 &
                        xvfb_pid=$!
                        export DISPLAY=:99
                    else
                        echo "DEBUG: Xvfb not available, skipping headless probe" >> "$LOGFILE"
                    fi
                fi

                local output=""
                local exitcode=0
                echo "DEBUG: Using FreeRDP candidate: $candidate (xvfb=$use_xvfb)" >> "$LOGFILE"
                if [[ "$candidate" == flatpak* ]]; then
                    # Run via flatpak with env prefix
                    output=$(timeout 30 bash -c "$env_prefix $candidate $base_args" 2>&1)
                    exitcode=$?
                elif [ "$candidate" = "xfreerdp3" ] || [ "$candidate" = "xfreerdp" ]; then
                    output=$(timeout 30 env $env_prefix $candidate $base_args 2>&1)
                    exitcode=$?
                else
                    output="Unsupported candidate: $candidate"
                    exitcode=1
                fi

                # Restore DISPLAY and cleanup Xvfb
                if [ -n "$xvfb_pid" ]; then
                    kill "$xvfb_pid" >/dev/null 2>&1 || true
                    wait "$xvfb_pid" 2>/dev/null || true
                    export DISPLAY="$display_backup"
                fi

                echo "DEBUG: FreeRDP output was:" >> "$LOGFILE"
                echo "$output" >> "$LOGFILE"
                echo "DEBUG: FreeRDP exit code was: $exitcode" >> "$LOGFILE"

                # Return via globals
                FREERDP_LAST_OUTPUT="$output"
                FREERDP_LAST_EXIT="$exitcode"
            }

            # Try minimal probe with all three candidates
            if [ $success -ne 1 ]; then
                for candidate in "${candidates_minimal[@]}"; do
                    # Skip if candidate is not available on system
                    if [[ "$candidate" == flatpak* ]]; then
                        if ! command -v flatpak >/dev/null 2>&1 || ! flatpak list --columns=application | grep -q "^com.freerdp.FreeRDP$"; then
                            continue
                        fi
                    else
                        if ! command -v "$candidate" >/dev/null 2>&1; then
                            continue
                        fi
                    fi
                    echo "DEBUG: Trying minimal probe with candidate: $candidate" >> "$LOGFILE"
                    local cand_output
                    if [[ "$candidate" == flatpak* ]]; then
                        cand_output=$(timeout 30 bash -c "$candidate /cert:ignore /u:MyWindowsUser /p:MyWindowsPassword /v:127.0.0.1 /port:3388 /app:program:cmd.exe,cmd:'/c tsdiscon'" 2>&1)
                    else
                        cand_output=$(timeout 30 $candidate \
                            /cert:ignore \
                            /u:MyWindowsUser \
                            /p:MyWindowsPassword \
                            /v:127.0.0.1 \
                            /port:3388 \
                            /app:program:cmd.exe,cmd:'/c tsdiscon' \
                            2>&1)
                    fi
                    local cand_exit=$?
                    echo "DEBUG: FreeRDP (minimal candidate) output was:" >> "$LOGFILE"
                    echo "$cand_output" >> "$LOGFILE"
                    echo "DEBUG: FreeRDP (minimal candidate) exit code was: $cand_exit" >> "$LOGFILE"
                    if echo "$cand_output" | grep -q "ERRINFO_LOGOFF_BY_USER"; then
                        print_success "RDP server is available (user logoff detected)"
                        success=1
                        break
                    fi
                done
            fi

            # If minimal did not succeed, try conservative variant with the originally detected command only
            if [ $success -ne 1 ]; then
                run_probe_variant "$FREERDP_COMMAND" "no"
                local freerdp_output="$FREERDP_LAST_OUTPUT"
                if echo "$freerdp_output" | grep -q "ERRINFO_LOGOFF_BY_USER"; then
                    print_success "RDP server is available (user logoff detected)"
                    success=1
                fi
            fi

            # If still not successful, do a single headless attempt with the originally detected command
            if [ $success -ne 1 ]; then
                run_probe_variant "$FREERDP_COMMAND" "yes"
                local freerdp_output="$FREERDP_LAST_OUTPUT"
                if echo "$freerdp_output" | grep -q "ERRINFO_LOGOFF_BY_USER"; then
                    print_success "RDP server is available (headless probe)"
                    success=1
                fi
            fi

            # If success was detected above, break out of attempt loop
            if [ $success -eq 1 ]; then
                break
            fi

            # If we've reached the reboot threshold and haven't rebooted yet, reboot the VM
            if [ $attempt -eq $reboot_threshold ] && [ "$vm_rebooted" = false ]; then
                print_info "Reached $reboot_threshold failed attempts. Rebooting Windows VM to restart Office installation..."
                vm_rebooted=true
                
                # Reboot the Windows VM using the existing reset functionality
                print_info "Rebooting Windows VM..."
                "$COMPOSE_COMMAND" --file "$COMPOSE_FILE" restart >>"$LOGFILE" 2>&1
                
                # Wait for container to restart
                local max_wait_time=120
                local wait_elapsed=0
                local check_interval=5
                
                print_info "Waiting for Windows VM to restart..."
                while [ $wait_elapsed -lt $max_wait_time ]; do
                    if timeout 1 bash -c ">/dev/tcp/127.0.0.1/3388" 2>/dev/null; then
                        print_success "Windows VM restarted successfully"
                        break
                    fi
                    sleep $check_interval
                    wait_elapsed=$((wait_elapsed + check_interval))
                    if [ $((wait_elapsed % 30)) -eq 0 ]; then
                        print_info "Still waiting for Windows VM to restart... ($((wait_elapsed/60)) minutes elapsed)"
                    fi
                done
                
                if [ $wait_elapsed -ge $max_wait_time ]; then
                    print_error "Timeout waiting for Windows VM to restart. Please check the container status."
                    return 1
                fi
                
                # Wait a bit more for Windows to fully boot
                print_info "Waiting for Windows to fully boot..."
                sleep 30
                
                # Now run the Office installation command
                print_info "Restarting Office installation after VM reboot..."
                echo "DEBUG: Running Office installation after VM reboot" >> "$LOGFILE"
                
                "$FREERDP_COMMAND" \
                    /cert:ignore \
                    +home-drive \
                    /u:MyWindowsUser \
                    /p:MyWindowsPassword \
                    /v:127.0.0.1 \
                    /port:3388 \
                    /app:program:powershell.exe,cmd:'-ExecutionPolicy Bypass -File C:\\OEM\\InstallOffice.ps1' \
                    >>"$LOGFILE" 2>&1 &
                
                print_info "Office installation command sent. Continuing with RDP connection attempts..."
            fi

            # If unable to connect, try again
            if [ $attempt -lt $max_attempts ]; then
                print_info "RDP server not ready yet, waiting 5 seconds..."
                sleep 5
            fi
            
            if [ $success -eq 1 ]; then
                print_success "RDP server is available"
                return 0
            fi

            # At this point we failed after max_attempts. Provide beginner-friendly guidance and offer retry.
            print_error "Failed to connect to the Windows VM via RDP after $max_attempts attempts."
            echo
            print_info "What to do now:"
            print_info "1) Open your web browser and go to: 127.0.0.1:8006"
            print_info "2) You will see the virtual machine. Log in using the password: MyWindowsPassword"
            print_info "3) In Windows, click Start (Windows logo), then click on 'MyWindowsUser' at the bottom-left, then click 'Sign out'."
            print_info "   Important: Do NOT shut down or restart the virtual machine. Just sign out."
            echo
            print_info "After signing out, we can try the connection again."
            echo
            # Machine-readable marker so the GUI can show a dialog and answer the prompt
            echo "PROMPT:VNC_SIGN_OUT_AND_RETRY"
            # Interactive prompt for terminal users
            local answer
            read -r -p "Try again now? [Y/n]: " answer
            answer=${answer:-Y}
            if [[ "$answer" =~ ^[Yy]$ ]]; then
                print_info "Okay, trying again..."
                continue
            else
                print_error "User chose not to retry RDP connection."
                return 1
            fi
        done
        done
    else
        print_success "Success file already exists"
        return 0
    fi
}

function check_success() {
    if [ -z "$FREERDP_COMMAND" ]; then
        detect_freerdp_command
    fi
    print_info "Checking if Office is installed"

    local freerdp_pid=""
    local elapsed_time=0
    local retry_count=0
    local max_retries=10
    # Keep connection alive for the whole installation monitoring period; do not kill early
    # local connection_timeout=60
    local check_interval=10  # Try again after 10 seconds
    local installation_timeout=1800 # 30 minutes timeout for Office download and installation
    
    # Function to cleanup FreeRDP process
    cleanup_freerdp() {
        if [ -n "$freerdp_pid" ] && kill -0 "$freerdp_pid" 2>/dev/null; then
            print_info "Cleaning up FreeRDP process (PID: $freerdp_pid)"
            kill -TERM "$freerdp_pid" 2>/dev/null || true
            sleep 3
            kill -KILL "$freerdp_pid" 2>/dev/null || true
        fi
    }

    # Register cleanup function to run on script exit
    trap cleanup_freerdp EXIT

    # Clear any existing success file once before attempting connections
    rm -f "$SUCCESS_FILE"

    # Retry loop for FreeRDP connection
    while [ $retry_count -lt $max_retries ]; do
        retry_count=$((retry_count + 1))
        print_info "Starting FreeRDP connection to mount home directory (Attempt $retry_count of $max_retries)..."
        
        # Start FreeRDP in the background with home-drive enabled
        if [[ "$FREERDP_COMMAND" == flatpak* ]]; then
            bash -c "$FREERDP_COMMAND /cert:ignore +home-drive /u:MyWindowsUser /p:MyWindowsPassword /v:127.0.0.1 /port:3388 /app:program:powershell.exe,cmd:'-ExecutionPolicy Bypass -File C:\\OEM\\FirstRDPRun.ps1'" >>"$LOGFILE" 2>&1 &
        else
            "$FREERDP_COMMAND" \
                /cert:ignore \
                +home-drive \
                /u:MyWindowsUser \
                /p:MyWindowsPassword \
                /v:127.0.0.1 \
                /port:3388 \
                /app:program:powershell.exe,cmd:'-ExecutionPolicy Bypass -File C:\\OEM\\FirstRDPRun.ps1' \
                >>"$LOGFILE" 2>&1 &
        fi
        
        freerdp_pid=$!
        
        # Wait briefly and check if FreeRDP started successfully
        sleep 5
        if kill -0 "$freerdp_pid" 2>/dev/null; then
            print_success "FreeRDP connection established successfully (PID: $freerdp_pid)"
            break
        else
            wait $freerdp_pid 2>/dev/null
            local exit_code=$?
            print_error "FreeRDP failed to start or exited immediately (exit code: $exit_code)"
            
            if [ $retry_count -lt $max_retries ]; then
                print_info "Retrying in 10 seconds..."
                sleep 10
            else
                print_error "Max retries ($max_retries) reached. Check log file at $LOGFILE for details."
                return 1
            fi
        fi
    done

    # Reset elapsed time for installation monitoring
    elapsed_time=0
    local last_check_time=$(date +%s)
    
    print_info "Waiting for Office installation to complete (timeout: $((installation_timeout/60)) minutes)..."
    
    # Monitor for success file creation
    while [ $elapsed_time -lt $installation_timeout ]; do
        # Check if success file exists
        if [ -f "$SUCCESS_FILE" ]; then
            print_success "Success file detected - Office installation is complete!"
            cleanup_freerdp
            return 0
        fi

        # Check if FreeRDP process is still running
        if ! kill -0 "$freerdp_pid" 2>/dev/null; then
            wait $freerdp_pid 2>/dev/null
            local exit_code=$?
            
            # Check if success file was created before process ended
            if [ -f "$SUCCESS_FILE" ]; then
                print_success "Success file detected - Office installation is complete!"
                return 0
            fi
            
            print_error "FreeRDP connection terminated (exit code: $exit_code)"
            print_info "Checking if success file was created..."
            
            sleep 2
            if [ -f "$SUCCESS_FILE" ]; then
                print_success "Success file found - Office installation completed successfully!"
                return 0
            else
                print_error "Success file not found. Installation may have failed."
                print_info "Check log file at $LOGFILE for details"
                return 1
            fi
        fi

        sleep $check_interval
        elapsed_time=$((elapsed_time + check_interval))
    done

    # Timeout reached
    print_error "Timeout waiting for Office installation to complete after $((installation_timeout / 60)) minutes"
    print_info "Check log file at $LOGFILE for details"
    
    # Final check for success file
    if [ -f "$SUCCESS_FILE" ]; then
        print_success "Success file found during cleanup - Office installation completed!"
        cleanup_freerdp
        return 0
    fi
    
    cleanup_freerdp
    return 1
}

function desktop_files() {
    print_step "8" "Installing .desktop files (app launchers)"
    
    # Check if required directories exist
    if [ ! -d "$DESKTOP_DIR" ]; then
        exit_with_error "Error: Desktop directory not found: $DESKTOP_DIR"
    fi

    if [ ! -d "$USER_APPLICATIONS_DIR" ]; then
        mkdir -p "$USER_APPLICATIONS_DIR" || exit_with_error "Failed to create $USER_APPLICATIONS_DIR"
    fi
    if [ ! -w "$USER_APPLICATIONS_DIR" ]; then
        exit_with_error "No write permissions for $USER_APPLICATIONS_DIR"
    fi

    # List of Office apps
    local apps=("excel" "word" "powerpoint" "onenote" "outlook", "linoffice")
    local INSTALLED_COUNT=0

    print_info "Processing .desktop files..."
    echo "Number of apps found: ${#apps[@]}"
    echo "Apps are: ${apps[*]}"
    
    for app in "${apps[@]}"; do
        echo "Starting to process app: $app"
        local desktop_file="$DESKTOP_DIR/$app.desktop"
        echo "Processing: $app.desktop"

        # Check if source file exists
        if [ ! -f "$desktop_file" ]; then
            echo "  Error: $app.desktop not found"
            continue
        fi

        # Create corrected .desktop file with absolute paths
        temp_file=$(mktemp) || {
            echo "  Error: Failed to create temporary file"
            continue
        }

        # Replace /PATH/ with LINOFFICE_DIR and write to temp file
        if ! sed "s|/PATH/|$LINOFFICE_DIR/|g" "$desktop_file" > "$temp_file"; then
            echo "  Error: sed command failed"
            rm -f "$temp_file"
            continue
        fi

        # Copy to user applications directory
        if ! cp "$temp_file" "${USER_APPLICATIONS_DIR}/$app.desktop"; then
            echo "  Error: Failed to copy to applications directory"
            rm -f "$temp_file"
            continue
        fi

        # Make it executable
        if ! chmod +x "${USER_APPLICATIONS_DIR}/$app.desktop"; then
            echo "  Error: Failed to make executable"
            rm -f "$temp_file"
            continue
        fi

        # Clean up temp file
        rm -f "$temp_file"

        echo "  Installed: ${USER_APPLICATIONS_DIR}/$app.desktop"
        ((INSTALLED_COUNT++))
        echo "Debug: Finished processing $app"
    done

    print_info "App launchers installed: $INSTALLED_COUNT"

    if [ $INSTALLED_COUNT -gt 0 ]; then
        print_info "Updating desktop database"
        if command -v update-desktop-database >/dev/null 2>&1; then
            update-desktop-database "$USER_APPLICATIONS_DIR" 2>/dev/null || true
            print_success "Desktop database updated"
        else
            print_info "Note: update-desktop-database not found, skipping database update"
        fi

        print_success "Installation complete! The applications should now appear in your application menu."
        print_info "Installed applications:"
        for app in "${apps[@]}"; do
            if [ -f "${USER_APPLICATIONS_DIR}/$app.desktop" ]; then
                display_name=$(grep "^Name=" "$DESKTOP_DIR/$app.desktop" | cut -d'=' -f2)
                echo "  - $display_name"
            fi
        done
        print_info "To uninstall, remove the files from: $USER_APPLICATIONS_DIR"
        print_info "To recreate them, run the script with the --desktop flag."
    else
        print_error "No files were installed."
        return 1
    fi
}

# Function to run InstallOffice.ps1 via FreeRDP
run_install_office_ps1() {
    if [ -z "$FREERDP_COMMAND" ]; then
        detect_freerdp_command
    fi

    print_info "Running Office installation script via FreeRDP..."

    # Build FreeRDP command wrapper (handles flatpak vs system install)
    local freerdp_cmd
    if [[ "$FREERDP_COMMAND" == flatpak* ]]; then
        freerdp_cmd=(bash -c "$FREERDP_COMMAND")
    else
        freerdp_cmd=("$FREERDP_COMMAND")
    fi

    # Connection timeout in milliseconds (e.g. 10000 = 10 seconds)
    local connection_timeout=10000

    # Run FreeRDP with connection timeout only (script keeps running after connect)
    "${freerdp_cmd[@]}" \
        /cert:ignore \
        +home-drive \
        /u:MyWindowsUser \
        /p:MyWindowsPassword \
        /v:127.0.0.1 \
        /port:3388 \
        /timeout:$connection_timeout \
        /app:program:powershell.exe,cmd:'-ExecutionPolicy Bypass -File C:\\OEM\\InstallOffice.ps1' \
        >>"$LOGFILE" 2>&1

    local exit_code=$?
    if [ $exit_code -eq 0 ]; then
        print_success "Office installation script executed successfully via FreeRDP."
        return 0
    else
        print_error "FreeRDP failed to run Office installation script (exit code: $exit_code)"
        return 1
    fi
}

run_install_evkey_ps1() {
    if [ -z "$FREERDP_COMMAND" ]; then
        detect_freerdp_command
    fi

    print_info "Running EVKey installation script via FreeRDP..."

    # Build FreeRDP command wrapper (handles flatpak vs system install)
    local freerdp_cmd
    if [[ "$FREERDP_COMMAND" == flatpak* ]]; then
        freerdp_cmd=(bash -c "$FREERDP_COMMAND")
    else
        freerdp_cmd=("$FREERDP_COMMAND")
    fi

    # Connection timeout in milliseconds (e.g. 10000 = 10 seconds)
    local connection_timeout=10000

    # Run FreeRDP with connection timeout only (script keeps running after connect)
    "${freerdp_cmd[@]}" \
        /cert:ignore \
        +home-drive \
        /u:MyWindowsUser \
        /p:MyWindowsPassword \
        /v:127.0.0.1 \
        /port:3388 \
        /timeout:$connection_timeout \
        /app:program:powershell.exe,cmd:'-ExecutionPolicy Bypass -File C:\\OEM\\InstallEVKey.ps1' \
        >>"$LOGFILE" 2>&1

    local exit_code=$?
    if [ $exit_code -eq 0 ]; then
        print_success "EVKey installation script executed successfully via FreeRDP."
        return 0
    else
        print_error "FreeRDP failed to run EVKey installation script (exit code: $exit_code)"
        return 1
    fi
}

# Main logic
# If --healthcheck flag is set, only run these tests without writing to progress file
if [ "$HEALTHCHECK" = true ]; then
    check_requirements
    check_linoffice_container
    verify_container_health
    if check_available; then
        print_success "Everything seems fine"
    else
        print_error "Unable to connect via RDP"
    fi
    exit 0
fi

init_progress_file

# If --installoffice flag is set, only run InstallOffice.ps1 via FreeRDP
if [ "$INSTALL_OFFICE_ONLY" = true ]; then
    print_info "Running InstallOffice.ps1 via FreeRDP (--installoffice mode)..."
    if run_install_office_ps1; then
        print_success "Office installation script executed successfully!"
    else
        exit_with_error "Failed to run Office installation script via FreeRDP."
    fi
    exit 0
fi

# If --installevkey flag is set, run InstallEVKey.ps1 via FreeRDP
if [ "$INSTALL_EVKEY_ONLY" = true ]; then
    print_info "Running InstallEVKey.ps1 via FreeRDP (--installevkey mode)..."
    if run_install_evkey_ps1; then
        print_success "EVKey installation script executed successfully!"
    else
        exit_with_error "Failed to run EVKey installation script via FreeRDP."
    fi
    exit 0
fi

# If --desktop flag is set, only run desktop_files
if [ "$DESKTOP_ONLY" = true ]; then
    print_info "Recreating desktop files..."
    if desktop_files; then
        mark_progress "$PROGRESS_DESKTOP"
        print_success "App launchers (.desktop files) created successfully!"
    else
        exit_with_error "Failed to create app launchers (.desktop files)"
    fi
    exit 0
fi

# If --firstrun is set, remove the office_installed progress marker
if [ "$FIRSTRUN" = true ]; then
    print_info "--firstrun specified: Forcing RDP and Office install checks."
    if [ -f "$PROGRESS_FILE" ]; then
        sed -i "/$PROGRESS_OFFICE/d" "$PROGRESS_FILE"
    fi
fi

# Check requirements if not already completed
if ! check_progress "$PROGRESS_REQUIREMENTS"; then
    if check_requirements; then
        mark_progress "$PROGRESS_REQUIREMENTS"
    else
        echo "Requirements check failed. Cannot proceed with container setup."
        exit 1
    fi
else
    print_info "Requirements check already completed, skipping..."
fi

# Check container status and create if needed
check_linoffice_container
if [ "$CONTAINER_EXISTS" -eq 0 ] && ! check_progress "$PROGRESS_CONTAINER"; then
    print_info "Container does not exist, proceeding with setup and creation."
    setup_logfile
    if create_container; then
        mark_progress "$PROGRESS_CONTAINER"
    else
        exit_with_error "Container creation failed"
    fi
else
    if check_progress "$PROGRESS_CONTAINER"; then
        print_info "Container already created, skipping creation step."
    else
        print_info "Skipping container creation as LinOffice container already exists."
    fi
fi

# Wait for RDP and check Office installation if not already completed or if --firstrun is set
if ! check_progress "$PROGRESS_OFFICE" || [ "$FIRSTRUN" = true ]; then
    # If --firstrun, ensure the container is running before checking RDP
    if [ "$FIRSTRUN" = true ]; then
        if ! podman ps -q --filter "name=$CONTAINER_NAME" | grep -q .; then
            print_info "Container is not running. Starting LinOffice container for --firstrun..."
            if ! $COMPOSE_COMMAND --file "$COMPOSE_FILE" start; then
                exit_with_error "Failed to start LinOffice container for --firstrun."
            fi
            print_info "Waiting 20 seconds for container to boot..."
            sleep 20
        fi
    fi
    if ! check_available; then
        exit_with_error "Failed to connect to RDP server"
    fi

    if ! check_success; then
        exit_with_error "Office installation failed or timed out"
    fi
    mark_progress "$PROGRESS_OFFICE"
else
    print_info "Office installation already completed, skipping..."
fi

# Install desktop files if not already completed
if ! check_progress "$PROGRESS_DESKTOP"; then
    if desktop_files; then
        mark_progress "$PROGRESS_DESKTOP"
    else
        exit_with_error "Failed to install desktop files"
    fi
else
    print_info "Desktop files already installed, skipping this step. To recreate them, run the script with the --desktop flag."
fi

# Clean up success file
rm -f "$SUCCESS_FILE"

print_success "LinOffice setup completed successfully!"
