#!/usr/bin/env bash

set -euo pipefail
APT_UPDATED=0

REPO_OWNER="eylenburg"
REPO_NAME="linoffice"
TARGET_DIR="$HOME/.local/bin/linoffice"
TMPDIR=$(mktemp -d)
GITHUB_API_URL="https://api.github.com/repos/${REPO_OWNER}/${REPO_NAME}/releases"

LINOFFICE_SCRIPT="$TARGET_DIR/gui/linoffice.py"

##################################################
# PART 1: INSTALL DEPENDENCIES
##################################################

# Detect package manager and distro
detect_package_manager() {
  if [ -f /etc/os-release ]; then
    . /etc/os-release
    DISTRO_ID=$ID
    DISTRO_LIKE=${ID_LIKE:-}
  else
    echo "Cannot determine OS version"
    exit 1
  fi

  # Reject immutable systems
  case "$DISTRO_ID" in
    nixos|guix|silverblue|coreos|kinoite|microos)
      echo "Unsupported system type: $DISTRO_ID"
      exit 1
      ;;
  esac

  for mgr in apt dnf yum zypper pacman xbps-install eopkg urpmi; do
    if command -v "$mgr" >/dev/null 2>&1; then
      PKG_MGR=$mgr
      return
    fi
  done

  echo "Unsupported package manager"
  exit 1
}

# Generic install function
install_pkg() {
  local pkg="$1"
  case "$PKG_MGR" in
    apt)
      if [ "$APT_UPDATED" -eq 0 ]; then
        sudo apt-get update
        APT_UPDATED=1
      fi
      sudo apt-get install -y "$pkg"
      ;;
    dnf|yum)
      sudo "$PKG_MGR" install -y "$pkg"
      ;;
    zypper)
      sudo zypper --non-interactive install "$pkg"
      ;;
    pacman)
      sudo pacman -Syu --noconfirm "$pkg"
      ;;
    xbps-install)
      sudo xbps-install -Sy "$pkg"
      ;;
    eopkg)
      sudo eopkg install -y "$pkg"
      ;;
    urpmi)
      sudo urpmi --auto "$pkg"
      ;;
    *)
      echo "Unknown package manager: $PKG_MGR"
      exit 1
      ;;
  esac
}

# Try list of possible package names
try_install_any() {
  for name in "$@"; do
    if install_pkg "$name"; then
      return 0
    fi
  done
  return 1
}

# Command check
pkg_exists() {
  command -v "$1" >/dev/null 2>&1
}

ensure_pip() {
  if ! pkg_exists pip3 && ! pkg_exists pip; then
    try_install_any python3-pip python-pip python-pip3 || {
      echo "Failed to install pip"
      exit 1
    }
  fi
}

freerdp_version_ok() {
  if command -v xfreerdp >/dev/null 2>&1; then
    ver=$(xfreerdp --version | grep -oP '\d+\.\d+\.\d+' | head -n1)
    major=$(echo "$ver" | cut -d. -f1)
    [ "$major" -ge 3 ]
  else
    return 1
  fi
}

# Flatpak fallback
install_freerdp_flatpak() {
  if ! pkg_exists flatpak; then
    try_install_any flatpak || { echo "Failed to install flatpak"; exit 1; }
  fi

  if ! flatpak remote-list | grep -q flathub; then
    flatpak remote-add --if-not-exists flathub https://flathub.org/repo/flathub.flatpakrepo
  fi

  flatpak install -y --user flathub com.freerdp.FreeRDP
}

# Main script for installing dependencies
dependencies_main() {
  detect_package_manager

  echo "Detected distro: $DISTRO_ID, package manager: $PKG_MGR"

  echo "Checking podman..."
  if ! pkg_exists podman; then
    try_install_any podman || { echo "Failed to install podman"; exit 1; }
  fi

  echo "Checking Python 3..."
  if ! pkg_exists python3 && ! pkg_exists python; then
    try_install_any python3 python || { echo "Failed to install Python 3"; exit 1; }
  fi

  echo "Checking FreeRDP (version >= 3)..."
  if freerdp_version_ok; then
    echo "FreeRDP is already version >= 3"
  else
    try_install_any freerdp3 freerdp3-x11 freerdp || true
    if ! freerdp_version_ok; then
      echo "Falling back to Flatpak for FreeRDP"
      install_freerdp_flatpak
    fi
  fi

  echo "Checking podman-compose..."
  if ! pkg_exists podman-compose; then
    try_install_any podman-compose || true
    if ! pkg_exists podman-compose; then
      echo "Trying pip fallback for podman-compose"
      ensure_pip
      pip3 install --user --break-system-packages podman-compose || pip install --user --break-system-packages podman-compose
      export PATH="$HOME/.local/bin:$PATH"
      if ! pkg_exists podman-compose; then
        echo "podman-compose still not available after pip install"
        exit 1
      fi
    fi
  fi

  echo "Checking PySide6..."
  if ! python3 -c "import PySide6" 2>/dev/null; then
    try_install_any python3-pyside6 python-pyside6 python3-pyside6.qtcore || true
    if ! python3 -c "import PySide6" 2>/dev/null; then
      echo "Falling back to pip for PySide6"
      ensure_pip
      pip3 install --user --break-system-packages PySide6 || pip install --user --break-system-packages PySide6
      if ! python3 -c "import PySide6" 2>/dev/null; then
        echo "Failed to install PySide6"
        exit 1
      fi
    fi
  fi


  echo "✅ All dependencies installed successfully!"
}


##################################################
# PART 2: DOWNLOAD LATEST LINOFFICE
##################################################

download_latest() {
  echo "Fetching latest LinOffice version from GitHub..."

  LATEST_VERSION=$(curl -sSL "$GITHUB_API_URL" | \
    grep -E '"tag_name":\s*"v[0-9]+\.[0-9]+\.[0-9]+"' | \
    grep -v '"prerelease": true' | \
    grep -v '"draft": true' | \
    head -n1 | \
    sed -E 's/.*"v([0-9]+\.[0-9]+\.[0-9]+)".*/\1/')

  if [[ -z "$LATEST_VERSION" ]]; then
    echo "Error: Could not determine latest version."
    exit 1
  fi

  ZIP_URL="https://github.com/${REPO_OWNER}/${REPO_NAME}/archive/refs/tags/v${LATEST_VERSION}.zip"
  ZIP_FILE="$TMPDIR/linoffice.zip"

  echo "Downloading Linoffice v${LATEST_VERSION} from:"
  echo "$ZIP_URL"

  # Modified: Add -L to follow redirects during status check
  HTTP_STATUS=$(curl -s -L -o /dev/null -w "%{http_code}" "$ZIP_URL")

  if [[ "$HTTP_STATUS" != "200" ]]; then
    echo "Error: File not found at $ZIP_URL (HTTP $HTTP_STATUS)"
    exit 1
  fi

  curl -L -o "$ZIP_FILE" "$ZIP_URL"

  # Check it’s really a zip
  if ! file "$ZIP_FILE" | grep -q "Zip archive data"; then
    echo "Error: Downloaded file is not a valid zip archive."
    file "$ZIP_FILE"
    exit 1
  fi

  echo "Unzipping..."
  unzip -q "$ZIP_FILE" -d "$TMPDIR"

  # Expected folder name: linoffice-${LATEST_VERSION}
  EXTRACTED_DIR="$TMPDIR/linoffice-${LATEST_VERSION}"

  if [[ ! -d "$EXTRACTED_DIR" ]]; then
    echo "Error: Expected folder 'linoffice-${LATEST_VERSION}' not found in zip."
    exit 1
  fi

  echo "Installing to $TARGET_DIR..."

  # Check if TARGET_DIR exists and contains linoffice.sh
  if [[ -d "$TARGET_DIR" && -f "$TARGET_DIR/linoffice.sh" ]]; then
    echo "Existing installation found. Updating files..."
    cp -r -u "$EXTRACTED_DIR/"* "$TARGET_DIR/"
  else
    # If TARGET_DIR doesn't exist or doesn't contain linoffice.sh, replace it
    rm -rf "$TARGET_DIR"
    mkdir -p "$(dirname "$TARGET_DIR")"
    mv "$EXTRACTED_DIR" "$TARGET_DIR"
  fi

  # Make everything executable
  find "$TARGET_DIR" -type f \( -name "*.py" -o -name "*.sh" \) -exec chmod +x {} \;

  rm -rf "$TMPDIR"

  echo "✅ Linoffice v${LATEST_VERSION} installed at $TARGET_DIR"
}

##################################################
# PART 3: Run LinOffice
##################################################

start_linoffice() {
  # Run the linoffice installer
  echo "Starting Linoffice..."

  # Check if python3 exists
  if command -v python3 >/dev/null 2>&1; then
    PYTHON_CMD="python3"
  elif command -v python >/dev/null 2>&1; then
    PYTHON_CMD="python"
  else
    echo "Error: Python not found. Please install Python."
    exit 1
  fi

  # Check if the linoffice.py script exists
  if [[ ! -f "$LINOFFICE_SCRIPT" ]]; then
    echo "Error: $LINOFFICE_SCRIPT not found. Please check the installation."
    exit 1
  fi

  echo "Running $LINOFFICE_SCRIPT with $PYTHON_CMD..."
  nohup "$PYTHON_CMD" "$LINOFFICE_SCRIPT" &
}

##################################################
# Main logic
##################################################

read -p "Welcome to the LinOffice installer. We will check and install dependencies, download the latest LinOffice release, and then run the main setup, which will install a Windows container with Microsoft Office. Are you sure you want to continue? (y/n): " confirmation
if [[ "$confirmation" == "y" || "$confirmation" == "Y" ]]; then
  dependencies_main "$@"
  download_latest
  start_linoffice
else
  echo "Cancelled."
  exit 1
fi
