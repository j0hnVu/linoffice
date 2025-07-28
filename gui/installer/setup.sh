#!/bin/bash
# this is just a tester script, not the actual one
# Colors for output
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

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

exit_with_error() {
    print_error "$1"
    exit 1
}

print_step "1" "Checking requirements"
sleep 1
echo "normal output"
print_info "info"
print_error "warning"
#exit_with_error "error!!!"
print_step "2" "Installing dependencies"
sleep 1
print_step "3" "Configuring environment"
sleep 1
print_info "info"
print_info "info"
print_info "info"
print_info "info"
print_info "info"
print_info "info"
print_step "4" "another step"
print_step "5" "another step"
print_step "6" "another step"
print_step "7" "another step"
print_step "8" "install desktop files"
print_success "all done"
