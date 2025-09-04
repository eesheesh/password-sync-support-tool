#!/bin/bash
#
# This script sets up a Debian-based Linux environment for testing the
# PasswordSyncSupportTool.ps1 script. It installs PowerShell and Pester.

# Exit on any error
set -e

# 1. Update package lists
echo "Updating package lists..."
sudo apt-get update

# 2. Install prerequisites
echo "Installing prerequisites..."
sudo apt-get install -y wget apt-transport-https software-properties-common

# 3. Add Microsoft package repository
echo "Adding Microsoft package repository..."
wget -q "https://packages.microsoft.com/config/ubuntu/22.04/packages-microsoft-prod.deb"
sudo dpkg -i packages-microsoft-prod.deb
rm packages-microsoft-prod.deb

# 4. Update package lists again
echo "Updating package lists again..."
sudo apt-get update

# 5. Install PowerShell
echo "Installing PowerShell..."
sudo apt-get install -y powershell

# 6. Install Pester module
echo "Installing Pester PowerShell module..."
pwsh -Command "Install-Module -Name Pester -Force -SkipPublisherCheck -AcceptLicense"

echo "Setup complete."
