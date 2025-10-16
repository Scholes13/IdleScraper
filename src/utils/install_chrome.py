import os
import sys
import platform
import subprocess
import urllib.request
import zipfile
import shutil
import tempfile
from pathlib import Path

def get_chrome_version():
    """Get the installed Chrome version"""
    system = platform.system()
    try:
        if system == "Windows":
            # Check common installation paths
            paths = [
                r"C:\Program Files\Google\Chrome\Application\chrome.exe",
                r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
                os.path.expanduser(r"~\AppData\Local\Google\Chrome\Application\chrome.exe")
            ]
            
            for path in paths:
                if os.path.exists(path):
                    # Use PowerShell to get the version
                    cmd = f'(Get-Item "{path}").VersionInfo.ProductVersion'
                    process = subprocess.run(["powershell", "-Command", cmd], capture_output=True, text=True)
                    version = process.stdout.strip()
                    print(f"Detected Chrome version: {version}")
                    return version
            
            print("Chrome not found in common locations")
            return None
            
        elif system == "Linux":
            process = subprocess.run(["google-chrome", "--version"], capture_output=True, text=True)
            version = process.stdout.strip().split()[-1]
            print(f"Detected Chrome version: {version}")
            return version
            
        elif system == "Darwin":  # macOS
            process = subprocess.run(["/Applications/Google Chrome.app/Contents/MacOS/Google Chrome", "--version"], 
                                   capture_output=True, text=True)
            version = process.stdout.strip().split()[-1]
            print(f"Detected Chrome version: {version}")
            return version
            
    except Exception as e:
        print(f"Error detecting Chrome version: {str(e)}")
        return None

def install_chrome():
    """Provide instructions to install Chrome browser"""
    system = platform.system()
    
    print("\n=== Chrome Installation Guide ===")
    print("You need to have Google Chrome installed to use this scraper.")
    
    if system == "Windows":
        print("Download Chrome from: https://www.google.com/chrome/")
        print("Or install it via winget (Windows Package Manager):")
        print("winget install Google.Chrome")
        
    elif system == "Linux":
        print("For Debian/Ubuntu systems:")
        print("wget https://dl.google.com/linux/direct/google-chrome-stable_current_amd64.deb")
        print("sudo apt install ./google-chrome-stable_current_amd64.deb")
        
        print("\nFor Fedora/RHEL systems:")
        print("sudo dnf install google-chrome-stable")
        
    elif system == "Darwin":  # macOS
        print("Download Chrome from: https://www.google.com/chrome/")
        print("Or install via Homebrew:")
        print("brew install --cask google-chrome")
    
    print("\nAfter installing Chrome, run this script again.")

def download_chromedriver(version_prefix):
    """Download the appropriate ChromeDriver version"""
    system = platform.system()
    if system == "Windows":
        platform_name = "win32"
    elif system == "Linux":
        platform_name = "linux64"
    elif system == "Darwin":
        # Check if M1/M2 Mac
        if platform.machine() == "arm64":
            platform_name = "mac_arm64"
        else:
            platform_name = "mac64"
    else:
        print(f"Unsupported platform: {system}")
        return None
    
    # Get major version
    major_version = version_prefix.split('.')[0]
    
    # ChromeDriver download URL
    if int(major_version) >= 115:
        # For Chrome 115+, use the new URL format
        base_url = f"https://storage.googleapis.com/chrome-for-testing-public/{version_prefix}"
        driver_url = f"{base_url}/chromedriver-{platform_name}.zip"
    else:
        # For older Chrome versions
        base_url = f"https://chromedriver.storage.googleapis.com/"
        driver_url = f"{base_url}{version_prefix}/chromedriver_{platform_name}.zip"
    
    print(f"Downloading ChromeDriver from: {driver_url}")
    
    try:
        # Create a temporary directory
        temp_dir = tempfile.mkdtemp()
        zip_path = os.path.join(temp_dir, "chromedriver.zip")
        
        # Download the file
        urllib.request.urlretrieve(driver_url, zip_path)
        
        # Extract the zip
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)
        
        # Find the chromedriver executable
        if system == "Windows":
            chromedriver_name = "chromedriver.exe"
        else:
            chromedriver_name = "chromedriver"
            
        # For Chrome 115+, the driver is in a subdirectory
        if int(major_version) >= 115:
            driver_path = os.path.join(temp_dir, f"chromedriver-{platform_name}", chromedriver_name)
        else:
            driver_path = os.path.join(temp_dir, chromedriver_name)
        
        if not os.path.exists(driver_path):
            # Try to find it recursively
            for root, dirs, files in os.walk(temp_dir):
                if chromedriver_name in files:
                    driver_path = os.path.join(root, chromedriver_name)
                    break
        
        # Make it executable on Unix
        if system != "Windows":
            os.chmod(driver_path, 0o755)
        
        # Create a directory for the driver if it doesn't exist
        driver_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "drivers")
        os.makedirs(driver_dir, exist_ok=True)
        
        # Copy the driver to our application directory
        dest_path = os.path.join(driver_dir, chromedriver_name)
        shutil.copy2(driver_path, dest_path)
        
        print(f"ChromeDriver installed successfully at: {dest_path}")
        return dest_path
        
    except Exception as e:
        print(f"Error downloading ChromeDriver: {str(e)}")
        return None
    finally:
        # Clean up the temporary directory
        try:
            shutil.rmtree(temp_dir)
        except:
            pass

def setup_environment():
    """Set up the environment for the scraper"""
    print("Setting up environment for Google Maps Scraper...")
    
    # Check if Chrome is installed
    chrome_version = get_chrome_version()
    
    if not chrome_version:
        print("Google Chrome is not installed or could not be detected.")
        install_chrome()
        return False
    
    # Get major.minor.build version (first 3 parts)
    version_parts = chrome_version.split('.')
    if len(version_parts) >= 3:
        version_prefix = '.'.join(version_parts[:3])
    else:
        version_prefix = chrome_version
    
    # Download compatible ChromeDriver
    driver_path = download_chromedriver(version_prefix)
    
    if driver_path:
        # Add the driver directory to PATH
        driver_dir = os.path.dirname(driver_path)
        if driver_dir not in os.environ['PATH']:
            os.environ['PATH'] = f"{driver_dir}{os.pathsep}{os.environ['PATH']}"
        
        print("\nEnvironment setup successful!")
        print(f"Chrome version: {chrome_version}")
        print(f"ChromeDriver path: {driver_path}")
        print("\nYou can now run the scraper scripts.")
        return True
    else:
        print("\nFailed to set up ChromeDriver.")
        print("Please try again or manually download ChromeDriver from:")
        print("https://chromedriver.chromium.org/downloads")
        return False

if __name__ == "__main__":
    print("Google Maps Scraper - Setup Tool")
    print("================================")
    
    setup_environment() 