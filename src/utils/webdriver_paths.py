import os
import sys
import platform

def get_base_path():
    """Get the base path for the application"""
    # When running as script
    if getattr(sys, 'frozen', False):
        # Running as compiled executable
        return os.path.dirname(sys.executable)
    else:
        # Running as a normal Python script
        return os.path.dirname(os.path.abspath(__file__))

def get_chromedriver_path():
    """Get the path to the chromedriver executable"""
    base_path = get_base_path()
    
    # Determine the appropriate chromedriver for the platform
    system = platform.system()
    if system == 'Windows':
        driver_path = os.path.join(base_path, 'chromedriver_dir', 'chromedriver.exe')
    elif system == 'Darwin':  # macOS
        driver_path = os.path.join(base_path, 'chromedriver_dir', 'chromedriver')
    else:  # Linux and others
        driver_path = os.path.join(base_path, 'chromedriver_dir', 'chromedriver')
    
    # Check if the driver exists
    if not os.path.exists(driver_path):
        print(f"ChromeDriver not found at {driver_path}")
        # Use ChromeDriverManager as fallback
        try:
            from webdriver_manager.chrome import ChromeDriverManager
            driver_path = ChromeDriverManager().install()
            print(f"Using ChromeDriverManager: {driver_path}")
        except Exception as e:
            print(f"Error using ChromeDriverManager: {str(e)}")
    
    return driver_path 