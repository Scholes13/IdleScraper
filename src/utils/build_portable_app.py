import os
import sys
import shutil
import subprocess
import zipfile
import platform
import requests
from io import BytesIO
from zipfile import ZipFile

def download_chromedriver():
    """Download the appropriate ChromeDriver based on platform and Chrome version"""
    print("Setting up ChromeDriver...")
    
    # Create chromedriver_dir if it doesn't exist
    if not os.path.exists("chromedriver_dir"):
        os.makedirs("chromedriver_dir")
    
    # Use a known working version instead of the latest
    version = "114.0.5735.90"
    print(f"Using ChromeDriver version: {version}")
    
    # Build the download URL based on platform
    system = platform.system()
    
    if system == "Windows":
        try:
            # Use the local chromedriver-win64.zip file
            local_zip_path = "chromedriver-win64.zip"
            if os.path.exists(local_zip_path):
                print(f"Using local ChromeDriver zip: {local_zip_path}")
                
                # Create a temporary directory for extraction
                temp_extract_dir = "temp_chromedriver_extract"
                if os.path.exists(temp_extract_dir):
                    shutil.rmtree(temp_extract_dir)
                os.makedirs(temp_extract_dir)
                
                # Extract the zip file to the temporary directory
                with zipfile.ZipFile(local_zip_path, 'r') as zip_ref:
                    zip_ref.extractall(temp_extract_dir)
                
                # The structure should be: temp_extract_dir/chromedriver-win64/chromedriver.exe
                nested_dir = os.path.join(temp_extract_dir, "chromedriver-win64")
                
                if os.path.exists(nested_dir):
                    # Copy all files from nested directory to chromedriver_dir
                    files_copied = 0
                    for file in os.listdir(nested_dir):
                        src_path = os.path.join(nested_dir, file)
                        dst_path = os.path.join("chromedriver_dir", file)
                        if os.path.isfile(src_path):
                            shutil.copy2(src_path, dst_path)
                            print(f"Copied {file} to chromedriver_dir")
                            files_copied += 1
                    
                    if files_copied > 0:
                        print(f"Successfully copied {files_copied} files including chromedriver.exe")
                    else:
                        print("No files were copied from the extracted archive")
                else:
                    print(f"Expected nested directory structure not found: {nested_dir}")
                    print("Contents of extracted archive:")
                    for root, dirs, files in os.walk(temp_extract_dir):
                        print(f"Directory: {root}")
                        for file in files:
                            print(f"  - {file}")
                
                # Clean up temporary directory
                shutil.rmtree(temp_extract_dir)
                
                # Verify that chromedriver.exe exists
                chromedriver_exe = os.path.join("chromedriver_dir", "chromedriver.exe")
                if os.path.exists(chromedriver_exe):
                    print(f"ChromeDriver executable found at: {chromedriver_exe}")
                    return True
                else:
                    print("ERROR: chromedriver.exe not found after extraction!")
                    print("Will try to download from internet as fallback")
            
            # Fallback to download if local file doesn't exist or extraction failed
            print("Downloading ChromeDriver from alternative source...")
            alt_url = "https://edgedl.me.gvt1.com/edgedl/chrome/chrome-for-testing/114.0.5735.90/win64/chromedriver-win64.zip"
            response = requests.get(alt_url, stream=True)
            response.raise_for_status()
            
            # Save the zip file
            zip_path = os.path.join("chromedriver_dir", "chromedriver.zip")
            with open(zip_path, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    f.write(chunk)
            
            # Extract the zip file directly to a temporary location
            temp_extract_dir = "temp_chromedriver_extract"
            if os.path.exists(temp_extract_dir):
                shutil.rmtree(temp_extract_dir)
            os.makedirs(temp_extract_dir)
            
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                zip_ref.extractall(temp_extract_dir)
            
            # Move the chromedriver.exe from the nested folder to the main chromedriver_dir
            nested_dir = os.path.join(temp_extract_dir, "chromedriver-win64")
            if os.path.exists(nested_dir):
                for file in os.listdir(nested_dir):
                    src_path = os.path.join(nested_dir, file)
                    dst_path = os.path.join("chromedriver_dir", file)
                    if os.path.isfile(src_path):
                        shutil.copy2(src_path, dst_path)
                        print(f"Copied {file} to chromedriver_dir")
            
            # Clean up
            shutil.rmtree(temp_extract_dir)
            
            # Verify chromedriver.exe exists
            chromedriver_exe = os.path.join("chromedriver_dir", "chromedriver.exe")
            if os.path.exists(chromedriver_exe):
                print(f"ChromeDriver executable found at: {chromedriver_exe}")
                return True
            else:
                print("ERROR: chromedriver.exe not found after download!")
                return False
                
        except Exception as e:
            print(f"Error setting up ChromeDriver: {e}")
            return False
    else:
        print("Unsupported platform for this script. Please download ChromeDriver manually.")
        return False

def run_pyinstaller():
    """Run PyInstaller to create the executable"""
    print("Building executable with PyInstaller...")
    
    # Check if logonew.ico exists and use it directly
    if os.path.exists("logonew.ico"):
        print("Using logonew.ico as application icon")
        # Copy logonew.ico to app_icon.ico for consistency with spec file
        shutil.copy2("logonew.ico", "app_icon.ico")
    # Otherwise, check if PNG logo exists and convert to ICO
    elif os.path.exists("app_logo.png"):
        try:
            # Import PIL only when needed
            from PIL import Image
            print("Converting app_logo.png to ICO format...")
            img = Image.open("app_logo.png")
            
            # Create icon file
            if not os.path.exists("app_icon.ico"):
                icon_sizes = [(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]
                img.save("app_icon.ico", format="ICO", sizes=icon_sizes)
                print("Successfully converted app_logo.png to app_icon.ico")
        except Exception as e:
            print(f"Error converting logo to ICO: {e}")
            # Fallback to old method if conversion fails
            if not os.path.exists("app_icon.ico"):
                try:
                    print("Falling back to default icon creation...")
                    import create_icon
                    create_icon.create_app_icon()
                except Exception as e2:
                    print(f"Error creating icon: {e2}")
    else:
        print("No custom icon found")
        # Fallback to default icon if PNG doesn't exist
        if not os.path.exists("app_icon.ico"):
            try:
                import create_icon
                create_icon.create_app_icon()
            except Exception as e:
                print(f"Error creating icon: {e}")
    
    # Run PyInstaller
    try:
        result = subprocess.run([
            "pyinstaller", 
            "--clean",
            "--noconfirm", 
            "maps_scraper_exe.spec"
        ], check=True)
        
        if result.returncode == 0:
            print("PyInstaller completed successfully")
            return True
        else:
            print(f"PyInstaller failed with return code {result.returncode}")
            return False
    except Exception as e:
        print(f"Error running PyInstaller: {e}")
        return False

def create_portable_package():
    """Create a portable package with all necessary files"""
    print("Creating portable package...")
    
    # Define source and destination paths
    dist_dir = os.path.join(os.getcwd(), "dist")
    output_dir = os.path.join(os.getcwd(), "Werkudara - Google Maps Scraper Portable")
    
    # Create output directory if it doesn't exist
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # Copy the executable and dependencies
    try:
        exe_name = "Werkudara - Google Maps Scraper.exe"
        exe_path = os.path.join(dist_dir, exe_name)
        
        if os.path.exists(exe_path):
            shutil.copy2(exe_path, output_dir)
            print(f"Copied {exe_name} to {output_dir}")
        else:
            print(f"Error: {exe_path} not found")
            return False
        
        # Create chromedriver directory in the output
        chrome_dir = os.path.join(output_dir, "chromedriver_dir")
        if not os.path.exists(chrome_dir):
            os.makedirs(chrome_dir)
        
        # Copy chromedriver files
        for file in os.listdir("chromedriver_dir"):
            src = os.path.join("chromedriver_dir", file)
            dst = os.path.join(chrome_dir, file)
            if os.path.isfile(src):
                shutil.copy2(src, dst)
                print(f"Copied {file} to {chrome_dir}")
        
        # Create a README file
        readme_path = os.path.join(output_dir, "README.txt")
        with open(readme_path, "w") as f:
            f.write("""Werkudara - Google Maps Scraper
===========================

Cara Penggunaan:
1. Double-click pada file "Werkudara - Google Maps Scraper.exe"
2. Pilih file Excel dengan kolom: 
   - Nama Perusahaan
   - Alamat
   - Kecamatan
3. Klik "Load Excel" untuk memuat data
4. Pilih apakah ingin mengaktifkan Smart Search 
   (pencarian otomatis perusahaan dengan nama serupa)
5. Klik "Process Data" untuk mulai mencari data perusahaan
6. Tunggu hingga proses selesai dan simpan hasilnya

Catatan:
- Program ini memerlukan koneksi internet
- Pastikan Google Chrome sudah terinstal di komputer Anda
- Jika terjadi masalah, coba nonaktifkan Smart Search

Version 1.0
""")
            print("Created README.txt")
        
        print(f"Portable package created successfully at {output_dir}")
        return True
    except Exception as e:
        print(f"Error creating portable package: {e}")
        return False

def create_zip_package():
    """Create a ZIP archive of the portable package"""
    print("Creating ZIP archive...")
    
    try:
        output_dir = "Werkudara - Google Maps Scraper Portable"
        zip_name = "Werkudara - Google Maps Scraper Portable.zip"
        
        with zipfile.ZipFile(zip_name, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, dirs, files in os.walk(output_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    zipf.write(file_path, os.path.relpath(file_path, os.path.dirname(output_dir)))
        
        print(f"ZIP archive created: {zip_name}")
        return True
    except Exception as e:
        print(f"Error creating ZIP archive: {e}")
        return False

if __name__ == "__main__":
    print("Starting build process...")
    
    # Step 1: Download ChromeDriver
    if not download_chromedriver():
        print("Failed to download ChromeDriver. Build aborted.")
        sys.exit(1)
    
    # Step 2: Run PyInstaller
    if not run_pyinstaller():
        print("Failed to build executable. Build aborted.")
        sys.exit(1)
    
    # Step 3: Create portable package
    if not create_portable_package():
        print("Failed to create portable package. Build aborted.")
        sys.exit(1)
    
    # Step 4: Create ZIP archive
    if not create_zip_package():
        print("Failed to create ZIP archive.")
    
    print("Build process completed successfully!")
    print("You can distribute 'Werkudara - Google Maps Scraper Portable.zip' to users.") 