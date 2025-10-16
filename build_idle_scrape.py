#!/usr/bin/env python
"""
Build script untuk Idle Scrape v2 menggunakan PyInstaller
"""
import os
import sys
import shutil
import subprocess
import PyInstaller.__main__

# Nama aplikasi
APP_NAME = "Idle Scrape v2"

# File ikon (jika ada)
ICON_FILE = None
if os.path.exists('important_files/app_icon.ico'):
    ICON_FILE = os.path.abspath('important_files/app_icon.ico')
elif os.path.exists('app_icon.ico'):
    ICON_FILE = os.path.abspath('app_icon.ico')

# Pastikan direktori build dan dist bersih
try:
    if os.path.exists("build"):
        shutil.rmtree("build")
    if os.path.exists("dist"):
        shutil.rmtree("dist")
except Exception as e:
    print(f"Warning: Could not remove old directories: {str(e)}")
    print("Continuing build process anyway...")

# Windows manifest content - untuk meminta akses file yang sesuai
manifest_content = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<assembly xmlns="urn:schemas-microsoft-com:asm.v1" manifestVersion="1.0">
  <assemblyIdentity type="win32" name="IdleScrape.Maps.Scraper" version="1.0.0.0" processorArchitecture="*"/>
  <trustInfo xmlns="urn:schemas-microsoft-com:asm.v3">
    <security>
      <requestedPrivileges>
        <requestedExecutionLevel level="asInvoker" uiAccess="false"/>
      </requestedPrivileges>
    </security>
  </trustInfo>
  <compatibility xmlns="urn:schemas-microsoft-com:compatibility.v1">
    <application>
      <supportedOS Id="{8e0f7a12-bfb3-4fe8-b9a5-48fd50a15a9a}"/>
      <supportedOS Id="{1f676c76-80e1-4239-95bb-83d0f6d0da78}"/>
      <supportedOS Id="{e2011457-1546-43c5-a5fe-008deee3d3f0}"/>
      <supportedOS Id="{35138b9a-5d96-4fbd-8e2d-a2440225f93a}"/>
      <supportedOS Id="{4a2f28e3-53b9-4441-ba9c-d69d4a4a6e38}"/>
    </application>
  </compatibility>
</assembly>
"""

# Tulis manifest ke file
manifest_file = "idle_scrape.manifest"
with open(manifest_file, "w") as f:
    f.write(manifest_content)

# Dapatkan semua file Python yang diperlukan
python_files = ["import_excel_gui.py"]

# Dapatkan direktori yang perlu disertakan
additional_dirs = []
if os.path.exists('src'):
    additional_dirs.append('src')
if os.path.exists('important_files'):
    additional_dirs.append('important_files')
if os.path.exists('chromedriver_dir'):
    additional_dirs.append('chromedriver_dir')

# Siapkan parameter PyInstaller
params = [
    "--name=%s" % APP_NAME,
    "--onedir",
    "--windowed",
    "--clean",
    "--add-data=requirements.txt;.",
    f"--manifest={manifest_file}"  # Tambahkan manifest
]

# Tambahkan ikon jika ada
if ICON_FILE:
    params.append("--icon=%s" % ICON_FILE)

# Tambahkan direktori tambahan
for dir_path in additional_dirs:
    params.append("--add-data=%s;%s" % (dir_path, dir_path))

# Tambahkan file utama
params.append("import_excel_gui.py")

# Cetak informasi build
print("="*60)
print(f"Building {APP_NAME}...")
print(f"Main files: {python_files}")
print(f"Additional directories: {additional_dirs}")
print(f"Icon file: {ICON_FILE}")
print(f"Using Windows manifest for file permissions")
print("="*60)

# Jalankan PyInstaller
PyInstaller.__main__.run(params)

# Salin file tambahan jika diperlukan
# (Contoh: README, license, dll)
if os.path.exists("README.md"):
    shutil.copy("README.md", f"dist/{APP_NAME}/")

# Hapus file manifest sementara
if os.path.exists(manifest_file):
    os.remove(manifest_file)

print("="*60)
print(f"Build selesai! Executable ada di dist/{APP_NAME}/")
print("="*60) 