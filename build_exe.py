#!/usr/bin/env python3
"""
Build executable for Glucose Converter
Creates standalone .exe file for Windows
"""

import sys
import os
import subprocess
import shutil
from pathlib import Path


def check_requirements():
    """Check if required packages are installed"""
    required = ['PyInstaller', 'openpyxl']
    missing = []
    
    for package in required:
        try:
            __import__(package.lower())
        except ImportError:
            missing.append(package)
    
    if missing:
        print(f"❌ Missing packages: {', '.join(missing)}")
        print("\nInstalling missing packages...")
        subprocess.check_call([sys.executable, '-m', 'pip', 'install'] + missing)
        print("✅ Packages installed successfully")
    else:
        print("✅ All required packages are installed")


def create_spec_file():
    """Create PyInstaller spec file for better control"""
    spec_content = """# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['glucose_converter_gui.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('glucose_converter.py', '.'),
    ],
    hiddenimports=['openpyxl', 'tkinter'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='GlucoseConverter',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # Set to False for GUI app (no console window)
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico' if os.path.exists('icon.ico') else None,
)
"""
    
    with open('glucose_converter.spec', 'w') as f:
        f.write(spec_content)
    
    print("✅ Created PyInstaller spec file")


def build_executable():
    """Build the executable using PyInstaller"""
    print("\n🔨 Building executable...")
    
    # Build command
    cmd = [
        sys.executable, '-m', 'PyInstaller',
        '--onefile',  # Single file executable
        '--windowed',  # No console window for GUI app
        '--name', 'GlucoseConverter',
        '--clean',  # Clean build
        '--noconfirm',  # Overwrite without asking
    ]
    
    # Add icon if exists
    if os.path.exists('icon.ico'):
        cmd.extend(['--icon', 'icon.ico'])
    
    # Add the GUI script
    cmd.append('glucose_converter_gui.py')
    
    # Run PyInstaller
    try:
        subprocess.check_call(cmd)
        print("✅ Build completed successfully!")
        
        # Check if exe was created
        exe_path = Path('dist') / 'GlucoseConverter.exe'
        if exe_path.exists():
            print(f"\n📦 Executable created: {exe_path}")
            print(f"   Size: {exe_path.stat().st_size / (1024*1024):.2f} MB")
            return str(exe_path)
        else:
            print("❌ Executable not found in dist folder")
            return None
            
    except subprocess.CalledProcessError as e:
        print(f"❌ Build failed: {e}")
        return None


def create_installer_script():
    """Create NSIS installer script for professional installation"""
    nsis_script = """!define APPNAME "Glucose Converter"
!define COMPANYNAME "Your Company"
!define DESCRIPTION "Convert Contour Plus glucose CSV to formatted XLSX"
!define VERSIONMAJOR 1
!define VERSIONMINOR 0
!define VERSIONBUILD 0
!define HELPURL "http://yourwebsite.com"
!define UPDATEURL "http://yourwebsite.com"
!define ABOUTURL "http://yourwebsite.com"

RequestExecutionLevel admin

InstallDir "$PROGRAMFILES\\${APPNAME}"

Name "${APPNAME}"
OutFile "GlucoseConverter_Setup.exe"

!include "MUI2.nsh"

!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH

!insertmacro MUI_UNPAGE_WELCOME
!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES
!insertmacro MUI_UNPAGE_FINISH

!insertmacro MUI_LANGUAGE "English"

Section "install"
    SetOutPath $INSTDIR
    
    File "dist\\GlucoseConverter.exe"
    
    # Create uninstaller
    WriteUninstaller "$INSTDIR\\uninstall.exe"
    
    # Start Menu
    CreateDirectory "$SMPROGRAMS\\${APPNAME}"
    CreateShortcut "$SMPROGRAMS\\${APPNAME}\\${APPNAME}.lnk" "$INSTDIR\\GlucoseConverter.exe"
    CreateShortcut "$SMPROGRAMS\\${APPNAME}\\Uninstall.lnk" "$INSTDIR\\uninstall.exe"
    
    # Desktop shortcut
    CreateShortcut "$DESKTOP\\${APPNAME}.lnk" "$INSTDIR\\GlucoseConverter.exe"
    
    # Registry information for add/remove programs
    WriteRegStr HKLM "Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\${APPNAME}" "DisplayName" "${APPNAME}"
    WriteRegStr HKLM "Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\${APPNAME}" "UninstallString" "$INSTDIR\\uninstall.exe"
    WriteRegStr HKLM "Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\${APPNAME}" "InstallLocation" "$INSTDIR"
    WriteRegStr HKLM "Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\${APPNAME}" "DisplayIcon" "$INSTDIR\\GlucoseConverter.exe"
    WriteRegStr HKLM "Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\${APPNAME}" "Publisher" "${COMPANYNAME}"
    WriteRegStr HKLM "Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\${APPNAME}" "HelpLink" "${HELPURL}"
    WriteRegStr HKLM "Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\${APPNAME}" "URLUpdateInfo" "${UPDATEURL}"
    WriteRegStr HKLM "Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\${APPNAME}" "URLInfoAbout" "${ABOUTURL}"
    WriteRegStr HKLM "Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\${APPNAME}" "DisplayVersion" "${VERSIONMAJOR}.${VERSIONMINOR}.${VERSIONBUILD}"
    WriteRegDWORD HKLM "Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\${APPNAME}" "VersionMajor" ${VERSIONMAJOR}
    WriteRegDWORD HKLM "Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\${APPNAME}" "VersionMinor" ${VERSIONMINOR}
    WriteRegDWORD HKLM "Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\${APPNAME}" "NoModify" 1
    WriteRegDWORD HKLM "Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\${APPNAME}" "NoRepair" 1
    
SectionEnd

Section "uninstall"
    Delete "$INSTDIR\\GlucoseConverter.exe"
    Delete "$INSTDIR\\uninstall.exe"
    
    Delete "$SMPROGRAMS\\${APPNAME}\\${APPNAME}.lnk"
    Delete "$SMPROGRAMS\\${APPNAME}\\Uninstall.lnk"
    RmDir "$SMPROGRAMS\\${APPNAME}"
    
    Delete "$DESKTOP\\${APPNAME}.lnk"
    
    DeleteRegKey HKLM "Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\${APPNAME}"
    
    RmDir $INSTDIR
SectionEnd
"""
    
    with open('installer.nsi', 'w') as f:
        f.write(nsis_script)
    
    print("✅ Created NSIS installer script (installer.nsi)")
    print("   To create installer, install NSIS and run: makensis installer.nsi")


def create_batch_launcher():
    """Create batch file for easy launching"""
    batch_content = """@echo off
title Glucose Converter

echo Starting Glucose Converter...

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo Python is not installed or not in PATH
    echo Please install Python from https://www.python.org/
    pause
    exit /b 1
)

REM Check and install requirements
echo Checking requirements...
pip install openpyxl tkinterdnd2 --quiet

REM Run the GUI
python glucose_converter_gui.py

if errorlevel 1 (
    echo.
    echo An error occurred. Press any key to exit...
    pause >nul
)
"""
    
    with open('run_glucose_converter.bat', 'w') as f:
        f.write(batch_content)
    
    print("✅ Created batch launcher (run_glucose_converter.bat)")


def main():
    """Main build process"""
    print("=" * 50)
    print("Glucose Converter - Build Executable")
    print("=" * 50)
    
    # Check OS
    if sys.platform != 'win32':
        print("⚠️  Warning: Building Windows executable on non-Windows system")
        print("   The executable will only work on Windows")
    
    # Check and install requirements
    check_requirements()
    
    # Build executable
    exe_path = build_executable()
    
    if exe_path:
        print("\n" + "=" * 50)
        print("✅ BUILD SUCCESSFUL!")
        print("=" * 50)
        print(f"\nExecutable location: {exe_path}")
        print("\nYou can now:")
        print("1. Run the executable directly")
        print("2. Copy it to any Windows computer")
        print("3. Create shortcuts as needed")
        
        # Create additional files
        create_batch_launcher()
        create_installer_script()
        
        print("\nAdditional files created:")
        print("- run_glucose_converter.bat (for running from source)")
        print("- installer.nsi (for creating professional installer)")
        
    else:
        print("\n❌ Build failed. Please check the error messages above.")
    
    # Cleanup option
    if Path('build').exists() or Path('__pycache__').exists():
        response = input("\nClean up build files? (y/n): ")
        if response.lower() == 'y':
            if Path('build').exists():
                shutil.rmtree('build')
            if Path('__pycache__').exists():
                shutil.rmtree('__pycache__')
            if Path('glucose_converter.spec').exists():
                os.remove('glucose_converter.spec')
            print("✅ Cleaned up build files")


if __name__ == '__main__':
    main()