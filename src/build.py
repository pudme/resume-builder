import PyInstaller.__main__
import os
import shutil
import platform
import sys

def build_executable():
    # Clean up previous builds
    if os.path.exists('build'):
        shutil.rmtree('build')
    if os.path.exists('dist'):
        shutil.rmtree('dist')
    
    # Base arguments for PyInstaller
    base_args = [
        'resume_builder_gui.py',  # Main script
        '--name=ResumeBuilder',   # Name of the executable
        '--onefile',             # Create a single executable file
        '--windowed',            # Don't show console window
        '--clean',               # Clean PyInstaller cache
        '--noconfirm',           # Replace existing build without asking
    ]
    
    # Add OS-specific arguments
    if platform.system() == 'Windows':
        base_args.extend([
            '--add-data=README.md;.',  # Windows path separator
            '--icon=icon.ico',         # Windows icon
        ])
    else:  # MacOS
        base_args.extend([
            '--add-data=README.md:.',  # Unix path separator
            '--icon=icon.icns',        # MacOS icon
        ])
    
    # Run PyInstaller
    PyInstaller.__main__.run(base_args)
    
    # Print success message with OS-specific information
    if platform.system() == 'Windows':
        print("\nBuild complete! The executable is in the 'dist' folder.")
        print("You can distribute the 'ResumeBuilder.exe' file to users.")
    else:
        print("\nBuild complete! The executable is in the 'dist' folder.")
        print("You can distribute the 'ResumeBuilder' file to users.")
        print("\nNote: For MacOS, you may want to create a .app bundle.")
        print("To create a .app bundle, run: python create_app_bundle.py")

def create_app_bundle():
    """Create a MacOS .app bundle"""
    if platform.system() != 'Darwin':  # Only run on MacOS
        print("This script can only be run on MacOS.")
        return
    
    # Create the .app bundle
    os.system('pyinstaller --name "ResumeBuilder" --windowed --onefile --add-data "README.md:." resume_builder_gui.py')
    
    # Create the .app structure
    app_name = "ResumeBuilder.app"
    if os.path.exists(app_name):
        shutil.rmtree(app_name)
    
    os.makedirs(f"{app_name}/Contents/MacOS", exist_ok=True)
    os.makedirs(f"{app_name}/Contents/Resources", exist_ok=True)
    
    # Move the executable
    shutil.move("dist/ResumeBuilder", f"{app_name}/Contents/MacOS/")
    
    # Create Info.plist
    info_plist = f"""<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>CFBundleExecutable</key>
    <string>ResumeBuilder</string>
    <key>CFBundleIdentifier</key>
    <string>com.resumebuilder.app</string>
    <key>CFBundleName</key>
    <string>ResumeBuilder</string>
    <key>CFBundlePackageType</key>
    <string>APPL</string>
    <key>CFBundleShortVersionString</key>
    <string>1.0</string>
    <key>LSMinimumSystemVersion</key>
    <string>10.13.0</string>
    <key>NSHighResolutionCapable</key>
    <true/>
    <key>NSIconFile</key>
    <string>icon.icns</string>
</dict>
</plist>"""
    
    with open(f"{app_name}/Contents/Info.plist", "w") as f:
        f.write(info_plist)
    
    print("\nApp bundle created successfully!")
    print(f"You can find the {app_name} in the current directory.")

if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--app-bundle":
        create_app_bundle()
    else:
        build_executable() 