import os
from PIL import Image
import platform

def create_ico(image_path, output_path, sizes=[16, 32, 48, 64, 128, 256]):
    """Create a Windows .ico file with multiple sizes"""
    img = Image.open(image_path)
    img.save(output_path, format='ICO', sizes=[(size, size) for size in sizes])

def create_icns(image_path, output_path):
    """Create a macOS .icns file"""
    if platform.system() != 'Darwin':
        print("Warning: .icns files can only be created on macOS")
        return False
    
    # Create iconset directory
    iconset_name = "resume_builder.iconset"
    if not os.path.exists(iconset_name):
        os.makedirs(iconset_name)
    
    # Generate icons of different sizes
    sizes = [16, 32, 64, 128, 256, 512, 1024]
    img = Image.open(image_path)
    
    for size in sizes:
        resized = img.resize((size, size), Image.Resampling.LANCZOS)
        resized.save(f"{iconset_name}/icon_{size}x{size}.png")
        if size <= 512:  # Also create @2x versions
            resized.save(f"{iconset_name}/icon_{size//2}x{size//2}@2x.png")
    
    # Convert iconset to .icns using iconutil (macOS only)
    os.system(f"iconutil -c icns {iconset_name}")
    
    # Move the .icns file to the output location
    if os.path.exists("resume_builder.icns"):
        os.rename("resume_builder.icns", output_path)
    
    # Clean up
    import shutil
    shutil.rmtree(iconset_name)
    return True

def main():
    # Create icons directory if it doesn't exist
    icons_dir = os.path.join("icons")
    if not os.path.exists(icons_dir):
        os.makedirs(icons_dir)
    
    # Base image path (you'll need to provide this)
    base_image = "base_icon.png"  # Replace with your base image
    
    if not os.path.exists(base_image):
        print(f"Error: Base image '{base_image}' not found!")
        print("Please provide a square PNG image (at least 1024x1024 pixels) named 'base_icon.png'")
        return
    
    # Generate Windows icon
    ico_path = os.path.join(icons_dir, "resume_builder.ico")
    create_ico(base_image, ico_path)
    print(f"Created Windows icon: {ico_path}")
    
    # Generate macOS icon
    icns_path = os.path.join(icons_dir, "resume_builder.icns")
    if create_icns(base_image, icns_path):
        print(f"Created macOS icon: {icns_path}")
    else:
        print("Skipped macOS icon creation (requires macOS)")

if __name__ == "__main__":
    main() 