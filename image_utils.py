import os
from PIL import Image, ImageOps
from logger import log

def resize_image_fixed(input_path, output_path, width, height):
    """Resize to fixed width/height, auto-correct orientation, and save as JPEG."""
    try:
        img = Image.open(input_path)

        # Auto-rotate based on EXIF orientation
        img = ImageOps.exif_transpose(img)

        # Convert to RGB always (JPEG requirement)
        if img.mode not in ("RGB", "L"):
            img = img.convert("RGB")

        # Resize
        img = img.resize((width, height), Image.LANCZOS)

        # Ensure parent folder exists
        os.makedirs(os.path.dirname(output_path), exist_ok=True)

        # Save
        img.save(output_path, format="JPEG")
        img.close()

    except Exception as e:
        raise RuntimeError(f"Resize error for {input_path}: {e}")
