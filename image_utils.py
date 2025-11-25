import os
from PIL import Image
from logger import log

def resize_image_fixed(input_path, output_path, width, height):
    """Resize to fixed width/height and ensure JPEG-compatible RGB mode."""
    try:
        img = Image.open(input_path)

        # Convert RGBA → RGB (JPEG does NOT support transparency)
        if img.mode in ("RGBA", "P"):
            img = img.convert("RGB")

        img = img.resize((width, height), Image.LANCZOS)
        # Ensure parent exists
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        # Save as JPEG for consistent handling
        img.save(output_path, format="JPEG")
        img.close()
    except Exception as e:
        raise RuntimeError(f"Resize error for {input_path}: {e}")
