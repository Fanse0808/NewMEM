from PIL import Image
import os

# Compress EmailBody.jpg to a smaller size
input_path = 'static/EmailBody.jpg'
output_path = 'static/EmailBody_compressed.jpg'

if os.path.exists(input_path):
    with Image.open(input_path) as img:
        # Resize to a reasonable size for email
        max_width = 800
        ratio = max_width / img.width
        new_height = int(img.height * ratio)
        
        # Resize the image
        img_resized = img.resize((max_width, new_height), Image.Resampling.LANCZOS)
        
        # Save with compression
        img_resized.save(output_path, 'JPEG', quality=70, optimize=True)
        
    print(f"Compressed image saved as {output_path}")
    print(f"Original size: {os.path.getsize(input_path) / 1024:.1f} KB")
    print(f"Compressed size: {os.path.getsize(output_path) / 1024:.1f} KB")
else:
    print("EmailBody.jpg not found") 