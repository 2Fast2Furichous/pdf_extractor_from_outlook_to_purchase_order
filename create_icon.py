"""
Create custom icon for PDF Extractor application
"""
from PIL import Image, ImageDraw, ImageFont
import os

def create_icon():
    """Create a professional PDF extraction icon"""

    # Create multiple sizes for ICO format
    sizes = [16, 32, 48, 64, 128, 256]
    images = []

    for size in sizes:
        # Create new image with transparent background
        img = Image.new('RGBA', (size, size), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)

        # Calculate proportions
        margin = size // 8
        doc_width = size - (2 * margin)
        doc_height = int(doc_width * 1.3)

        # Center the document
        doc_x = margin
        doc_y = (size - doc_height) // 2

        # Draw PDF document shape (rectangle with folded corner)
        fold_size = doc_width // 4

        # Main document rectangle
        points = [
            (doc_x, doc_y),  # Top left
            (doc_x + doc_width - fold_size, doc_y),  # Top right (before fold)
            (doc_x + doc_width, doc_y + fold_size),  # Fold corner
            (doc_x + doc_width, doc_y + doc_height),  # Bottom right
            (doc_x, doc_y + doc_height),  # Bottom left
        ]

        # Draw shadow (offset slightly)
        shadow_offset = max(2, size // 64)
        shadow_points = [(x + shadow_offset, y + shadow_offset) for x, y in points]
        draw.polygon(shadow_points, fill=(0, 0, 0, 80))

        # Draw document with gradient-like effect (red PDF color)
        draw.polygon(points, fill=(220, 50, 50, 255), outline=(180, 30, 30, 255), width=max(1, size // 64))

        # Draw folded corner
        fold_points = [
            (doc_x + doc_width - fold_size, doc_y),
            (doc_x + doc_width, doc_y + fold_size),
            (doc_x + doc_width - fold_size, doc_y + fold_size),
        ]
        draw.polygon(fold_points, fill=(180, 30, 30, 255), outline=(180, 30, 30, 255))

        # Draw "PDF" text if size is large enough
        if size >= 48:
            try:
                # Try to use a bold font
                font_size = size // 6
                try:
                    font = ImageFont.truetype("arialbd.ttf", font_size)
                except:
                    try:
                        font = ImageFont.truetype("arial.ttf", font_size)
                    except:
                        font = ImageFont.load_default()

                text = "PDF"
                # Get text bounding box
                bbox = draw.textbbox((0, 0), text, font=font)
                text_width = bbox[2] - bbox[0]
                text_height = bbox[3] - bbox[1]

                text_x = doc_x + (doc_width - text_width) // 2
                text_y = doc_y + (doc_height - text_height) // 2 - size // 16

                # Draw text with slight shadow
                draw.text((text_x + 1, text_y + 1), text, fill=(100, 0, 0, 200), font=font)
                draw.text((text_x, text_y), text, fill=(255, 255, 255, 255), font=font)
            except:
                pass

        # Draw extraction arrow if size is large enough
        if size >= 64:
            arrow_y = doc_y + doc_height - doc_height // 4
            arrow_start_x = doc_x + doc_width // 4
            arrow_end_x = doc_x + (3 * doc_width) // 4
            arrow_size = size // 16

            # Arrow shaft
            draw.line([(arrow_start_x, arrow_y), (arrow_end_x, arrow_y)],
                     fill=(255, 255, 255, 255), width=max(2, size // 48))

            # Arrow head
            arrow_head = [
                (arrow_end_x, arrow_y),
                (arrow_end_x - arrow_size, arrow_y - arrow_size),
                (arrow_end_x - arrow_size, arrow_y + arrow_size),
            ]
            draw.polygon(arrow_head, fill=(255, 255, 255, 255))

        images.append(img)

    # Save as PNG (largest size for preview)
    images[-1].save('icon.png', 'PNG')
    print("Created icon.png")

    # Save as ICO (multi-resolution)
    images[0].save('icon.ico', format='ICO', sizes=[(s, s) for s in sizes])
    print("Created icon.ico")

    print(f"Icon files created successfully!")
    print(f"  - icon.png ({sizes[-1]}x{sizes[-1]})")
    print(f"  - icon.ico (multi-size: {', '.join(map(str, sizes))})")

if __name__ == "__main__":
    create_icon()
