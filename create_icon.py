"""
Pic2PPT Icon Generator
Create Windows ICO format icon using PIL
"""

from PIL import Image, ImageDraw


def create_pic2ppt_icon():
    """Generate Pic2PPT multi-size icon"""
    sizes = [16, 24, 32, 48, 64, 128, 256]
    images = []

    for size in sizes:
        img = create_icon_size(size)
        images.append(img)

    # Save as ICO file
    images[0].save(
        'pic2ppt.ico',
        format='ICO',
        sizes=[(s, s) for s in sizes],
        append_images=images[1:]
    )
    print("Icon generated: pic2ppt.ico")
    print(f"Sizes: {sizes}")


def create_icon_size(size):
    """Generate icon at specific size"""
    img = Image.new('RGBA', (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    # Colors (Orange theme)
    color_left = (249, 115, 22)      # Deep orange - Image
    color_right = (245, 158, 11)     # Amber - PPT
    bg_color = (255, 251, 235)       # Light yellow background

    # Background rounded rectangle
    padding = max(1, size // 32)
    corner = size // 5
    draw.rounded_rectangle(
        [padding, padding, size - padding, size - padding],
        radius=corner,
        fill=bg_color
    )

    # Calculate block size and positions
    gap = max(1, size // 32)
    block_size = (size - 2 * padding - 3 * gap) // 2

    # Block positions
    positions = [
        (padding + gap, padding + gap),  # Top-left - Image
        (padding + 2 * gap + block_size, padding + gap),  # Top-right - PPT
        (padding + gap, padding + 2 * gap + block_size),  # Bottom-left - Image
        (padding + 2 * gap + block_size, padding + 2 * gap + block_size),  # Bottom-right - PPT
    ]

    colors = [color_left, color_right, color_left, color_right]
    block_radius = max(1, block_size // 5)

    for (x, y), color in zip(positions, colors):
        draw.rounded_rectangle(
            [x, y, x + block_size, y + block_size],
            radius=block_radius,
            fill=color
        )

    return img


if __name__ == '__main__':
    print("Generating Pic2PPT icon...")
    create_pic2ppt_icon()
    print("\nOutput: pic2ppt.ico")
    print("PyInstaller: icon='pic2ppt.ico'")
