from PIL import Image


def png_to_multi_ico(png_path: str, ico_path: str):
    """
    Convert a single PNG file into a Windows .ico containing
    multiple resolutions (16×16, 32×32, 48×48, 64×64, 128×128, 256×256).

    Args:
        png_path: Path to the source PNG (e.g. your treasure chest).
        ico_path: Path where to write the .ico file.
    """
    # Load source image
    img = Image.open(png_path).convert("RGBA")

    # Define the sizes we want in the ICO
    sizes = [(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]

    # Resize (lanczos for best quality)
    icons = [img.resize(s, Image.LANCZOS) for s in sizes]

    # Save as a multi-icon file
    # Pillow will pack all the specified sizes into the single .ico
    icons[0].save(ico_path, format="ICO", sizes=sizes)


png_to_multi_ico("resources/loot_box_icon.png", "resources/loot_box_icon.ico")
