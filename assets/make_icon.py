from PIL import Image

img = Image.open("assets/icon.png").convert("RGBA")
img.save("assets/icon.ico", format="ICO", sizes=[(16, 16), (32, 32), (48, 48), (256, 256)])