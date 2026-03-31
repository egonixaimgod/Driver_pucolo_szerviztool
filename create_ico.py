from PIL import Image, ImageDraw

# Egy 256x256-as átlátszó alapon kék pajzs / fogaskerék szerű egyszerű ikon rajzolása
img = Image.new('RGBA', (256, 256), (0, 0, 0, 0))
draw = ImageDraw.Draw(img)

# Külső pajzs / körkörös forma vastag fekete kerettel és kék belsővel
draw.ellipse((20, 20, 236, 236), fill=(30, 144, 255), outline=(10, 50, 150), width=15)

# Belső ábrázolás egy X-el vagy egy sima villáskulcs "fejjel", ami egy tools / kezelő appra utal
draw.rectangle((100, 60, 156, 196), fill=(255, 255, 255))
draw.rectangle((60, 100, 196, 156), fill=(255, 255, 255))

img.save('icon.ico')
print("Ikon generálva a következő néven: icon.ico")
