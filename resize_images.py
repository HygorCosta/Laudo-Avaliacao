from PIL import Image,ExifTags, ImageOps 
from tkinter import Tk, filedialog
import os 

root = Tk()
root.withdraw()
root.attributes('-topmost', True)
folder = filedialog.askdirectory()

print('Mudando o formato das imagens...')
image_prefix = input("prefixo: ")

for count, picture in enumerate(os.listdir(folder)):
    file = f"{folder}/{picture}"
    img = Image.open(file)
    fixed_image = ImageOps.exif_transpose(img)
    fixed_image = fixed_image.resize((400, 600), Image.ANTIALIAS)
    fixed_image.save(folder + '/' + image_prefix + '_' + f"{count:02d}" + '.jpg', 'JPEG')
