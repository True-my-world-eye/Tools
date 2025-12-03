import os
from PIL import Image

SRC = r'f:\AAAAclass\python\翻译练习\软件图标.png'
OUT = r'f:\AAAAclass\python\翻译练习\assets\app.ico'

def main():
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    img = Image.open(SRC).convert('RGBA')
    sizes = [(16,16),(32,32),(48,48),(64,64),(128,128),(256,256)]
    img.save(OUT, format='ICO', sizes=sizes)
    print('ICO written:', OUT)

if __name__ == '__main__':
    main()

