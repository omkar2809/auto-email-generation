from PIL import Image
from PIL import ImageFont
from PIL import ImageDraw


def make_certificate(name):
    img = Image.open('certificate.jpg')
    draw = ImageDraw.Draw(img)
    selectFont = ImageFont.truetype('LHANDW.ttf', size=64)
    width, height = img.size
    w, h = draw.textsize(name, selectFont)
    draw.text(((width - w) / 2, (height - h) / 2), name, '#7daef7', selectFont)
    img.save('certi.pdf', 'PDF')
    img.save('certi.png', 'png')
