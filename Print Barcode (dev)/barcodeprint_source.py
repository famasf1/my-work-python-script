from barcode import Code128, Code39
from barcode.writer import ImageWriter

def barcode_generator():
    word = input('')
    if word:
        with open(rf"D:\Workstuff\my-work-python-script\Print Barcode (dev)\result\{word}.png", "wb+") as files:
            Code39(word, writer=ImageWriter()).write(files)

if __name__ in '__main__':
    barcode_generator()