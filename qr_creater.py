from tkinter import Scale
import pyqrcode
from pyqrcode import QRCode
import re

n=int(input("How many student you want to make QR Code : "))

j = r"20SOECE110[0-9][0-9]"
for i in range(n):
    print("Name : ",end='')
    s1 = input()
    while True:
        print('+++++++ Note: Enrollment Must be start with "20SOECE11__"  _= 0 to 9')
        print("Enrollment No : ",end='')
        s2 = input()
        if (re.fullmatch(j,s2)):
            break
    t = pyqrcode.create(s2+s1)
    t.svg(f'{s2}.svg', scale = 6)