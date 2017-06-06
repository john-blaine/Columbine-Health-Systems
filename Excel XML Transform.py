try:
    # for Python2
    from Tkinter import *   ## notice capitalized T in Tkinter 
except ImportError:
    # for Python3
    from tkinter import *   ## notice lowercase 't' in tkinter here

import os
from xml.etree.ElementTree import Element, SubElement, ElementTree, parse
from tkinter import filedialog

Tk().withdraw() # we don't want a full GUI, so keep the root window from appearing
filename = filedialog.askopenfilename(title="Open File",
initialdir=("G:\\"))
# show an "Open" dialog box and return the path to the selected file
head, tail = (os.path.split(filename))

tail1, tail2 = tail.split(".")
tail1 = tail1 + "_FINAL."
tail = tail1 + tail2

tree = ElementTree(file=filename)
elem = tree.getroot()

a = elem[0][0]
b = elem[0][1]
c = elem[0][2]
d = elem[0][3]

for node in elem:
        node.remove(node[0])
        node.remove(node[0])
        node.remove(node[0])
        node.remove(node[0])

elem.insert(0, d)
elem.insert(0, c)
elem.insert(0, b)
elem.insert(0, a)

os.chdir(head)

tree.write(tail)
