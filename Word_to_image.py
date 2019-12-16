import xlsxwriter
from PIL import Image

textfile = open('/Users/Shanyl/Documents/Python_Projects/PythonWordtoImage/PythonReferral.txt', 'r')
image = Image.open("/Users/Shanyl/Documents/Python_Projects/PythonWordtoImage/Bee.jpg", 'r')
bgimage = Image.open("/Users/Shanyl/Documents/Python_Projects/PythonWordtoImage/Beebackground.jpg", 'r')
pixval = list(image.getdata())
bgpixval = list(bgimage.getdata())

temp = []
charbank = []
for row in range(0, 252, 1):
    for col in range(0, 199, 1):
        index = row*199 + col
        textfile.seek(index)
        temp.append(textfile.read(1))
    charbank.append(temp)
    temp = []

temp = []
hexcolpix = []
for row in range(0, 252, 1):
    for col in range(0, 199, 1):
        index = row*199 + col
        temp.append(pixval[index])
    hexcolpix.append(temp)
    temp = []

temp = []
flatbghexcolpix = []
for row in range(0, 252, 1):
    for col in range(0, 199, 1):
        index = row*199 + col
        temp.append(bgpixval[index])
    flatbghexcolpix.append(temp)
    temp = []

temp = []
fonthexcolpix = []
for x in range(0, 252, 1):
    for y in range(0, 199, 1):
        temp_a = hex(hexcolpix[x][y][0])
        if len(temp_a) == 3:
            RR = str("0" + temp_a[2])
        if len(temp_a) == 4:
            RR = str(temp_a[2] + temp_a[3])
        temp_b = hex(hexcolpix[x][y][1])
        if len(temp_b) == 3:
            GG = str("0" + temp_b[2])
        if len(temp_b) == 4:
            GG = str(temp_b[2] + temp_b[3])
        temp_c = hex(hexcolpix[x][y][2])
        if len(temp_c) == 3:
            BB = str("0" + temp_c[2])
        if len(temp_c) == 4:
            BB = str(temp_c[2] + temp_c[3])
        temp.append(str("#" + RR + GG + BB))
    fonthexcolpix.append(temp)
    temp = []

temp = []
bghexcolpix = []
for x in range(0, 252, 1):
    for y in range(0, 199, 1):
        temp_a = hex(flatbghexcolpix[x][y][0])
        if len(temp_a) == 3:
            RR = str("0" + temp_a[2])
        if len(temp_a) == 4:
            RR = str(temp_a[2] + temp_a[3])
        temp_b = hex(flatbghexcolpix[x][y][1])
        if len(temp_b) == 3:
            GG = str("0" + temp_b[2])
        if len(temp_b) == 4:
            GG = str(temp_b[2] + temp_b[3])
        temp_c = hex(flatbghexcolpix[x][y][2])
        if len(temp_c) == 3:
            BB = str("0" + temp_c[2])
        if len(temp_c) == 4:
            BB = str(temp_c[2] + temp_c[3])
        temp.append(str("#" + RR + GG + BB))
    bghexcolpix.append(temp)
    temp = []

Xfile = xlsxwriter.Workbook('/Users/Shanyl/Documents/Python_Projects/PythonWordtoImage/Output.xlsx')
Xsheet = Xfile.add_worksheet()
for i in range(0, 252, 1):
    Xsheet.set_row(i, 18)
Xsheet.set_column(0, 198, 1.89)

temp = []
formatbank = []
for x in range(0, 252, 1):
    for y in range(0, 199, 1):
        temp.append(Xfile.add_format({'font_name': 'Impact', 'font_size': 16, 'font_color': fonthexcolpix[x][y], 'bg_color': bghexcolpix[x][y]}))
    formatbank.append(temp)
    temp = []

for row in range(0, 252, 1):
    for col in range(0, 199, 1):
            Xsheet.write_string(row, col, charbank[row][col], formatbank[row][col])

Xfile.close()
