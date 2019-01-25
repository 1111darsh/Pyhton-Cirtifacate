import xlrd
from PIL import Image
from PIL import ImageDraw
from PIL import ImageFont

img = Image.open("sample_in.jpg")


# Give the location of the file
loc = "Detail.xlsx"

# To open Workbook
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

# For row 0 and column 0
for i in range(1,sheet.nrows):
    img = Image.open("ProjectCompletion.jpg")
    fnt = ImageFont.truetype('Font/OLDENGL.ttf', 20)
    name = sheet.cell_value(i, 0)
    date = sheet.cell_value(i, 1)
    Project_title = sheet.cell_value(i,2)
    print(date)
    draw = ImageDraw.Draw(img)
    # font = ImageFont.truetype(<font-file>, <font-size>
    draw.text((100, 361), name, font=fnt, fill=(0, 0, 0))
    draw.text((830, 361), date, font=fnt, fill=(0, 0, 0))
    draw.text((340, 418), Project_title, font=fnt, fill=(0, 0, 0))
    img.save('out/sample-out'+str(i)+'.jpg')