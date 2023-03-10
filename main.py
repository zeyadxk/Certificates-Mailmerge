from PIL import Image
from PIL import ImageDraw
from PIL import ImageFont
from openpyxl import load_workbook
import arabic_reshaper
from bidi.algorithm import get_display

# Load the workbook
wb = load_workbook('6.xlsx')

# Get the active sheet
ws = wb.active

# Create a list of the names
names = [cell.value for cell in ws['A']]
files = [cell.value for cell in ws['C']]

# Load the image
img = Image.open('Picture6.jpg')
W,H = img.size

# Iterate through the names
for name, filename in zip(names, files):
    
    #filename = name

    name = arabic_reshaper.reshape(name)
    name = get_display(name)
    # Generate a new image
    new_img = img.copy()
    W,H = new_img.size

    # Write the name on the image
    
    draw = ImageDraw.Draw(new_img)
    font = ImageFont.truetype('Arial Unicode', 35)
    _, _, w, h = draw.textbbox((0, 0), name, font=font)
    draw.text(((W-w)/2, 285), name, font=font, stroke_width=1, fill=(53, 87, 80))

    # Save the new image as a PDF
    new_img.save(f'6/{filename}.pdf')