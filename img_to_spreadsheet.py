from PIL import Image
from tkinter import Tk, filedialog
from openpyxl import Workbook
from openpyxl.formatting.rule import ColorScaleRule
import webbrowser


def column_string(n):
    
    #takes a column integer value and converts it to an alphabetical column value used by excel
    string = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        string = chr(65 + remainder) + string
    return string

def conditionalFormatting(ws, redRow, greenRow, blueRow):
    
    #formats the red row by assigning the 0 to black and 255 as red
    ws.conditional_formatting.add(redRow, ColorScaleRule(start_type='num', start_value=0, start_color='000000', end_type='num', end_value=255, end_color='AA0000'))

    #formats the green row by assigning the 0 to black and 255 as green
    ws.conditional_formatting.add(greenRow, ColorScaleRule(start_type='num', start_value=0, start_color='000000', end_type='num', end_value=255, end_color='00AA00'))

    #formats the blue row by assigning the 0 to black and 255 as blue
    ws.conditional_formatting.add(blueRow, ColorScaleRule(start_type='num', start_value=0, start_color='000000', end_type='num', end_value=255, end_color='0000AA'))

def getImage():
    
    #opens up file explorer in the users "Pictures" directory and asks them to select a photo
    root = Tk().withdraw()
    filename =  filedialog.askopenfilename(initialdir = "~\Pictures",title = "Select file",filetypes = (("Image File","*.jpg *.png"),("all files","*.*")))
    try:
        #opens the image using PIL and reduces the resolution to better fit the spreadsheet
        img = Image.open(filename, "r")
        img.thumbnail((200, 200))
        width, height = img.size
    except:
        exit()

    #extracts the subpixel values into a 2D array
    pixelValues = list(img.getdata())
    renderSpreadsheet(pixelValues, width, height)

def renderSpreadsheet(pixelValues, width, height):
    #gets the string for the largest column in the spreadsheet eg: AB
    maxColumn = column_string(width)
    
    wb = Workbook()
    # grab the active worksheet
    ws = wb.active

    #goes through each subpixel and write 0-255 depending on colour value
    pixel = 0
    for row in range(height):
        for column in range(width):
            for subPixel in range(3):
                ws.cell(row=((((row + 1) * 3) - 2) + subPixel), column=column+1, value = pixelValues[pixel][subPixel])
            pixel += 1
        #creates a string for the current row of subpixels eg: A1:BU1
        redRow = "A" + str(row*3 + 1) + ":" + str(maxColumn) + str(row*3 + 1)
        greenRow = "A" + str(row*3 + 2) + ":" + str(maxColumn) + str(row*3 + 2)
        blueRow = "A" + str(row*3 + 3) + ":" + str(maxColumn) + str(row*3 + 3)

        conditionalFormatting(ws, redRow, greenRow, blueRow)

    # Save the file
    wb.save("Image.xlsx")
    webbrowser.open("Image.xlsx")

getImage()