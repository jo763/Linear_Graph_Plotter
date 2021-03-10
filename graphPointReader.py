print("Loading modules")
from os import environ
environ['PYGAME_HIDE_SUPPORT_PROMPT'] = "hide"
import pygame
import openpyxl 
from easygui import *
from time import sleep





def exInputOpenSave(rowInput, columnInput, cellInput, fileName):
    excel = openpyxl.load_workbook(fileName)
    excelSht = excel.active
    excelSht.cell (row = rowInput, column = columnInput).value = cellInput
    excel.save(fileName)
    pass

#exInputOpenSave(1,1,"reee", sourceFile)

def graph():
    '''
    Displays the graph on the window. This also essentially cleans the graph from previous markings
    '''
    screen.blit(pygame.transform.scale(graphImg, (size)),(graphX,graphY))
    
    
def drawTarget(x,y,colour):
    '''
    Draws a crosshair on the screen
    '''
    pygame.draw.rect(screen, colour, [x,y-1, 3000, 2])
    pygame.draw.rect(screen, colour, [x-1,y, 2, 3000])
    pygame.draw.rect(screen, colour, [ -2, -2, x+1,y+1], 2)
    
def message_display(text):
    '''
    Creates & displays the message saying the x and y values
    '''
    largeText = pygame.font.Font(None,30)
    TextSurf, TextRect = text_objects(text, largeText)
    TextRect.center = ((width/2),(18))
    screen.blit(TextSurf, TextRect)

def text_objects(text, font):
    '''
    Renders the text
    '''
    textSurface = font.render(text, True, blue)
    return textSurface, textSurface.get_rect()

def messageUpdate(xPos, yPos, roundX, roundY):
    '''
    Creates the display message
    '''
    message_display(f"X Axis: {round(xPos,roundX)}, Y Axis: {round(yPos,roundY)}")


# Name that will be on all the captions for the selection boxes
programName = "Joe Prollins' Graphical Reader"


msg = "Enter the name that you want your excel to be named. Don't include the file extension. \nNote: this will overwrite a file with the same name."
title = programName
sourceFile = []
sourceFile = enterbox(msg, title, sourceFile)
sourceFile = sourceFile + ".xlsx"



excel = openpyxl.Workbook()
#sourceFile = "test.xlsx"
excel.save(sourceFile)


# Intialises the pygame module
pygame.init()


msgbox("Select the graph you wish to upload.")

# Selection of the image file. Creates a pop up box, then creates a file browser
#msgbox("Please selecct the picture file of the graph you want to use")
imageFile = fileopenbox()
#imageFile = "GhoosData.PNG"

# Creation of a selection of screen sizes
screenSizeX = [640, 720, 800, 1024, 1152, 1280, 1366, 1500, 1680, 1920]
screenSizeY = [480, 400, 600, 768, 864, 720, 1024, 768, 900, 1050, 1080]
screenSizeChoices = []
for i in range(len(screenSizeX)):
    screenSizeChoices.append(str(screenSizeX[i]) + " x " + str(screenSizeY[i]))

# Asks the user which screen size they want to select via a multiple-choice pop up box
msg ="Please choose the size of the screen."
title = programName + ": "+ "Resolution Size Selection Box"
choices = screenSizeChoices
screenChoice = choicebox(msg, title, choices)

# Defines the size of the screen size
indexChoice = screenSizeChoices.index(screenChoice)
size = width, height = screenSizeX[indexChoice], screenSizeY[indexChoice]



msgbox("Select the origin point and then press the 'enter' key.\nThen click on a point that you know the value of (that's not on either axis) and press the 'enter' key. The further from the origin, the more accurate it will be.")

# Setting the size of the screen in pygame
screen = pygame.display.set_mode(size)

# Setting the caption of the program window
pygame.display.set_caption(programName)


# Loads in the graphical image
graphImg = pygame.image.load(imageFile)
# Settomg the coordinates for where the graph will be placed
graphX = 0
graphY = 0
# Setting up the colours 
red = (255,0,0)
blue = (0,0,255)
green = (0, 255, 0)



 

calibration = True
calibrateAxis = False
allowEvent = True
dataEntryRow = 2
exInputOpenSave(1,1,"x axis", sourceFile)
exInputOpenSave(1,2,"y axis", sourceFile)




running = True
graph()
while running:
    for event in pygame.event.get():
        if event.type == pygame.QUIT:
            pygame.display.quit()
            break
            running = False

    if calibration == True:
        if event.type == pygame.MOUSEBUTTONDOWN:
            graph()
            pos = pygame.mouse.get_pos()
            drawTarget(pos[0],pos[1],red)
            allowEvent = True
        if event.type == pygame.KEYDOWN and allowEvent == True:
            if event.key == pygame.K_RETURN:
                originX = pos[0]
                originY = pos[1]
                graph()
                drawTarget(pos[0],pos[1],green)
                calibrateAxis = True
                calibration = False
                allowEvent = False


    if calibrateAxis == True:
        if event.type == pygame.MOUSEBUTTONDOWN:
            graph()
            pos = pygame.mouse.get_pos()
            drawTarget(pos[0],pos[1],red)
            allowEvent = True
        if event.type == pygame.KEYDOWN and allowEvent == True:
            if event.key == pygame.K_RETURN:
                selectionX = pos[0]
                selectionY = pos[1]
                graph()
                drawTarget(pos[0],pos[1],green)
                pygame.display.update()
                print(pos[0],pos[1])
                calibrateAxis = False
                msg = "What is the value of the point you have selected?\nHow many decimal places do you want the results to round to?\nAfter submiting the values, clicking anywhere will yield the values, pressing the 'enter' key will log the values in an excel file." 
                title = programName
                fieldNames = ["x value", "y value", "x rounding", "y rounding"]
                fieldValues =[]
                fieldValues = multenterbox(msg,title, fieldNames)
                trueX = float(fieldValues[0])
                trueY = float(fieldValues[1])
                scaleX = (selectionX - originX) / trueX
                scaleY = (selectionY - originY) / trueY
                roundX = int(fieldValues[2])
                roundY = int(fieldValues[3])
                allowEvent = False
                sleep(.2)


    
    if calibration == False and calibrateAxis == False:
        if event.type == pygame.MOUSEBUTTONDOWN:
            graph()
            pos = pygame.mouse.get_pos()
            drawTarget(pos[0],pos[1],blue)
            xPos = ((pos[0]-originX)/scaleX)
            yPos = ((pos[1]-originY)/scaleY)
            print("\n\n===================================\n")
            print ("X Axis = ", round(xPos,roundX), "\nY Axis = ", round(yPos,roundY))
            messageUpdate(xPos, yPos, roundX, roundY)
            sleep(.2)
            allowEvent = True
            pass
        if event.type == pygame.KEYDOWN and allowEvent == True:
            if event.key == pygame.K_RETURN:
                graph()
                messageUpdate(xPos, yPos, roundX, roundY)
                print("data logged")
                drawTarget(pos[0],pos[1],green)
                pygame.display.update()
                exInputOpenSave(dataEntryRow, 1, round(xPos,roundX), sourceFile)
                exInputOpenSave(dataEntryRow, 2, round(yPos,roundX), sourceFile)
                dataEntryRow += 1
                allowEvent = False
                sleep(.2)

        try:
            drawTarget(pos[0], pos[1], blue)
        except:
            pass
            

    pygame.display.update()










