import win32com.client as wc
from PIL import Image
import os, random

pw = wc.Dispatch('PowerPoint.Application')

presentation = pw.Presentations.Add()

# slides are 1280 px (w) by 720 px (h)

chrisPictures = 333
chrisHolder = 0
sarahPictures = 213
sarahHolder = 0

for x in range(chrisPictures + sarahPictures):
    choice = random.randint(0,1) # if 1 then chris, otherwise sarah
    try:
        if choice == 1 and chrisHolder < chrisPictures:
            chrisHolder += 1
            slide = presentation.Slides.Add(x + 1, 1)
            filePath = 'C:\\Users\\jeremyshank\\OneDrive - BMSS\\Chris and Sarah Pictures\\Chris\\' + str(chrisHolder - 1) + '.png'
            img = Image.open(filePath)
            widthPosition = img.width / 2
            heightPosition = img.height / 2
            picture = slide.Shapes.AddPicture(FileName=filePath, LinkToFile=False, SaveWithDocument=True, Left=100, Top=0, Width=-1, Height=-1)
            picture.Left = (pw.ActivePresentation.PageSetup.SlideWidth - picture.Width) / 2
            picture.Top = (pw.ActivePresentation.PageSetup.SlideHeight - picture.Height) / 2

            presentation.Slides[x].SlideShowTransition.AdvanceOnTime = True
            presentation.Slides[x].SlideShowTransition.AdvanceTime = 3
        
        if choice == 0 and sarahHolder < sarahPictures:
            sarahHolder += 1
            slide = presentation.Slides.Add(x + 1, 1)
            filePath = 'C:\\Users\\jeremyshank\\OneDrive - BMSS\\Chris and Sarah Pictures\\Sarah\\Final\\' + str(sarahHolder - 1) + '.png'
            img = Image.open(filePath)
            widthPosition = img.width / 2
            heightPosition = img.height / 2
            picture = slide.Shapes.AddPicture(FileName=filePath, LinkToFile=False, SaveWithDocument=True, Left=100, Top=0, Width=-1, Height=-1)
            picture.Left = (pw.ActivePresentation.PageSetup.SlideWidth - picture.Width) / 2
            picture.Top = (pw.ActivePresentation.PageSetup.SlideHeight - picture.Height) / 2

            presentation.Slides[x].SlideShowTransition.AdvanceOnTime = True
            presentation.Slides[x].SlideShowTransition.AdvanceTime = 4

    except:
        continue

print(chrisHolder)
print(sarahHolder)