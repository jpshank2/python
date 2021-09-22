# from pdf2image import convert_from_path
import os

# path = r'C:\Users\jeremyshank\desktop\pdfs'

# with os.scandir(path) as folder:
#     for file in folder:
#         print(os.path.abspath(file))
#         pages = convert_from_path(os.path.abspath(file), 500, poppler_path=r'C:\Python39\poppler-21.09.0\Library\bin')
#         for page in pages:
#             page.save('C:\\Users\\jeremyshank\\OneDrive - BMSS\\Chris and Sarah Pictures\\Sarah\\' + file.name + '.png', 'PNG')

pic_path = 'C:\\Users\\jeremyshank\\OneDrive - BMSS\\Chris and Sarah Pictures\\Sarah\\'

with os.scandir(pic_path) as folder:
    for count, file in enumerate(folder):
        if not os.path.isdir(os.path.abspath(file)):
            print(count)
            print(os.path.abspath(file))
            os.rename(os.path.abspath(file), os.path.join(pic_path, 'Fixed\\', str(count) + '.png'))
    # dst = str(count) + '.png'
    # src = str('C:\\Users\\jeremyshank\\OneDrive - BMSS\\Chris and Sarah Pictures\\Sarah\\') + filename
    # dst = str('C:\\Users\\jeremyshank\\OneDrive - BMSS\\Chris and Sarah Pictures\\Sarah\\') + dst

    # os.rename(src, dst)