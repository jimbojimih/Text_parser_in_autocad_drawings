import win32com.client
import os
from time import sleep

import pandas as pd
import pyautocad


drawings = os.listdir()
drawings_after_filter = list(filter(lambda x: '.dwg' in x, drawings))

def search_engine(drawings_list, need_text):
    drawing_info = dict()

    for drawing in drawings_list:
        text_list = list()

        #get patch
        folder = os.getcwd()
        filename = drawing
        drawing_file = os.path.join(folder, filename)

        #get access to autocad file management
        acad32 = win32com.client.dynamic.Dispatch("AutoCAD.Application")
        doc = acad32.Documents.Open(drawing_file)

        acad = pyautocad.Autocad()
        name = acad.doc.Name
        for text in acad.iter_objects('Text'):
            text = text.TextString
            if need_text in text:
                text_list.append(text)
        drawing_info[name] = text_list
        sleep(1)
    print(drawing_info)
    df = pd.DataFrame.from_dict(data=drawing_info, orient='index')
    df.to_excel('information.xlsx')

if __name__ == "__main__" :
    
    need_text = 'металл'
    search_engine(drawings_after_filter, need_text)