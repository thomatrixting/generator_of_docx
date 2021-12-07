from traceback import print_tb
from docx import Document
import pandas as pd

def run():
    df = pd.DataFrame({'nombre':['thomas','jeronimo'],'cedula':[2356,2567]})
    df = df.rename(columns={'nombre':'name','cedula':'id'})
    
    for index,row in df.iterrows():
        doc = Document('templeates/certificate_templeate.docx')
        text_to_change = doc.paragraphs[16].text 
        text_to_change = text_to_change.format(row['name'],row['id'])
        doc.paragraphs[16].text  = text_to_change 
        doc.save('files_words/demo2.docx')
        print(doc.paragraphs[16].text)

if __name__ == '__main__':
    run()