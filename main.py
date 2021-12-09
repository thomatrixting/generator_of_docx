from traceback import print_tb
from docx import Document
import pandas as pd

def filter_df (df):

    df.columns = [''] * len(df.columns)
    colums_list = df.iloc[[2]].values.tolist()[0]
    df.columns  = [value.replace(' ','').lower() for value in colums_list] #set colims
    df = df.drop([0, 1, 2])#eliminate 2 firts rows that creates probelms ans the 3 with the colums names
    return df

def run():
    df = pd.read_excel('exel/certifications_test.xlsx')
    df = filter_df(df)
    
    df = df.rename(columns={'nonbre':'name','cedula':'id'})

    for index,row in df.iterrows():
        doc = Document('templeates/certificate_templeate.docx')
        text_to_change = doc.paragraphs[16].text 
        text_to_change = text_to_change.format(row['name'],row['id'])
        doc.paragraphs[16].text  = text_to_change 
        doc.save('files_words/demo3.docx')
        print(doc.paragraphs[16].text)

if __name__ == '__main__':
    run()