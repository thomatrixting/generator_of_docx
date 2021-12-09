from docx import Document
import pandas as pd

def filter_df (df):

    df.columns = [''] * len(df.columns)
    colums_list = df.iloc[[2]].values.tolist()[0]
    df.columns  = [value.replace(' ','').lower() for value in colums_list] #set colums

    df = df[(df.iloc[:, 0] == df.iloc[:, 0])] #eliminate the nan
    df = df.drop([2]) #eliminate row with colums names
    return df 

def main(relative_path_excel,path_to_save_words):
    df = pd.read_excel(relative_path_excel)
    df = filter_df(df)

    df = df.rename(columns={'nonbre':'name','cedula':'id'})

    for index,row in df.iterrows():
        doc = Document('templeates/certificate_templeate.docx')
        text_to_change = doc.paragraphs[16].text 
        text_to_change = text_to_change.format(row['name'],row['id'])
        doc.paragraphs[16].text  = text_to_change 
        name = row['name'].replace(' ','_')
        doc.save(f'{path_to_save_words}/{name}_certificacion.docx')
        print(doc.paragraphs[16].text)

def run():
        main('exel/certifications_test.xlsx','files_words')
if __name__ == '__main__':
    run()