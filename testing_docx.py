from docx import Document

doc = Document('word_test.docx')


print(doc.paragraphs[0].text)

