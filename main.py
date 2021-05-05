print('hello')
# hi cam
from pathlib import Path
from docx2python import docx2python

word_doc = Path("J:\C04100_C04199\C04147_5_Semple_St_Porirua_Wellington\C04147100_Due_Diligence\007_Work\Reporting\DSI\C04147100R001_FINAL.docx")


print(word_doc)
word = docx2python('C04147100R001_FINAL.docx')
text = word.body
exec = text[2]