from docx import Document
from config import RESUME_KEY
import os, csv, sys
from collections import OrderedDict

def docx2list(fn):
    text = []
    new_text = []
    doc = Document(fn)

    for t in doc.tables:
        for row in range(len(t.rows)):
            for column in range(len(t.columns)):
                text.append(t.cell(row, column).text)

    for item in text:
        if item not in new_text:
            new_text.append(item)

    return new_text

def list2dict(li):
    di = OrderedDict()
    value = None
    for i in range(len(li)):
        data = li.pop()
        if data in RESUME_KEY and value is None:
            di[data] = None
        elif data in RESUME_KEY and value:
            di[data] = value
            value = None
        elif data:
            value = data

    ret_di = OrderedDict()
    for key in reversed(di):
        ret_di[key] = di[key]
    return ret_di

def dict2csv(di, fn, new_fn=None):

    if new_fn is None:
        new_fn = os.path.splitext(os.path.basename(fn))[0] + '.csv'
    else:
        new_fn = fn

    with open(new_fn, 'w', encoding='utf8') as fp:
        writer = csv.DictWriter(fp, list(di.keys()))
        writer.writeheader()
        writer.writerow(di)

if __name__ == '__main__':

    """
    usage: 
    1, Run the command: python docx2csv.py your-dir-name
        it will convert all of the *.docx file to *.csv in the specified dir
    2, Run the python program directly, it will convert all of the *.docx file to *.csv
        in the same folder.
    
    Please note, it only have the effort on specified format of resume. 
        For other docx file, the output will be meaningless
    """

    if os.path.exists(str(sys.argv[1:2])):
        dir = sys.argv[1]
    else:
        dir = os.path.dirname(os.path.abspath(__file__))

    for file in os.listdir(dir):
        file_type = file.split('.')[-1]
        fn =  os.path.abspath(file)
        if 'docx' in file_type:
            data = docx2list(fn)
            di = list2dict(data)
            dict2csv(di, fn)