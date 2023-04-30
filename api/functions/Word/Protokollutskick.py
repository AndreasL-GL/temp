from __future__ import print_function
import mailmerge
import docx
from PIL import Image
import base64
import io
import json

if __name__ == '__main__':
    import get_template
    from add_image_to_zipfile import add_icon_to_word_file, autoorient
    import create_json
else:
    import functions.Word.get_template as get_template
    import functions.Word.create_json as create_json
    from functions.Word.add_image_to_zipfile import add_icon_to_word_file, autoorient


def add_template_data(file, tbl):
    doc = mailmerge.MailMerge(file)
    content= tbl["content"]
    info = {k:v for k,v in tbl.items() if k!="content"}
    
    print(doc.get_merge_fields())
    doc.merge(**info)
    doc.merge_rows('Område',content)
    
    
    
    for item in content:
        item["Område2"] = item["Område"]
    file = io.BytesIO()
    doc.write(file)
    file.seek(0)
    return file
def add_images_to_doc(file,tbl):
    doc = docx.Document(file)
    
    doc.add_page_break()
    doc.add_heading('Bilder och Kommentarer', 0)
    table = doc.add_table(rows=1, cols=3)
    table.style = 'Plain Table 5'
    for item in tbl["images"]:
        row = table.add_row().cells
        Adress = item["Område"]
        Bildkommentar = item["Bildkommentar"]
        Bild = item["Bild"]
        row[0].text = Adress
        row[1].text = Bildkommentar

        img = Image.open(Bild)
        img =autoorient(img)
        img = img.resize((200,250))
        file=io.BytesIO()
        img.save(file, format='JPEG')
        file.seek(0)
        row[2].add_paragraph().add_run().add_picture(file)
    return doc

def change_icon_in_header(doc):
    img = Image.open("bild2.jpg")
    img = img.resize((100,100))
    fb = io.BytesIO()
    img.save(fb, 'PNG')
    fb.seek(0)
    docbyte = io.BytesIO()
    doc.save(docbyte)
    docbyte.seek(0)
    doc = add_icon_to_word_file(docbyte.read(), icon_file=fb.read())
    doc = docx.Document(doc)
    return doc

def run_functions(js):
    js = create_json.create_json_for_word_functions(js)
    file= download_template_file()
    file = add_template_data(file, js)
    doc = add_images_to_doc(file, js)
    #doc = change_icon_in_header(doc)
    doc.save('gfg.docx')
    file_content = io.BytesIO()
    doc.save(file_content)
    file_content.seek(0)
    file_content=base64.b64encode(file_content.getvalue()).decode('utf-8')
    return {"content":file_content, "filename": "filename.docx"}

def download_template_file():
    l = get_template.Download_template("","word_template.docx")
    print(type(l.content))
    file_content = io.BytesIO(l.content)
    file_content.seek(0)
    return file_content
if __name__ == '__main__':
    
    with open('ActionOutputs (2).json', 'r', encoding="utf-8") as f:
        js = json.load(f)
    run_functions(js)