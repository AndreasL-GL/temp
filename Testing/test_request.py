import requests
from PIL import Image,ExifTags
import io, json
import os
URL = "http://51.124.105.224/api/" #51.124.105.224
#URL = "http://localhost/api/"
API_KEY = "ABCDEFG"
def test_image_resize():
    url = URL+f'image_resizer?filename=filnamn.jpg&API_KEY={API_KEY}&height=500&width=500'
    files = {'media': open(os.path.join(os.path.dirname(__file__),'test.jpg'), 'rb')}
    imm = Image.open(os.path.join(os.path.dirname(__file__),'test.jpg'))
    im = requests.post(url, files=files)
    try:
        img = Image.open(io.BytesIO(im.content))
        img.show()
        imm.show()
    except Exception as e:
        print(im.text)

def test_sharepoint():
    url = URL+'flows/get_sharepoint_columns?API_KEY='+API_KEY
    with open(os.path.join(os.path.dirname(__file__),'sharepointlog.json'), encoding="utf-8") as f:
        js = json.load(f)

    rs = requests.post(url,json=js)
    print(rs.content)    

def test_sharepointv2():
    site="GLMalmAB-EgenkontrollerVellingebostder"
    list_ = 'MKB Egenkontroll Periodiska 2023'
    # rs = requests.get("http://51.124.105.224/api/flows/get_sharepoint_columns_v2?listt=MKB%20Egenkontroll%20Periodiska%202023&site=GLMalmAB-EgenkontrollerVellingebostder&API_KEY=ABCDEFG")
    rs = requests.get(f'{URL}flows/get_sharepoint_columns_v2?list=MKB Egenkontroll Periodiska 2023&site=GLMalmAB-EgenkontrollerVellingebostder&API_KEY='+API_KEY)
    print(f'{URL}flows/get_sharepoint_columns_v2?list={list_}&site={site}&API_KEY=ABCDEFG')
    print(rs.text)
import base64
def test_excel_functions():
    url = URL+f'excel_dagbok?application=Dagbok&filename="abcdefg.xlsx"'
    with open(os.path.join(os.path.dirname(__file__),'Dagbok (2).xlsx'), 'rb') as file:
        print(url)
        response = requests.post(url, data=file.read())
    with open(os.path.join(os.path.dirname(__file__),'result.xlsx'), 'wb') as file:
        file.write(response.content)
    fdict = eval(response.content)
    file = base64.b64decode(fdict["content"])
    with open(os.path.join(os.path.dirname(__file__),fdict["filename"]), 'wb') as f:
        f.write(file)
        
        
        
        
def test_excel_functions2():
    url = URL+f'excel_dagbok?application=Dagbok&filename="abcdefg.xlsx"'
    
    files = {'media': open(os.path.join(os.path.dirname(__file__),'Dagbok (2).xlsx'), 'rb')}
    im = requests.post(url, files=files)
    with open(os.path.join(os.path.dirname(__file__),'horse.xlsx'), 'wb') as f:
        f.write(im.content)
#test_excel_functions()

def test_Word_functions():
    url = URL + f'word_dokument_for_protokoll?API_KEY=ABCDEFG'
    print(url)
    with open('ActionOutputs (2).json', 'r', encoding="utf-8") as f:
        js = json.load(f)
    rs = requests.post(url, json.dumps(js), headers={"Content-Type": "application/json"})
    base64_file = json.loads(rs.content)["content"]
    file_content=base64.b64decode(base64_file)
    file_content=io.BytesIO(file_content)
    file_content.seek(0)
    with open('Test.docx', 'wb') as f:
        f.write(file_content.read())
test_Word_functions()