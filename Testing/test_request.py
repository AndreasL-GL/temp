import requests
from PIL import Image,ExifTags
import io, json
import os
URL = "http://localhost:3000/api/flows/get_sharepoint_columns_v2?listt=MKB Egenkontroll Periodiska 2023&site=GLMalmAB-EgenkontrollerVellingebostder&API_KEY="
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
    rs = requests.get("http://51.124.105.224/api/flows/get_sharepoint_columns_v2?listt=MKB%20Egenkontroll%20Periodiska%202023&site=GLMalmAB-EgenkontrollerVellingebostder&API_KEY=ABCDEFG")
#     rs = requests.get(f'{URL}flows/get_sharepoint_columns_v2?listt=MKB Egenkontroll Periodiska 2023&site=GLMalmAB-EgenkontrollerVellingebostder&API_KEY='+API_KEY)
    print(f'{URL}flows/get_sharepoint_columns_v2?list={list_}&site={site}&API_KEY='+API_KEY)
    print(rs.text)
    
test_sharepointv2()