import requests
from PIL import Image,ExifTags
import io, json
import os
URL = "http://localhost/api/"
API_KEY = "ABCDEFG"
def test_image_resize():
    url = URL+f'image_resizer?filename=filnamn.jpg&API_KEY={API_KEY}'
    files = {'media': open('test.jpg', 'rb')}
    imm = Image.open("test.jpg")
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
    
test_sharepoint()