import json
import os
import requests
from functions.downloader.sharepoint import get_sharepoint_access_headers_through_client_id,get_item_based_on_id, get_all_items, get_by_url, Download_icon
import datetime
import io
import mailmerge
from PIL import Image
import docx
from docx.shared import Pt, Inches, RGBColor
from Office_helper_functions.Word.form_field import set_checkbox_value
from Office_helper_functions.Image.Image_operations import resize_and_autoorient
from functions.Word.add_image_to_zipfile import add_icon_to_word_file
import base64
from docx.enum.text import WD_BREAK


icons = {
         "Green Landscaping AB":"https://greenlandscapingmalmo.sharepoint.com/sites/Funktionskontrolllekplatsdemo/_api/web/GetFileByServerRelativeUrl('/sites/Funktionskontrolllekplatsdemo/Delade Dokument/Protokoll lekplats besiktning demo/Template/Loggor/GLAB_png.png')/$value",
         "Green Landscaping Skåne AB":"https://greenlandscapingmalmo.sharepoint.com/sites/Funktionskontrolllekplatsdemo/_api/web/GetFileByServerRelativeUrl('/sites/Funktionskontrolllekplatsdemo/Delade Dokument/Protokoll lekplats besiktning demo/Template/Loggor/GLAB_png.png')/$value",
         "Svensk Markservice":"https://greenlandscapingmalmo.sharepoint.com/sites/Funktionskontrolllekplatsdemo/_api/web/GetFileByServerRelativeUrl('/sites/Funktionskontrolllekplatsdemo/Delade Dokument/Protokoll lekplats besiktning demo/Template/Loggor/SvMS_png.png')/$value",
         "Green Landscaping Malmö AB":"https://greenlandscapingmalmo.sharepoint.com/sites/Funktionskontrolllekplatsdemo/_api/web/GetFileByServerRelativeUrl('/sites/Funktionskontrolllekplatsdemo/Delade Dokument/Protokoll lekplats besiktning demo/Template/Loggor/GL_Malmo_png.png')/$value",
         "Markservice Stockholm AB":"https://greenlandscapingmalmo.sharepoint.com/sites/Funktionskontrolllekplatsdemo/_api/web/GetFileByServerRelativeUrl('/sites/Funktionskontrolllekplatsdemo/Delade Dokument/Protokoll lekplats besiktning demo/Template/Loggor/markservice_sthlm_png.png')/$value",
         "PARK i Syd AB":"https://greenlandscapingmalmo.sharepoint.com/sites/Funktionskontrolllekplatsdemo/_api/web/GetFileByServerRelativeUrl('/sites/Funktionskontrolllekplatsdemo/Delade Dokument/Protokoll lekplats besiktning demo/Template/Loggor/PARK green 4F.png')/$value",
         "Tranemo Trädgårdstjänst AB":"https://greenlandscapingmalmo.sharepoint.com/sites/Funktionskontrolllekplatsdemo/_api/web/GetFileByServerRelativeUrl('/sites/Funktionskontrolllekplatsdemo/Delade Dokument/Protokoll lekplats besiktning demo/Template/Loggor/TTJ_logo_ute_webb.png')/$value",
         "Grön Stad":"https://greenlandscapingmalmo.sharepoint.com/sites/Funktionskontrolllekplatsdemo/_api/web/GetFileByServerRelativeUrl('/sites/Funktionskontrolllekplatsdemo/Delade Dokument/Protokoll lekplats besiktning demo/Template/Loggor/Gron_Stad_logo_tag_RGB_270px.png')/$value"
         }




def get_cert_no(site,certlist, fitness):
    items = get_all_items(site=site, list_=certlist)
    itemsurl = items["d"]["Items"]["__deferred"]["uri"]
    items = requests.get(itemsurl)


def create_protocol(site, lista, js):
    certifikatjs = js['Certifikatinfo']['value'][0]
    trigger = js['Trigger']
    
    js1 = js["Items"]['value'][0]

    
    if js1["Telefonnummer"] == 0 or js1["Telefonnummer"] == "0":
        return {"status":"failed"}
    js1['Hemsida'] = certifikatjs['Hemsida']
    js1['Email'] = js1['Author']['Email']
    js1['Telefonnummer'] = certifikatjs['Telefonnummer']
    js1['Dagensdatum'] = datetime.datetime.strftime(datetime.datetime.now(), "%Y-%m-%d")
    js1['Bolag'] = certifikatjs['Bolag']['Value']
    js1["Certifieringstext2"] = """Klagomål

I det fall Ni har synpunkter på utförd säkerhetsbesiktning kan Ni Kontakta: 

SERENO Certifiering AB
Box 5604

114 86  STOCKHOLM 
Telefon: 08-556 953 30 
"""
    js1["Certifieringstext"] = """Innehar Certifikat nr Click or tap here to enter text. utfärdat av
SERENO Certifiering AB
Platser för motion eller annan utevistelse
Samt av fitnessutrustning
"""
    
    doc = populate_template(js1,certifikatjs,js,trigger)
    return doc
    
    
def populate_template(js1, certifikatjs, js, trigger):
    trigger['DigitalsignaturLekplats'] = False
    
    
    if trigger['DigitalsignaturLekplats'] or trigger['DigitalsignaturUtegym']:
        js1['Digital signatur'] = "Härmed intygas att besiktningen utförts enligt gällande regler."
        js1['Digital signatur 2'] = "Digitalt signerad av",js1['Author']['DisplayName']+',', js1['Created'].split('T')[0]

    js1["Besmantelefonnummer"] = certifikatjs["Telefonnummer"]
    js1["Adresstillprotokoll"] = certifikatjs['Adresstillprotokoll']
    js1['Created'] = js1['Created'].split('T')[0]
    if not trigger['DigitalsignaturUtegym'] and js1['Fitnessbesiktning']:
        doc = mailmerge.MailMerge(os.path.join(os.path.dirname(__file__), 'Fitness mall ej cert.docx'))
    elif trigger['DigitalsignaturUtegym'] and js1['Fitnessbesiktning']:
        doc = mailmerge.MailMerge(os.path.join(os.path.dirname(__file__), 'Fitness mall cert.docx'))
    elif not trigger['DigitalsignaturLekplats'] and not js1['Fitnessbesiktning']:
        doc = mailmerge.MailMerge(os.path.join(os.path.dirname(__file__), 'Lekplatsbesiktning mall cert.docx'))
    elif trigger['DigitalsignaturUtegym'] and not js1['Fitnessbesiktning']:
        doc = mailmerge.MailMerge(os.path.join(os.path.dirname(__file__), 'Lekplatsbesiktning mall ej cert.docx'))
    
        
 

    js1["H_x00e4_nvisningsskylt"] = "Ja" if js1["H_x00e4_nvisningsskylt"]==True else "Nej"
    js1["Rutinerolyckor_x002f_incidenter"] = "Ja" if js1["Rutinerolyckor_x002f_incidenter"]==True else "Nej"
    js1["Rutinerf_x00f6_runderh_x00e5_ll"] = "Ja" if js1["Rutinerf_x00f6_runderh_x00e5_ll"]==True else "Nej"
    js1["Leverant_x00f6_rensanvisningar"] = "Ja" if js1["Leverant_x00f6_rensanvisningar"]==True else "Nej"
    js1["Sparandeavdokument"] = "Ja" if js1["Sparandeavdokument"]==True else "Nej"

    mergefields = {key:value for key,value in js1.items() if key in [x for x in doc.get_merge_fields()]}
    icon_file = requests.get(icons[js1['Bolag']], headers=get_sharepoint_access_headers_through_client_id()).content

    doc.merge(**mergefields)
    
    
    ### TYPE CHANGE: type(doc) = docx.Document object from here.
    doc = change_icon_in_header(doc, icon_file, "word/media/image1.png")
    if js1['Typavbesiktning']['Value'] == "Installationsbesiktning":
        file = io.BytesIO()
        doc.save(file)
        file.seek(0)
        doc = docx.Document(set_checkbox_value(file.read(), 1,1))
    else: 
        file = io.BytesIO()
        doc.save(file)
        file.seek(0)
        doc = docx.Document(set_checkbox_value(file.read(), 1,0)) #Årlig besiktning
    
    
    if not js1["Fitnessbesiktning"]:
        print(js1['Fysiskomfattning']['Value'])
    
        match js1["Fysiskomfattning"]['Value']:
            case "Endast lekredskap":
                file = io.BytesIO()
                doc.save(file)
                file.seek(0)
                doc = docx.Document(set_checkbox_value(file.read(), 1,3)) # Årlig besiktning
            case "Lekplats inklusive":
                file = io.BytesIO()
                doc.save(file)
                file.seek(0)
                doc = docx.Document(set_checkbox_value(file.read(), 1,3)) # Lekplats inklusive
            case "Lekplats inkl staket mm":
                file = io.BytesIO()
                doc.save(file)
                file.seek(0)
                doc = docx.Document(set_checkbox_value(file.read(), 1,3)) # Lekplats inklusive
            case "Område på karta":
                file = io.BytesIO()
                doc.save(file)
                file.seek(0)
                doc = docx.Document(set_checkbox_value(file.read(), 1,5)) # Område på karta
        match js1["Drift_x002d_ochunderh_x00e5_llsp"]:
            case True:
                file = io.BytesIO()
                doc.save(file)
                file.seek(0)
                doc = docx.Document(set_checkbox_value(file.read(), 1,6)) # Underhållsplan True
            case False:
                file = io.BytesIO()
                doc.save(file)
                file.seek(0)
                doc = docx.Document(set_checkbox_value(file.read(), 1,7)) # Underhållsplan False
    st = [style.type for style in doc.styles if style.name == 'Heading 1'][0]
    bighead = doc.styles.add_style('Big heading', st)
    bighead.font.size = Pt(22)
    bighead.font.color.rgb = RGBColor(100,200,100)

    doc.add_page_break()
    add_översiktsbild(doc,js)
    add_utrustning(doc,js)
    add_page_break(doc)
    add_anmärkningar(doc,js)
    add_page_break(doc)
    add_underlag(doc,js)
    if any(js['Staket']):
        add_grindar(doc,js)
    if any(js['Brunnar']):
        add_brunnar(doc,js)
    doc.save(os.path.join(os.path.dirname(__file__), 'steg_1.docx'), )
    return doc



def add_översiktsbild(doc,js):
    if js['Trigger']['Sida2']:
        image = io.BytesIO(base64.b64decode(js['Översiktsbild']))
        image.seek(0)
        h=doc.add_heading('Översiktsbild av lekplatsen', 0)
        h.style = 'Big heading'
        img = resize_and_autoorient(image,300,350)
        table = doc.add_table(rows=1,cols=1)
        row = table.add_row().cells
        row[0].add_paragraph().add_run().add_picture(img)
    return None
def add_utrustning(doc,js):
    print(js["Utrustning"][0]['Items'].keys())
    # Add the table for the Utrustnings items:
    h = doc.add_heading('Produktbeskrivning')
    h.style = 'Big heading'
    subheading = doc.styles.add_style('subheading', h.style.type)
    subheading.font.size = Pt(16)
    subheading.font.color.rgb = RGBColor(100,200,100)
    small = doc.styles.add_style('small', h.style.type)
    small.font.size = Pt(10)
    small.font.italic = True
    vsmall= doc.styles.add_style('vsmall', h.style.type)
    vsmall.font.size = Pt(8)
    # h.style.font.color.rgb = RGBColor(100,200,100)
    table = doc.add_table(rows=1, cols=5)
    table.style = 'Grid Table 1 Light'
    row = table.rows[0].cells
    row[0].text = 'Nr'
    row[0].width = Inches(0.2)
    row[1].text = "Produkt"
    row[2].text = "Tillverkare/artnr"
    row[3].text = "Årtal"
    row[4].text = "Bild nr"
    
    for i, item in enumerate(js["Utrustning"]):
        row = table.add_row().cells
        row[0].text = str(i+1)
        row[0].width = Inches(0.2)
        row[1].text = item['Items']['Utrustning']['Value']
        row[2].text = item['Items']['Tillverkare_x002f_artnr']
        row[3].text = item['Items']['OData__x00c5_rtal']
        row[4].text = "Bild: "+str(i+1)
    
    for cell in table.columns[0].cells:
        cell.width = Inches(0.5)
    for cell in table.columns[1].cells:
        cell.width = Inches(2.5)
    for cell in table.columns[2].cells:
        cell.width = Inches(2.5)
    for cell in table.columns[3].cells:
        cell.width = Inches(0.7)
    for cell in table.columns[4].cells:
        cell.width = Inches(0.8)
    doc.add_page_break()
    doc.add_heading('Besiktningsresultat', 0).style = 'Big heading'
    table = doc.add_table(rows=1, cols=3)
    index = 0
    images = [x['Image'][0]['content'] for x in js['Utrustning']]
    while index < len(images):
        row = table.add_row().cells
        for i in range(0,3):
            
            file = io.BytesIO(base64.b64decode(js['Utrustning'][index]['Image'][0]['content']))
            file.seek(0)
            file = resize_and_autoorient(file,120,120)
            row[i].add_paragraph().text = "Bild: "+str(index+1)
            row[i].add_paragraph().add_run().add_picture(file)
            index +=1
            if index == len(images):break
    return None

def add_underlag(doc,js):
    doc.add_paragraph()
    hh = doc.add_heading('Stötdämpande underlag:', 0)
    hh.style = 'Big heading'
    #images = [img['Image'][0] for img in js['Anmärkningar']]
    
    for i, item in enumerate(js['Underlag']):
        table = doc.add_table(rows=1, cols=2)
        table.style = 'Grid Table Light'
        row = table.rows[0].cells
        row[0].text = item['Kommentar']
        row[0].style = 'vsmall'
        row[1].text = item['Bed_x00f6_mning']['Value']
        p = doc.add_paragraph()
        p.text = "Enligt SS-EN 1176-1:4.2.8.5"
        p.style = 'small'
        for cell in table.columns[0].cells:
            cell.width = Inches(6)
        for cell in table.columns[1].cells:
            cell.width = Inches(0.4)
    return None
    
    
def add_anmärkningar(doc, js):
    doc.add_paragraph()
    hh = doc.add_heading('Anmärkningar:', 0)
    hh.style = 'Big heading'
    #images = [img['Image'][0] for img in js['Anmärkningar']]
    for i, item in enumerate(js['Anmärkningar']):
        h = doc.add_heading('Produkt '+str(i+1)+', '+ item['Items']['Utrustningstyp0'], 0)
        h.style= 'subheading'
        if item['Items']['{HasAttachments}']:
            if item['Image'][0]['content']:
                img = item['Image'][0]['content']
                file = io.BytesIO(base64.b64decode(img))
                file.seek(0)
                img = resize_and_autoorient(file, 80,80)
                doc.add_paragraph().add_run().add_picture(img)
        table = doc.add_table(rows=1, cols=2)
        table.style = 'Grid Table Light'
        row = table.rows[0].cells
        row[0].text = item['Items']['Kommentar']
        row[1].text = item['Items']['Bed_x00f6_mning']['Value']
        p = doc.add_paragraph()
        p.text = item['Items']['Utrustningstyp']['Value']
        p.style = 'small'
        for cell in table.columns[0].cells:
            cell.width = Inches(6)
        for cell in table.columns[1].cells:
            cell.width = Inches(0.4)
    return None

def add_grindar(doc, js):
    doc.add_paragraph()
    hh = doc.add_heading('Grindar och staket', 0)
    hh.style = 'Big heading'
    for i, item in enumerate(js['Staket']):
        
        if item['Items']['{HasAttachments}']:
            if item['Images'][0]['content']:
                img = item['Images'][0]['content']
                file = io.BytesIO(base64.b64decode(img))
                file.seek(0)
                img = resize_and_autoorient(file, 80,80)
                doc.add_paragraph().add_run().add_picture(img)
        table = doc.add_table(rows=1, cols=2)
        table.style = 'Grid Table Light'
        row = table.rows[0].cells
        row[0].text = item['Items']['Kommentar']
        row[1].text = item['Items']['Bed_x00f6_mning']['Value']
        p = doc.add_paragraph()
        p.text = item['Items']['Utrustningstyp']['Value']
        p.style = 'small'
        for cell in table.columns[0].cells:
            cell.width = Inches(6)
        for cell in table.columns[1].cells:
            cell.width = Inches(0.4)
    return None

def add_brunnar(doc,js):
    doc.add_paragraph()
    hh = doc.add_heading('Brunnar:', 0)
    hh.style = 'Big heading'
    #images = [img['Image'][0] for img in js['Anmärkningar']]
    for i, item in enumerate(js['Brunnar']):
        if item['Items']['{HasAttachments}']:
            if item['Image'][0]['content']:
                img = item['Image'][0]['content']
                file = io.BytesIO(base64.b64decode())
                file.seek(0)
                img = resize_and_autoorient(file, 80,80)
                doc.add_paragraph().add_run().add_picture(img)
        table = doc.add_table(rows=1, cols=2)
        table.style = 'Grid Table Light'
        row = table.rows[0].cells
        row[1].text = item['Items']['Anm_x00e4_rkning']
        row[0].text = item['Items']['Status']['Value']
        p = doc.add_paragraph()
        p.text = "Enligt ordningslagen 3 kap 5§"
        p.style = 'small'
        for cell in table.columns[0].cells:
            cell.width = Inches(6)
        for cell in table.columns[1].cells:
            cell.width = Inches(0.4)
    return None



def change_icon_in_header(doc, icon_file, imagepath):
    """Changes the icon in the header of a word file by replacing a pre-formatted placeholder image."""
    file_content = io.BytesIO(icon_file)
    file_content.seek(0)
    img = Image.open(file_content)
    fb = io.BytesIO()
    img.save(fb, 'PNG')
    fb.seek(0)
    docbyte = io.BytesIO()
    if "mailmerge" in str(type(doc)):
        doc.write(docbyte)
    else: doc.save(docbyte)
    docbyte.seek(0)
    doc = add_icon_to_word_file(docbyte.read(), icon_file=fb.read(), imagepath=imagepath)
    doc = docx.Document(doc)
    return doc


def add_page_break(document):
    body_height = document.element.body.get("height", 0)
    
    page_height = document.sections[-1].page_height.cm
    
    ratio = body_height / page_height
    
    if ratio >= 0.5:
        document.add_page_break()
    else:
        p = document.add_paragraph()
        p.add_run().add_break(WD_BREAK.PAGE)

def run_functions(js):
    print(js.keys())
    doc = create_protocol('Funktionskontrolllekplatsdemo',"Lista_lekplats_besiktningsprotokoll",js)
    file = io.BytesIO()
    doc.save(file)
    file.seek(0)
    return {"content": base64.b64encode(file.getvalue()).decode('utf-8'), "filename": "filename.docx"}


if __name__ == '__main__':
    with open(os.path.join(os.path.dirname(__file__),'sample.json'),'r', encoding="utf-8") as f:
        js = json.load(f)
    doc = create_protocol('Funktionskontrolllekplatsdemo',"Lista_lekplats_besiktningsprotokoll",js)
    
    
    'Fitness mall ej cert.docx'
    #get_cert_no("Funktionskontrolllekplatsdemo","Certifikatinformation","True")

    
    """"""
    # url = "https://greenlandscapingmalmo.sharepoint.com/sites/Funktionskontrolllekplatsdemo/_api/web/GetFileByServerRelativeUrl('/sites/Funktionskontrolllekplatsdemo/Lists/Lista_lekplats_besiktningsprotokoll/Attachments/2171/4929E5D2-6190-495A-AF7F-2523EC1D5AA9.jpg')/$value"
    # img = requests.get(url, get_sharepoint_access_headers_through_client_id())
    

    
    
    ### TODO:
    """
    Kontrollera att Besiktningsmannens telefonnummer verkligen är "Telefonnummer" i js1. Annars ändra i alla templates.
    Fortsätt med att skriva in värden för brunnar anmärkningar osv osv. enligt flödet.
    
    
    """