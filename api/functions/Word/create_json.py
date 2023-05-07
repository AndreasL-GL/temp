import json
import io
import base64
import datetime

def decode_file(base64_file):
    file_content=base64.b64decode(base64_file)
    file_content=io.BytesIO(file_content)
    file_content.seek(0)
    return file_content
    
    

def create_json_for_word_functions(js):
    """Creates the json structure needed specifically for GL Groups Egenkontroller protokoll.

    Args:
        js (dict): Json dictionary from POST-request.

    Returns:
        dict: Restructured json-dictionary.
    """
    bad_keys = ["Bildkommentar","'Återkommande kontroll","Funktionskontroll klar","Test","Egenkontroll klar","Fastighetstyp","Objektnummer" "Egenkontroll klar", "Förvaltare","Ort","Extra tillsyn","Team","Turordning","Objektnummer","Fastighetstyp"]
    js['Kontrollmoment'] = [item for item in js["Kontrollmoment"] if item["Moment"] not in bad_keys]
    links = [i['link'] for i in js['Kontrollmoment']]
    kontroller = []
    for item in js['Items']['value']:
        kontrollmoment = [link for link in item.keys() if link in links and item[link] and link != "Test"]
        kontrollmomentklar = [item[moment] for moment in kontrollmoment]
        kontrollmoment = [js["Kontrollmoment"][key]["Moment"] for key in range(len(js["Kontrollmoment"])) if js["Kontrollmoment"][key]["link"] in kontrollmoment]
        
        kontroller.append({
            "Område": item['Omr_x00e5_de'] + ', ' + item['Ort'],
            "Kommentar":[item["Bildkommentar"] if "Bildkommentar" in item.keys() else ""][0],
            "Klar":["Ja" if item["Egenkontrollklar"] else "Nej"][0],
            "Moment": ', '.join(kontrollmoment)
        })
    fID = [{"Bild":decode_file(item['Item']), "ID":item["Item_ID"]} for item in js['Attachments']]
    items_list = []
    for item in js['Items']['value']:
        items_list.append([{
            "Bild": fIDItem["Bild"],
            "Område": item['Omr_x00e5_de'] + ', ' + item['Ort'],
            "Bildkommentar":[item["Bildkommentar"] if "Bildkommentar" in item.keys() else ""][0]
        } for fIDItem in fID if fIDItem['ID']==item['ID']])
    komplett_bildlista = []
    for i in items_list:
        for j in i:
            komplett_bildlista.append(j)
            
    return {"Vecka":str(datetime.datetime.now().isocalendar().week),"Utskriftsdatum":datetime.datetime.now().strftime("%Y-%m-%d"),"images": komplett_bildlista, "content":kontroller}




if __name__ == '__main__':
        
    with open('ActionOutputs (2).json', 'r', encoding="utf-8") as f:
        js = json.load(f)
    create_json_for_word_functions(js)
        
    