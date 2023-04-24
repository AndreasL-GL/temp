import pandas as pd
from flask import jsonify, request
import json, os
import Levenshtein


class SharepointColumns():
    def get(body):
        print(body, type(body))
        with open('./logs/sharepointlog.json', 'w') as f:
            json.dump(body,f,indent=4)
        body = SharepointColumns.remove_columns(body)
        return jsonify(body)
    
    def remove_columns(js):
        body = js['value'][0]
        templist = []
        columns_to_remove = ['@odata','Modified', 'Created', 'Author', 'Author#Claims', 'Editor', 'Editor#Claims', \
            'Identifier', 'IsFolder', 'Thumbnail', 'Link', 'Name', 'FilenameWithExtension', 'Path', 'FullPath', \
                'HasAttachments', 'VersionNumber','ItemInternalId', 'ID', 'Title', 'Ort', \
                    'Omr_x00e5_de', 'Kontrollmoment#Id', 'Infotillpersonal', 'OData__x00c5_terkommandekontroll', 'Egenkontrollklar']
        kontrollmoment = [x['Value'] for x in body.pop('Kontrollmoment')]
        for key in body.keys():
            
            
            for column in columns_to_remove:
                if column.lower() in key.lower():
                    templist.append(key)
        for key in templist: 
            try: body.pop(key)
            except Exception as e:
                print(e)
        print(kontrollmoment)
        kontrollpunkter_from_link = [x.replace('_x00f6','').replace('_x002c','') if '_x00f6' or '_x002c' in x else x for x in [x.replace('_x00e4_','ä').replace('_x00e5_','å').replace('_x00f6_','ö')\
            .replace('_x002d_','-').replace('_x002f_','/') \
                .replace('_x002c_', ',').replace('_x00e9_',' é').replace('Vårstädningutf', 'Vårstädning')
                for x in list(body.keys())]]
        
        kontrollpunkter_from_link = [x.split(',')[0] if ',' in x else x for x in kontrollpunkter_from_link]
        resultdict = SharepointColumns.levenshtein_dictionary(kontrollmoment,kontrollpunkter_from_link)
        resultdict = {key: list(body.keys())[kontrollpunkter_from_link.index(value)] for key,value in resultdict.items()}
        try: resultdict["Vårsopning"] = "Vårsopning".replace('å','_x00e5_')
        except Exception as e: print(e)
        return resultdict
    
    def levenshtein_dictionary(list1,list2):
        result_dict = {}

        for item1 in list1:
            max_similarity = 0
            most_similar_item = ''
            for item2 in list2:
                similarity = 1 - (Levenshtein.distance(item1, item2) / max(len(item1), len(item2)))
                if similarity > max_similarity:
                    max_similarity = similarity
                    most_similar_item = item2

            result_dict[item1] = most_similar_item
        return result_dict




if __name__ == '__main__':
    with open(os.path.join(os.path.join(os.path.dirname(os.path.dirname(__file__)), "logs"),'sharepointlog.json'), encoding="utf-8") as f:
        js = json.load(f)
        jsons = SharepointColumns.remove_columns(js)
        # print(json.dumps(jsons, indent=3))