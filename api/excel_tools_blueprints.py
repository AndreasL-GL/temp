from flask import Blueprint, render_template, request, send_file,jsonify, abort
from functions.authentication import require_api_key
from functions.Excel.Get_Excel_data_to_json import  convert_file_to_workbook
import os,io,base64, mimetypes
excel_dagbok = Blueprint('excel_dagbok_trädexperterna', __name__)

@excel_dagbok.route("/api/excel_dagbok", methods=["POST"])
@require_api_key
  
def get_excel_file(): #WORKING
    file_content = request.json.get('content')
    
    file_content=base64.b64decode(file_content)
    file_content=io.BytesIO(file_content)
    fixBadZipfile(file_content)
    excel_file,filename=convert_file_to_workbook(file_content)
    file_content_base64 = base64.b64encode(excel_file.read()).decode('utf-8')
    return jsonify({"content":file_content_base64,"filename":filename})

def fixBadZipfile(zipFile):  
 f = open(zipFile, 'r+b')  
 data = f.read()  
 pos = data.find('\x50\x4b\x05\x06') # End of central directory signature  
 if (pos > 0):  
     f.seek(pos + 22)   # size of 'ZIP end of central directory record' 
     f.truncate()  
     f.close()  
 else:  
     # raise error, file is truncated 
     abort('Hello')