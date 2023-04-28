from flask import Blueprint, render_template, request, send_file,jsonify, abort
from functions.authentication import require_api_key
from functions.Excel.Get_Excel_data_to_json import  convert_file_to_workbook
import os,io,base64, mimetypes, openpyxl
excel_dagbok = Blueprint('excel_dagbok_trädexperterna', __name__)

@excel_dagbok.route("/api/excel_dagbok", methods=["POST"])
@require_api_key
  
  
 
def upload():
    # Get the file from the request
    print("Hello")
    file = request.files['file']
    wb = openpyxl.load_workbook(file)
    wb.save(os.path.join(os.path.dirname(__file__),'temp.xlsx'))
    # Do whatever you need to do with the file here
    # ...

    return jsonify({"content": "Hello"}) 
  
  
  
  
  
def get_excel_file(): #WORKING
    file_content = request.json.get('content')
    file_content=base64.b64decode(file_content)
    file_content=io.BytesIO(file_content)
    excel_file,filename=convert_file_to_workbook(file_content)
    file_content_base64 = base64.b64encode(excel_file.read()).decode('utf-8')
    return jsonify({"content":file_content_base64,"filename":filename})