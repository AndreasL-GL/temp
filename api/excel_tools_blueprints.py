from flask import Blueprint, render_template, request, send_file,jsonify, abort
from functions.authentication import require_api_key
from functions.Excel.Get_Excel_data_to_json import  convert_file_to_workbook
import os,io,base64, zipfile
excel_dagbok = Blueprint('excel_dagbok_trädexperterna', __name__)

@excel_dagbok.route("/api/excel_dagbok", methods=["POST"])
@require_api_key

def get_excel_file():
    file_content = request.json.get('content')
    file_content=base64.b64decode(file_content)
    #file_content=io.BytesIO(file_content)
    file_content=convert_xmlpk_to_xlsx(file_content)
    excel_file,filename=convert_file_to_workbook(file_content)
    file_content_base64 = base64.b64encode(excel_file.read()).decode('utf-8')
    return jsonify({"content":file_content_base64,"filename":filename})
def get_excel_file(): #WORKING
    file_content = request.json.get('content')
    file_content=base64.b64decode(file_content)
    file_content=io.BytesIO(file_content)
    excel_file,filename=convert_file_to_workbook(file_content)
    file_content_base64 = base64.b64encode(excel_file.read()).decode('utf-8')
    return jsonify({"content":file_content_base64,"filename":filename})


def get_excel_file():
    file = request.files['document']
    excel_file,filename=convert_file_to_workbook(file)
    file_content_base64 = base64.b64encode(excel_file.read()).decode('utf-8')
    return jsonify({"content":file_content_base64,"filename":filename})

def post_excel_file():
    """Upload a file and resize it to 300x300 px or specified input"""
    filename = request.args.get('filename')
    application = request.args.get('application')
    excel_file = request.data
    if not excel_file: excel_file=request.files['content']
    with open(os.path.join(os.path.dirname(__file__),'hs'), 'wb') as f:
        f.write(request.files)
    if application=='Dagbok': excel_file=convert_file_to_workbook(excel_file)


    return send_file(excel_file,download_name=filename,as_attachment=True)

def convert_xmlpk_to_xlsx(xmlpk_file):

    with zipfile.ZipFile(io.BytesIO(xmlpk_file), 'r') as zip_ref:
        for file in zip_ref.namelist():
            if file.endswith('.xml'):
                workbook_xml = zip_ref.read(file)
                workbook_name = file[:-4] + '.xlsx'
                with zipfile.ZipFile(io.BytesIO(), mode='w') as xlsx_file:
                    xlsx_file.writestr('xl/workbook.xml', workbook_xml)
                    xlsx_file.writestr('xl/_rels/workbook.xml.rels', '')
                    xlsx_file.writestr('xl/theme/theme1.xml', '')
                    xlsx_file.writestr('xl/styles.xml', '')
                    xlsx_file.writestr('[Content_Types].xml', '')
                    xlsx_data = xlsx_file.read()

    return io.BytesIO(xlsx_data)
   