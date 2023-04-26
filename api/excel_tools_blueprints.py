from flask import Blueprint, render_template, request, send_file,jsonify
from functions.authentication import require_api_key
from functions.Excel.Get_Excel_data_to_json import  convert_file_to_workbook
import os,io,base64
excel_dagbok = Blueprint('excel_dagbok_trädexperterna', __name__)

@excel_dagbok.route("/api/excel_dagbok", methods=["POST"])
@require_api_key
def get_excel_file():
    file = request.files['document']
    excel_file=convert_file_to_workbook(file)
    file_content_base64 = base64.b64encode(file.read()).decode('utf-8')
    return jsonify(file_content_base64)

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