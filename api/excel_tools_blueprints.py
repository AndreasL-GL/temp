from flask import Blueprint, render_template, request, send_file
from functions.authentication import require_api_key
from functions.Excel.Get_Excel_data_to_json import  convert_file_to_workbook
excel_dagbok = Blueprint('excel_dagbok_trädexperterna', __name__)

@excel_dagbok.route("/api/excel_dagbok", methods=["POST"])
@require_api_key
def post_excel_file():
    """Upload a file and resize it to 300x300 px or specified input"""
    filename = request.args.get('filename')
    application = request.args.get('application')
    excel_file = request.files['media']
    if application=='Dagbok': excel_file=convert_file_to_workbook(excel_file)


    return send_file(excel_file,download_name=filename,as_attachment=True)