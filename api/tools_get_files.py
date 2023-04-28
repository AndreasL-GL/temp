from flask import request, jsonify,base64
import io
def get_file_from_content():
    # Get the file from the request
    file_data = request.files
    file_data = request.get_data()
    file = io.BytesIO(file_data)
    
    filebytes = io.BytesIO()
    filebytes.write(file.getvalue())
    filebytes.seek(0)
    return filebytes
    
def return_file_as_base64(file):    
    file_content_base64 = base64.b64encode(file.read()).decode('utf-8')
    return jsonify(file_content_base64)