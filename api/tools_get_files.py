from flask import request, jsonify,base64
def get_file():
    file = request.files['document']
    return file
    
    file_content_base64 = base64.b64encode(file.read()).decode('utf-8')
    return jsonify(file_content_base64)