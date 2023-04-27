
from flask import Blueprint, request, jsonify
from functions.Image_api import resize_and_autoorient
from functions.authentication import require_api_key
import base64, io

image_resize = Blueprint('image_resize', __name__)

@image_resize.route("/api/image_resizer", methods=["POST"])
@require_api_key
def post_file():
    file_content = request.json.get('content')
    width = int(request.json.get('width'))
    height = int(request.json.get('height'))
    file_content=base64.b64decode(file_content)
    file_content=io.BytesIO(file_content)
    img_file=resize_and_autoorient(file_content, height,width)
    file_content=base64.b64encode(img_file.read()).decode('utf-8')
    return jsonify({"content":file_content})
    
def post_file():
    """Upload a file and resize it to 300x300 px or specified input"""
    file_content_base64 = request.json['content']
    width = request.json.get('width')
    height = request.json.get('height')
    if not width:width=300
    if not height:height=300
    width,height = int(width),int(height)
    # decode the base64-encoded file content
    file_content = base64.b64decode(file_content_base64)
    img_file = resize_and_autoorient(file_content, width,height)
    file_content_base64 = base64.b64encode(img_file.read()).decode('utf-8')
    return jsonify({"content":file_content_base64})