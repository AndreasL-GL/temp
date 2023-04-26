
from flask import Blueprint, request, jsonify
from functions.Image_api import resize_and_autoorient
from functions.authentication import require_api_key
import base64

image_resize = Blueprint('image_resize', __name__)

@image_resize.route("/api/image_resizer", methods=["POST"])
@require_api_key
def post_file():
    """Upload a file and resize it to 300x300 px or specified input"""
    file = request.files['document']
    img_file = resize_and_autoorient(file)

    file_content_base64 = base64.b64encode(img_file.read()).decode('utf-8')
    return jsonify(file_content_base64)