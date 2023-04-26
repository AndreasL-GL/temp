
from flask import Blueprint, request, send_file
from functions.Image_api import resize_and_autoorient
from functions.authentication import require_api_key

image_resize = Blueprint('image_resize', __name__)

@image_resize.route("/api/image_resizer", methods=["POST"])
@require_api_key
def post_file():
    """Upload a file and resize it to 300x300 px or specified input"""
    filename = request.args.get('filename')
    file = request.files['media']
    img_file = resize_and_autoorient(file)

    return send_file(img_file,download_name=filename,as_attachment=True)
