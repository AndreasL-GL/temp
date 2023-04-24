
from flask import Blueprint, request, send_file, jsonify
from functions.Image_api import Resize_Image
from functions.authentication import require_api_key

image_resize = Blueprint('image_resize', __name__)

@image_resize.route("/api/image_resizer", methods=["POST"])
@require_api_key
def post_file():
    """Upload a file and resize it to 300x300 px."""
    filename = request.args.get('filename')
    file = request.files['media']
    img_file = Resize_Image.get(file)

    return send_file(img_file,download_name=filename,mimetype=file.content_type,as_attachment=True)
