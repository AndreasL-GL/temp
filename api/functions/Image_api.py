from PIL import Image
import io, os
import configparser
from flask import request
config = configparser.ConfigParser()
config.read(os.path.join(os.path.join(os.path.dirname(os.path.dirname(__file__)),'config'),"config.ini"))



def autoorient(image):
    """Accepts a PIL image item as input, returns a PIL image item as output."""
    # get exif data from image
    try:
        exif = image._getexif()
    except AttributeError:
        exif = None

    # if image has no exif data, return it unmodified
    if exif is None:
        return image

    # define exif orientation values and corresponding transformations
    ORIENTATIONS = {
        3: (Image.ROTATE_180,),
        6: (Image.ROTATE_270,),
        8: (Image.ROTATE_90,),
        2: (Image.FLIP_LEFT_RIGHT,),
        4: (Image.FLIP_TOP_BOTTOM, Image.ROTATE_180),
        5: (Image.FLIP_LEFT_RIGHT, Image.ROTATE_270),
        7: (Image.FLIP_LEFT_RIGHT, Image.ROTATE_90),
    }
    
    orientation = exif.get(274)
    if orientation not in ORIENTATIONS:
        return image

    for transform in ORIENTATIONS[orientation]:
        image = image.transpose(transform)


    return image

def resize_and_autoorient(file):
    """Accepts a file bytes object and returns a file bytes object
    Resizes an image based on specifications in the config."""
    f = Image.open(file)
    f = autoorient(f)
    height, width = request.args.get("height"),request.args.get("width")
    if not height or not width:
        f = f.resize(eval(config["IMAGE_API"]["IMAGE_SIZE_AFTER_RESIZE"]))
    else:
        f=f.resize((int(height),int(width)))
    # Create a bytes object to send in response
    img_file = io.BytesIO()
    f.save(img_file, format='JPEG')
    img_file.seek(0)
    return img_file