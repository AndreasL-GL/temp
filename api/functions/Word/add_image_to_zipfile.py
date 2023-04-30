import io
import zipfile
import zlib
from PIL import Image
def add_icon_to_word_file(word_file, icon_file):
    with zipfile.ZipFile(io.BytesIO(word_file), mode='r') as zf:

        files = {x: zf.read(x) for x in zf.namelist()}
        files['word/media/image2.png'] = icon_file

    # Create a new in-memory ZipFile object
    new_word_file = io.BytesIO()
    with zipfile.ZipFile(new_word_file, mode='w', compression=zipfile.ZIP_DEFLATED) as zkf:
        for name, file in files.items():
            # Add each file to the new zipfile
            zkf.writestr(name, file)

    # Reset the output_zip's file pointer to the beginning
    new_word_file.seek(0)
    return new_word_file
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