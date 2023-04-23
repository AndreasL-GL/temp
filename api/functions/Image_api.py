from PIL import Image
import io
from config import ImageConfig
class Resize_Image():
    def autoorient(image):
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

            # get orientation value from exif data
            orientation = exif.get(274)

            # if orientation value is not recognized, return image unmodified
            if orientation not in ORIENTATIONS:
                return image

            # apply transformations based on orientation value
            for transform in ORIENTATIONS[orientation]:
                image = image.transpose(transform)

            # remove orientation tag from exif data
            exif.pop(274, None)

            # set modified exif data on image and return it
            # image.info["exif"] = ExifTags.TAGS.items()
            return image

    def get(file):
        f = Image.open(file)
        print(f.getexif())
        f = Resize_Image.autoorient(f)
        print(f.getexif())
        f = f.resize(ImageConfig.image_size_after_resize)
        img_file = io.BytesIO()
        f.save(img_file, format='JPEG')
        img_file.seek(0)
        return img_file