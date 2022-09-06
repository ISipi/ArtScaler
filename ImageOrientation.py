from PIL import Image
from PIL.ExifTags import TAGS


class ImageOrientation:
    def __init__(self, file):
        self.input_img = file
        self.image = Image.open(self.input_img)

        # extract other basic metadata
        self.info_dict = {
            "Filename": self.image.filename,
            "Image Size": self.image.size,
            "Image Height": self.image.height,
            "Image Width": self.image.width,
            "Image Format": self.image.format,
            "Image Mode": self.image.mode,
            "Image is Animated": getattr(self.image, "is_animated", False),
            "Frames in Image": getattr(self.image, "n_frames", 1)
        }

        self.exifdata = self.image.getexif()

    def _print_info(self):
        for label, value in self.info_dict.items():
            print(f"{label:25}: {value}")
        self.re_orient(True)

    def re_orient(self, print_data=False):
        for tag_id in self.exifdata:
            # get the tag name, instead of human unreadable tag id
            tag = TAGS.get(tag_id, tag_id)
            data = self.exifdata.get(tag_id)
            # decode bytes
            if isinstance(data, bytes):
                try:
                    data = data.decode()
                except Exception:
                    continue
            if print_data:
                print(f"{tag:25}: {data}")
            if tag == 'Orientation':
                #print(tag, str(data), self.input_img)
                if str(data) in ['2', '4', '5', '7']:
                    self.image = self.image.transpose(method=Image.Transpose.FLIP_LEFT_RIGHT)
                if str(data) in ['5', '6']:
                    #print('image should be rotated -90')
                    self.image = self.image.rotate(-90, expand=True)
                elif str(data) in ['7', '8']:
                    self.image = self.image.rotate(90, expand=True)
                elif str(data) in ['3', '4']:
                    #print('image should be rotated 180')
                    self.image = self.image.rotate(180, expand=True)
        return self.image