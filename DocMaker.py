"""
DocMaker:

Two classes:
    1) BuildCSV creates a CSV file that lists all jpg and png files in a given folder
    2) Scaler scales the images in that CSV file and saves them in a docx document.
"""

from docx import Document
from docx.shared import Cm
from numpy import nan
import pandas as pd
from os import listdir
from os.path import join


class BuildCSV:
    def __init__(self, image_files):
        self.image_files = image_files
        self.image_files_cleaned = self.clean()

    def clean(self) -> list:
        assert isinstance(self.image_files, list)
        new_img_lst = []
        for img in listdir(self.image_files[0]):
            if img[-4:].lower() in ['.jpg', '.png']:
                new_img_lst.append(join(self.image_files[0], img))
        return new_img_lst

    def build(self, filename:str, scale:float) -> str:

        dataframe = pd.DataFrame({'Image_file': self.image_files_cleaned})
        dataframe['Scale'] = float(scale)
        dataframe['Artwork_life_size'] = nan
        dataframe['Artwork_frame_life_size'] = nan
        dataframe.to_csv(f'{filename}.csv')

        return f'{filename}.csv'


class Scaler:
    def __init__(self, csv_file, save_location, docx_filename):

        print(csv_file, save_location, docx_filename)
        self.csv_file = csv_file  # location of the csv file
        self.save_location = save_location  # location to save the docx
        self.docx_filename = join(self.save_location, f'{docx_filename}.docx')  # final file name
        self.document = Document()

    def iterate_images(self):
        df = pd.read_csv(self.csv_file)

        for index, row in df.iterrows():
            try:
                real_height = row['Artwork_life_size']
                scale = row['Scale']
                target_img_size_in_cm = self.calculate_size_of_img_in_cm(real_height, scale)
                self.document.add_picture(row['Image_file'], height=Cm(target_img_size_in_cm))
                self.document.add_paragraph(str(row['Image_file']))
            except ValueError:
                continue

    def calculate_size_of_img_in_cm(self, object_height: float, scale: int) -> float:
        target_size_of_img_in_cm = object_height / scale
        return target_size_of_img_in_cm

    def run_scaler(self):
        self.iterate_images()
        self.document.save(self.docx_filename)
