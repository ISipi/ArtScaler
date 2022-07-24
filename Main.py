"""
Main.py:

Two classes:
    1) BrowseToItem inherits from FileBrowser and creates a modular popup window to choose files and folders
    2) MainView is the main app itself and provides an 8 step system to interactively create a Word document of scaled
        images from a given folder.

The steps for the user:
    1) They create a folder with images that they wish to scale
    2) They select the folder where they want to save the intermediate CSV file and the final .docx
    3) They give a filename for the intermediate CSV file
    4) They give a filename for the final .docx
    5) They assign a scale (1 returning just the image).
    6) They can now press 'Prepare CSV', which populates the CSV file with the filenames from the folder provided in
        step 1), as well as the scale.
    7) The user can now open the intermediate CSV file and input the life size of the images and modify the scale if
        they want. They can also assign the size of the frame for the artwork, but this is not implemented yet.
    8) Once all the rows of interest are populated, the user can now produce the final scaled images.

Alternatively, if the user has a CSV file with all the necessary information:
    1) They need to make sure the following columns are present (note capital letters):
            Image_file, Scale, Artwork_life_size, Artwork_frame_life_size
    2) They can now skip steps 1), 3), 5), 6), and 7)
    3) Instead of going to step 8, they must press Choose CSV and find the pre-populated CSV.
    4) Now, run the program.

"""

import kivy
kivy.require('2.1.0')

from kivy.app import App
from kivy.uix.gridlayout import GridLayout
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.label import Label
from kivy.uix.scrollview import ScrollView
from kivy.uix.textinput import TextInput
from kivy.uix.popup import Popup
from FileBrowser import FileBrowser
from os.path import sep, dirname, expanduser, join
import sys
from DocMaker import BuildCSV, Scaler


class ExitApp(Popup):
    def build(self):
        # create a layout widget within which we place two buttons:
        box = BoxLayout()
        cancel_action = Button(text='Stay in app')
        exit_app = Button(text='Exit app')
        box.add_widget(cancel_action)
        box.add_widget(exit_app)

        # use the above widget as content:
        popup = Popup(content=box, title='Are you sure you want to close this app?', auto_dismiss=False,
                      size_hint=(None, None), size=(200, 120))

        # create button bindings:
        exit_app.bind(on_press=self.exit)
        cancel_action.bind(on_press=popup.dismiss)

        # open the popup:
        popup.open()

    def exit(self, instance):
        App.get_running_app().stop()


class BrowseToItem(FileBrowser):

    def __init__(self, msg_to_log, select_folder=False, **kwargs):
        super().__init__(**kwargs)
        self.msg_to_log = msg_to_log
        if sys.platform == 'win':
            user_path = dirname(expanduser('~')) + sep + 'Documents'
        else:
            user_path = expanduser('~') + sep + 'Documents'
        self.browser = FileBrowser(select_string='Select',
                              favorites=[(user_path, 'Documents')],
                              dirselect=select_folder,
                              show_hidden=True)
        self.popup = Popup(content=self.browser, title='Select a file', auto_dismiss=False)
        self.browser.bind(on_success=self._fbrowser_success, on_canceled=self._fbrowser_canceled)

    def build(self):
        self.popup.open()
        return self.browser

    def _fbrowser_canceled(self, instance):
        self.popup.dismiss()

    def _fbrowser_success(self, instance):
        self.selection = instance.selection
        self.popup.dismiss()
        app = App.get_running_app()
        current_text = app.window.ids['error_text'].text
        new_text = f'{current_text}\n{self.msg_to_log}: {instance.selection[0]}.'
        app.window.ids['error_text'].text = new_text

    def return_selection(self):
        return self.selection


class MainView(App):

    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.window = None
        self.img_folder_popup = None
        self.get_csv_popup = None
        self.csv_file = None
        self.save_folder_popup = None

    def build(self):

        # main window:
        self.window = GridLayout(size_hint=(1, 1), pos_hint={'center_x':.5})
        self.window.cols = 1

        ### define widgets: ###
        image_folder_loc = Button(text='Step 1. Choose image folder', size_hint=(0.5, 0.5))
        save_location = Button(text='Step 2. Save location (for CSV and final document)', size_hint=(0.5, 0.5))
        input_grid = GridLayout(size_hint=(1,2))
        input_grid.cols = 2
        csv_filename_label = Label(text='Step 3. Give a filename for CSV file:')
        self.csv_filename = TextInput(text='Give a filename for CSV file', multiline=False)
        doc_filename_label = Label(text='Step 4. Give a filename for DOC file:')
        self.doc_filename = TextInput(text='Give a filename for DOC file', multiline=False)
        scale_label = Label(text='Step 5. Assign a scale:')
        self.scale = TextInput(text='Give a numeric value', multiline=False)
        prepare_csv = Button(text='Step 6a. Prepare CSV', size_hint=(0.5, 0.5))
        choose_csv = Button(text='Step 6b. Choose CSV', size_hint=(0.5, 0.5))
        run_job = Button(text='Step 7. Run the program', size_hint=(0.5, 0.5))

        # add scroll view within which we add a grid layout,
        # which then consists of one label that acts as our logging output
        scroll_view = ScrollView(do_scroll_y=True, do_scroll_x=False, size_hint=(1, 2))
        inner_grid = GridLayout(size=(scroll_view.width, scroll_view.height), size_hint=(1, None))
        inner_grid.cols = 1
        output_errors_and_success = Label(text='### LOGGING OUTPUT ###\n')

        # assign id to logger:
        self.window.ids['error_text'] = output_errors_and_success

        # add widgets to main window:
        self.window.add_widget(image_folder_loc)
        self.window.add_widget(save_location)
        input_grid.add_widget(csv_filename_label)  # sub_widget
        input_grid.add_widget(self.csv_filename)  # sub_widget
        input_grid.add_widget(doc_filename_label)  # sub_widget
        input_grid.add_widget(self.doc_filename)  # sub_widget
        input_grid.add_widget(scale_label)  # sub_widget
        input_grid.add_widget(self.scale)  # sub_widget
        self.window.add_widget(input_grid)
        self.window.add_widget(prepare_csv)
        self.window.add_widget(choose_csv)
        self.window.add_widget(run_job)

        inner_grid.add_widget(output_errors_and_success)  # sub_widget
        scroll_view.add_widget(inner_grid)  # sub_widget
        self.window.add_widget(scroll_view)

        # bind actions to widgets:
        image_folder_loc.bind(on_press=self.select_image_folder)  # button widget
        save_location.bind(on_press=self.select_save_folder)  # button widget
        self.csv_filename.bind(on_text_validate=self.store_csv_name)  # textinput widget
        self.doc_filename.bind(on_text_validate=self.store_doc_name)  # textinput widget
        self.scale.bind(on_text_validate=self.check_numeric)  # textinput widget
        prepare_csv.bind(on_press=self.build_csv)  # button widget
        choose_csv.bind(on_press=self.select_csv_file)  # button widget
        run_job.bind(on_press=self.create_final_document)

        return self.window

    def check_numeric(self, instance):
        try:
            assert isinstance(float(instance.text), float)
            msg_to_log = f'Set scale to'
            current_text = self.window.ids['error_text'].text
            new_text = f'{current_text}.\n{msg_to_log}: 1:{str(instance.text)} or {1 / float(instance.text) * 100:.2f}% life size'
            self.window.ids['error_text'].text = new_text
        except ValueError as e:
            action_to_take = 'Please give a number'
            self.error_msg(e, action_to_take)

    def store_csv_name(self, instance):
        msg_to_log = f'Currently saving CSV to'
        current_text = self.window.ids['error_text'].text
        new_text = f'{current_text}.\n{msg_to_log}: {instance.text}'
        self.window.ids['error_text'].text = new_text

    def store_doc_name(self, instance):
        msg_to_log = f'Currently saving DOC to'
        current_text = self.window.ids['error_text'].text
        new_text = f'{current_text}.\n{msg_to_log}: {instance.text}'
        self.window.ids['error_text'].text = new_text

    def select_image_folder(self, instance, select_folder=True):
        msg_to_log = 'Image folder selected'
        self.img_folder_popup = BrowseToItem(msg_to_log, select_folder=select_folder)
        self.img_folder_popup.build()

    def select_csv_file(self, instance, select_folder=False):
        msg_to_log = 'CSV file selected'
        self.get_csv_popup = BrowseToItem(msg_to_log, select_folder=select_folder)
        self.get_csv_popup.build()

    def select_save_folder(self, instance, select_folder=True):
        msg_to_log = 'Save folder selected'
        self.save_folder_popup = BrowseToItem(msg_to_log, select_folder=select_folder)
        self.save_folder_popup.build()

    def build_csv(self, instance):
        try:
            image_files = self.img_folder_popup.return_selection()
            csv_build = BuildCSV(image_files)
            filename = join(self.save_folder_popup.return_selection()[0], self.csv_filename.text)
            self.csv_file = csv_build.build(filename=filename, scale=self.scale.text)

            # log the file selection
            msg_to_log = 'CSV file selected'
            current_text = self.window.ids['error_text'].text
            new_text = f'{current_text}.\n{msg_to_log}: {self.csv_file}'
            self.window.ids['error_text'].text = new_text

        except AttributeError as e:
            action_to_take = 'Image folder not selected, perform Step 1'
            self.error_msg(e, action_to_take)

        except IndexError as e:
            action_to_take = 'Check if you have images in the folder you selected or select a new folder'
            self.error_msg(e, action_to_take)

    def create_final_document(self, instance):

        current_text = self.window.ids['error_text'].text
        new_text = f'{current_text}.\nRunning scaling....'
        self.window.ids['error_text'].text = new_text

        if self.csv_file is None:
            self.csv_file = self.get_csv_popup.return_selection()[0]

        try:
            assert isinstance(self.csv_file, str)
        except (AttributeError, AssertionError) as e:
            action_to_take = 'Did you select the right CSV file?'
            self.error_msg(e, action_to_take)
        try:
            assert isinstance(self.save_folder_popup.return_selection()[0], str)

            csv_file = self.csv_file  # csv file name
            save_location = self.save_folder_popup.return_selection()[0]  # base location to save the docx output
            docx_filename = self.doc_filename.text  # file name used for saving the docx

            scaler = Scaler(csv_file=csv_file, save_location=save_location, docx_filename=docx_filename)
            scaler.run_scaler()

            current_text = self.window.ids['error_text'].text
            new_text = f'{current_text}.\nFinished scaling successfully. ' \
                       f'Your file can be found at {join(save_location, docx_filename)}.docx'
            self.window.ids['error_text'].text = new_text

        except (AttributeError, AssertionError) as e:
            action_to_take = 'Did you select a save folder?'
            self.error_msg(e, action_to_take)

    def error_msg(self, e, action_to_take):
        grid = GridLayout()
        grid.cols = 1
        grid.add_widget(Label(text=action_to_take))
        dismiss_button = Button(text='Okay')
        grid.add_widget(dismiss_button)
        error_popup = Popup(content=grid, title=str(e), auto_dismiss=False, size_hint=(None, None), size=(600, 200))
        error_popup.open()
        dismiss_button.bind(on_press=error_popup.dismiss)
        return error_popup


if __name__ == '__main__':
    MainView().run()
