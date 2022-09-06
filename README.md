# ArtScaler
 Image scaling based on life size. The FileBrowser.py was copied from https://github.com/kivy-garden/filebrowser.



## Compiling

The app was designed with Kivy 2 and the idea was to make it into a transportable executable. Because of this aim, I suggest compiling the app by taking the following actions:

1) create a virtual environment and install packages in requirements.txt 

2) create an empty folder 

3) cd into that folder while in the virtual environment 

4) type: python -m PyInstaller --onefile --hiddenimport win32timezone --name ArtScaler Full\path\to\Main.py

5) open the ArtScaler.spec with a text editor

6) add to the top of the document: from kivy_deps import sdl2, glew  

7) go to EXE at the bottom of the document and replace with (change the folder path name for Tree(), though - don't include filename in path):

exe = EXE(
    pyz, Tree('Full\\path\\to\\folder\\where\Main.py\is\located'),
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    *[Tree(p) for p in (sdl2.dep_bins + glew.dep_bins)],
    name='ArtScaler',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

 8) and finally, run python -m PyInstaller ArtScaler.spec in the terminal. The transportable executable will be found in the *dist* folder. If you get "ModuleNotFoundError: No module named 'kivy_deps'", run python -m pip install kivy-deps.sdl2 and python -m pip install kivy-deps.glew in your terminal.

## Using the app:

```
The steps for the user:
    1) They create a folder with images that they wish to scale
    2) They select the folder where they want to save the intermediate CSV file and the final .docx
    3) They give a filename for the intermediate CSV file
    4) They give a filename for the final .docx
    5) They assign a scale (1 returning just the image).
    6) They can now press 'Prepare CSV', which populates the CSV file with the filenames from the folder 	        provided in step 1), as well as the scale.
    7) The user can now open the intermediate CSV file and input the life size of the images and modify the        scale if they want. They can also assign the size of the frame for the artwork, as well as the shape        of the frame. The shape keywords 'square', 'rectangle', and 'none' create similar frame as long as 		   frame size is provided. Likewise, keywords 'circle' and 'oval' produce similar frames as long as 	        frame size is provided.
    8) Once all the rows of interest are populated, the user can now produce the final scaled images.

Alternatively, if the user has a CSV file with all the necessary information:
    1) They need to make sure the following columns are present (note capital letters):
            Image_file, Scale, Artwork_life_size_height, Artwork_life_size_width, Artwork_frame_height, 				Artwork_frame_width, Frame_type
       The information for the first four columns must be populated if images are to be scaled. If frames 		   are also to be scaled, then the last three columns need to be populated. Image_file must include the 	   full path to each image.
    2) They can now skip steps 1), 3), 5), 6), and 7)
    3) Instead of going to step 8, they must press Choose CSV and find the pre-populated CSV.
    4) Now, run the program.
```
