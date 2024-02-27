# Python-GUI-PDF_converter

A project to create a automation tool allowing me to convert DOCX files to PDF files with ease. I used Python for easy integration with third-party libraires and quick development. The application allows a user to select a DOCX file through a graphical user interface (GUI), and then convert it to PDF with simple button click.

- Python with object oriented programming (OOP) was used to structure the program.
- Created the GUI interface with Tkinter, allow cross-platform compatibility.
- Used Ttkbootstrap enhancing the application's visual aesthetics.
- Employed pyinstaller to bundle the applicaiton into a standalone executable.

## Requirements

- Python 3.10 or later (Tested with only 3.10)
- Microsoft Visual Studio Code or your preffered code editor for Python
- Program tested on Windows.

## How it works and looks:

## How to run it?

1. Clone the project or download the all the folders.
2. Launch the application without executing the code by running the _pdfConverter.exe_ in the _dist_ folder.
3. Otherwise, open the code in your preffered code editor.
4. Open the terminal, and using pip install the missing packages such as docx2pdf package and confirm they are installed.
   - To confirm they are installed, in the terminal, type enter the command python. This will allow you to execute the python code, and simply import the packages, if there are no errors then the package is installed.
5. Execute the code. To execute from the terminal, ensure you are in the directory of the file. Enter the following: _python ./[name_of_file].py_
6. To create a executable of the program, install pyinstaller. May need to add to the PATH in system variables. Then in the terminal: _pyinstaller ./[name_of_file].py_. The executable will be located in the dist folder.
