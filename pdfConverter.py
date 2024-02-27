# GUI Libraries
import tkinter as tk
from tkinter import filedialog
import ttkbootstrap as ttk

# For converting the docx to pdf
from docx2pdf import convert

# Accessing the files
import os
import time


# Class for aGUI application to convert DOCX files to PDFs
class DocxToPdfConverterApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("DOCX to PDF Converter") # Set window title
        self.geometry("600x500") # Set window size
        
        # Label to display application purpose
        self.label = ttk.Label(self, text="Convert docx files to PDFs")
        self.label.pack(pady=30)
        self.label.config(font=("Arial", 20, "bold")) # Set label font
        
        # A frame to organise widgets
        self.frame = ttk.Frame(self)
        self.frame.pack(pady=15, padx=10, fill="x")
        
        # Button to select DOCX file
        self.select_button = ttk.Button(self.frame, text="Select DOCX File", command=self.select_file)
        self.select_button.pack()
        
        # Label to display selected file path
        self.file_path_label = ttk.Label(self.frame, text="")
        self.file_path_label.pack()
        
        # Button to start conversion
        self.convert_button = ttk.Button(self.frame, text="Convert to PDF", command=self.convert_to_pdf, state=tk.DISABLED)
        self.convert_button.pack()
        
        # Progress bar to indicate conversion progress
        self.progress_bar = ttk.Progressbar(self.frame, orient="horizontal", mode="determinate", length=280)
        self.progress_bar.pack()
        
        # Label to indicate conversion completion
        self.done_label = ttk.Label(self.frame, text="")
        self.done_label.pack()
    
    # Method to handle file selection
    def select_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Word files", "*.docx")])
        if file_path:
            self.convert_button.config(state=tk.NORMAL)
            self.file_path_label.config(text=file_path)
    
    # Method to handle conversion completion       
    def on_conversion_complete(self):
        print("Conversion completed successfully!")
    
    # Function to convert DOCX file to PDF
    def convert_to_pdf(self):
        self.done_label.config(text="")
        docx_file = self.file_path_label.cget("text")
        self.convert_button.config(state=tk.DISABLED)
        
        # Generate PDF file path
        base_name, ext = os.path.splitext(os.path.basename(docx_file))
        pdf_file = os.path.join(os.path.dirname(docx_file), f"{base_name}.pdf")
        count = 0
        fileNameExists = False
        
        # Ensure unique file name
        while os.path.exists(pdf_file):
            count += 1
            pdf_file = os.path.join(os.path.dirname(docx_file), f"{base_name}-{count}.pdf")
           
        if(count > 0):
            fileNameExists = True
        
        # Perform conversion  
        convert(docx_file, pdf_file)
        
        # Update progress bar
        for progress in range(101):
            self.progress_bar["value"] = progress
            self.update_idletasks()
            self.after(50)
        
        # Reset progress bar and display converion completion message
        time.sleep(3)
        self.on_conversion_complete()
        self.progress_bar["value"] = 0 
        self.file_path_label.config(text="")
        if(fileNameExists):
            self.done_label.config(text="File has been converted. Name of file: " + base_name + "-" + str(count) + ".pdf")
        else:
            self.done_label.config(text="File has been converted. Name of file: " + base_name + ".pdf")

        
if __name__ == "__main__":
    app = DocxToPdfConverterApp()
    app.mainloop() # Start the application