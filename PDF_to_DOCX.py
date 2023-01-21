# import all required packages
import os.path
from tkinter import ttk, Tk, messagebox
import tkinter.filedialog as fd
from pdf2docx import Converter


def open_pdf_file():
    # it will ask for a pdf file to open it
    file_name = fd.askopenfilename()

    # get the file name and split the file name from its path
    file_path = os.path.basename(file_name).split('/')[-1]

    # change the file type
    change_file_type = file_path.replace('.pdf', '.docx')

    pdf_file = file_name
    docx_file = change_file_type

    cv = Converter(pdf_file)
    cv.convert(docx_file)
    cv.close()

    messagebox.showinfo('Done', 'Successfully Converted...\n\nFor more information you can contact me on this number:'
                                ' 0797822830')


root = Tk()
root.title('PDF2Word Converter v1.0')

label_title = ttk.Label(root, text='Offline PDF_to_Word Converter!')
label_open_pdf = ttk.Label(root, text='Open PDF: ')
label_open_word = ttk.Label(root, text='Open Word: ')
label_Developer = ttk.Label(root, text='Developer: Behrooz Sharify')

open_pdf_button = ttk.Button(root, text='Open', command=open_pdf_file)

label_title.grid(row=0, column=1, columnspan=2)
label_open_pdf.grid(row=2, column=0)
label_Developer.grid(row=4, column=1, columnspan=2)
open_pdf_button.grid(row=2, column=3)

root.geometry('315x80')
root.resizable(False, False)
root.mainloop()
