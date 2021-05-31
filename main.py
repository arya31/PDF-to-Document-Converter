from pdf2jpg import pdf2jpg
from tkinter import *
from tkinter import messagebox
from tkinter import filedialog
import os
from docx import Document
from docx.shared import Inches
from natsort import natsorted
import shutil


def take_screenshot():
	print("Generating Screenshots")
	path_pdf = path1.get()
	pdf2jpg.convert_pdf2jpg(path_pdf,os.path.dirname(path_pdf), dpi=100, pages="ALL")
	print("Screenshots generated")
	create_docx(path_pdf)


def create_docx(path_pdf):
    print("Adding Screenshots to document")
    image_path = path_pdf + "_dir/"
    image_list = natsorted(os.listdir(image_path))
    docx_filename = path_pdf[:-3] + ".docx"

    root = Document()
    image_index = 0
    while (image_index < len(image_list)):
        print("Added image ", image_index)
        image = image_path + image_list[image_index]
        root.add_picture(image, height=Inches(8))
        image_index += 1
    root.save(docx_filename)
    shutil.rmtree(image_path)
    messagebox.showinfo("SUCCESS", "DONE")


if __name__ == "__main__":
	window = Tk()
	window.title("My First App")
	window.geometry("350x150")
	window.configure(bg='white')
	window.attributes('-alpha',0.90)
	path1 = StringVar()

	Label(window, text="Choose Pdf File", bg='white').place(x=0,y=30)
	Entry(window, width=30, bg='lightgrey',textvariable=path1).place(x=110,y=30)
	btn1 = Label(window, text="...", width=4, bg='white')
	btn1.place(x=300,y=30)
	btn1.bind("<Button>",lambda e: path1.set(filedialog.askopenfilename(initialdir=os.getcwd(),title="Select Pdf File")))
	
	Button(window, text="RUN", font="none 11", width=4, command=take_screenshot, bg='lightgreen').place(x=150,y=100)
	window.mainloop()