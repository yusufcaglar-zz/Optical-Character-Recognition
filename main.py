import os
import tkinter as tk
from tkinter.filedialog import askopenfile
import PIL
import cv2
import numpy as np
import pytesseract
from PIL import Image, ImageTk
import docx
import fpdf


# Global Variables
img = ""
text = ""
cvs_image = ""
cvs_image2 = ""
image = ""

# Creating Window
window = tk.Tk()
window.title("Optical Character Recognition")
window.iconbitmap("yildiz_logo.ico")
window.configure(background="#D9D9D9")
window.minsize(900, 570)


def browse():
    def pdf_to_img(pdf_file):
        # PDF to Image
        pdftoppmpath = r"C:\Users\sturu\PycharmProjects\imageProject\poppler-0.68.0\bin"

        from pdf2image import convert_from_path
        pages = convert_from_path(pdf_path=pdf_file, poppler_path=pdftoppmpath)
        pages[0].save('out.jpg', 'JPEG')

        return "out.jpg"

    global img
    global cvs_image

    basedir = os.path.dirname(os.path.abspath(__file__))

    file = askopenfile(
        initialdir=basedir,
        title="Choose the image file",
        filetypes=[("Image or PDF Files", "*.pdf *.jpg *.jpeg *.png *.bmp *.tmp")]
    )

    if file is not None:
        name, extension = os.path.splitext(file.name)
        if extension == ".pdf" or extension == ".PDF":
            img = pdf_to_img(file.name)
        else:
            img = file.name

        cvs_image = ImageTk.PhotoImage(image=Image.open(img))
        canvas[0].create_image(200, 200, image=cvs_image)


def convert():
    global img
    global text
    global cvs_image2
    global image
    global textarea

    if img != "":
        # OCR
        pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

        image = cv2.imread(img)
        image = cv2.cvtColor(image, cv2.COLOR_RGB2GRAY)

        boxes = pytesseract.image_to_data(image)

        image = cv2.cvtColor(image, cv2.COLOR_GRAY2RGB)

        text = pytesseract.image_to_string(img)

        for x, b in enumerate(boxes.splitlines()):
            if x != 0:
                b = b.split()
                if len(b) == 12:
                    x, y, w, h = int(b[6]), int(b[7]), int(b[8]), int(b[9])
                    cv2.rectangle(image, (x, y), (w + x, h + y), (0, 0, 255), 1)
        # OCR

        cvs_image2 = PIL.ImageTk.PhotoImage(image=PIL.Image.fromarray(image))
        canvas[1].create_image(200, 200, image=cvs_image2)

        textarea.configure(state="normal")
        textarea.delete("1.0", tk.END)
        textarea.insert(tk.INSERT, pytesseract.image_to_string(img))
        textarea.configure(state="disabled")


def open_photo(num):
    global img
    global cvs_image
    global cvs_image2

    if num == 0 and img != "":
        cv2.imshow("input", cv2.imread(img))
    elif num == 1 and image != "":
        cv2.imshow("output", image)


def export():
    # string to pdf
    global cvs_image2
    global image

    if image != "":
        pdf = fpdf.FPDF()
        pdf.add_page()
        pdf.add_font('Arial Unicode MS', '', 'arial-unicode-ms.ttf', uni=True)
        pdf.set_font('Arial Unicode MS', '', 16)
        pdf.multi_cell(40, 10, text)
        pdf.output('exported.pdf', 'F')

        # string to docx
        document = docx.Document()
        document.add_paragraph(text)
        document.save('exported.docx')

        # image to pdf
        cv2.imwrite('exported.jpg', image)

        # open output files
        os.startfile("exported.pdf")
        os.startfile("exported.docx")
        os.startfile("exported.jpg")


img_out = np.zeros([100, 100, 3], dtype=np.uint8)

window.columnconfigure(0, weight=1)
window.columnconfigure(1, weight=1)
window.columnconfigure(2, weight=1)

# Canvases
canvas = (tk.Canvas(window, width=400, height=400, bg="#F2F2F2"), tk.Canvas(window, width=400, height=400, bg="#F2F2F2"))
canvas[0].config(highlightthickness=2, highlightbackground="#8C8C8C")
canvas[1].config(highlightthickness=2, highlightbackground="#8C8C8C")

canvas[0].grid(row=0, column=0, sticky=tk.NSEW, padx=2, pady=2)
canvas[1].grid(row=0, column=2, sticky=tk.NSEW, padx=2, pady=2)

canvas[0].bind("<Button-1>", lambda event: open_photo(0))
canvas[1].bind("<Button-1>", lambda event: open_photo(1))

# Buttons
button = (
    tk.Button(window, text=f'Import', font='Helvetica 12', command=browse),
    tk.Button(window, text=f'Export', font='Helvetica 12', command=export),
    tk.Button(window, text=f'Convert', font='Helvetica 12', command=convert))

button[0].grid(row=1, column=0, sticky=tk.EW, ipadx=2, ipady=2, padx=5, pady=5)
button[1].grid(row=1, column=2, sticky=tk.EW, ipadx=2, ipady=2, padx=5, pady=5)
button[2].grid(row=0, column=1, sticky=tk.EW, ipadx=2, ipady=2, padx=2, pady=2)

# Labels
instructions = "INSTRUCTIONS\n 1) Select a photo or pdf file\n 2) Click Convert button\n 3) Click export button to export files\n 4) Exported pdf, docx and jpg files will be on source file\n*You can also view images by clicking on them*"

label = (tk.Label(window, text=instructions, font='Helvetica 12', background="#D9D9D9"))
label.grid(row=2, column=0, sticky=tk.NSEW, padx=2, pady=2)

# Textarea
textarea = tk.Text(window, height=5, width=5)
textarea.grid(row=2, column=2, sticky=tk.NSEW, padx=2, pady=2)
scrollbar = tk.Scrollbar(window, orient='vertical', command=textarea.yview)
scrollbar.grid(row=2, column=3, sticky='ns')

window.mainloop()
