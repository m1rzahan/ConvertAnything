
from tkinter import *
from tkinter import filedialog as fd
from docx import Document
import cv2
import pytesseract as pytesseract
from pdf2image import convert_from_path
from docx2pdf import convert
from PIL import Image
from tkinter import messagebox
from moviepy.editor import *
# import PyPDF2
from pdf2docx import Converter
import webbrowser
root = Tk()

def aboutLabel():
   webbrowser.open_new(r"https://www.mirzahan.dev")

menubar = Menu(root)
filemenu = Menu(menubar, tearoff=0)

helpmenu = Menu(menubar, tearoff=0)

#helpmenu.add_command(label="About...", command=aboutLabel)
menubar.add_command(label="About", command=aboutLabel)
root.title("Convert Anything")
def jpgToPng():
    global im1

    import_filename = fd.askopenfilename()
    if import_filename.endswith(".jpg") or import_filename.endswith(".jpeg"):

        im1 = Image.open(import_filename)
        export_filename = fd.asksaveasfilename(defaultextension=".png")
        im1.save(export_filename)

        # displaying the Messaging box with the Success
        messagebox.showinfo("success ", "your Image converted to Png")
    else:

        # if Image select is not with the Format of .jpg
        # then display the Error
        Label_2 = Label(root, text="Error!", width=20,
                        fg="red", font=("bold", 15))
        Label_2.place(x=80, y=280)
        messagebox.showerror("Fail!!", "Something Went Wrong...")


def pngToJpg():
    global im1
    import_filename = fd.askopenfilename()

    if import_filename.endswith(".png"):
        im1 = Image.open(import_filename)
        export_filename = fd.asksaveasfilename(defaultextension=".jpg")
        im1.save(export_filename)

        messagebox.showinfo("success ", "your Image converted to jpg ")
    else:
        Label_2 = Label(root, text="Error!", width=20,
                        fg="red", font=("bold", 15))
        Label_2.place(x=80, y=280)

        messagebox.showerror("Fail!!", "Something Went Wrong...")


def mp4ToMp3():
    import_filename = fd.askopenfilename()
    if import_filename.endswith(".mp4"):
        video = AudioFileClip(import_filename)
        export_filename = fd.asksaveasfilename(defaultextension=".mp3")
        video.write_audiofile(export_filename)

        messagebox.showinfo("success ", "your  mp4 video converted to mp3 ")
    else:
        Label_2 = Label(root, text="Error!", width=20,
                        fg="red", font=("bold", 15))
        Label_2.place(x=80, y=280)

        messagebox.showerror("Fail!!", "Something Went Wrong...")


def wordToPdf():
    import_filename = fd.askopenfilename()
    if import_filename.endswith(".docx"):

        convert(import_filename)
        messagebox.showinfo("success ", "your  docx file converted to pdf ")
    else:
        Label_2 = Label(root, text="Error!", width=20,
                        fg="red", font=("bold", 15))
        Label_2.place(x=80, y=280)

        messagebox.showerror("Fail!!", "Something Went Wrong...")


def pdfToWord():
    import_filename = fd.askopenfilename()
    if import_filename.endswith(".pdf"):
        cv = Converter(import_filename)
        messagebox.showinfo("Warning",
                            "Please create the docx file to save! You dont need to write .docx extension end of the filename")

        export_filename = fd.asksaveasfilename(defaultextension=".docx")
        messagebox.showinfo("Warning",
                            "The duration of the conversion process may vary depending on the number of pages. Please wait until the successful message is displayed.")
        cv.convert(export_filename)
        cv.close()

        messagebox.showinfo("success ", "your  docx file converted to pdf ")
    else:
        Label_2 = Label(root, text="Error!", width=20,
                        fg="red", font=("bold", 15))
        Label_2.place(x=80, y=280)

        messagebox.showerror("Fail!!", "Something Went Wrong...")


def pdfToJpg():
    import_filename = fd.askopenfilename()
    if import_filename.endswith(".pdf"):
        pages = convert_from_path(import_filename, 500)
        for page in pages:
            messagebox.showinfo("Be careful!! ", "Please create name and save  for each page")
            export_filename = fd.asksaveasfilename(defaultextension=".jpg")
            page.save(export_filename)

        messagebox.showinfo("success ", "your  pdf file converted to jpg ")
    else:
        Label_2 = Label(root, text="Error!", width=20,
                        fg="red", font=("bold", 15))
        Label_2.place(x=80, y=280)

        messagebox.showerror("Fail!!", "Something Went Wrong...")

def ocrCore(img):
    text= pytesseract.image_to_string(img)
    return text
def getGrayscale(image):
    return cv2.cvtColor(image,cv2.COLOR_BGR2GRAY)
def removeNoise(image):
    return cv2.medianBlur(image,5)
def thresholding(image):
    return cv2.threshold(image,0,255,cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1]
def imageToText():
    import_filename = fd.askopenfilename()
    if import_filename.endswith(".png"):
        img = cv2.imread(import_filename)
        img = getGrayscale(img)
        img = thresholding(img)
        img = removeNoise(img)
        text = ocrCore(img)
        print(text)
        newText = text.replace("\x0c","")
        document = Document()
        document.add_paragraph(newText)
        export_filename = fd.asksaveasfilename(defaultextension=".docx")
        document.save(export_filename)
jpgToPngButton = Button(root, text="JPG & JPEG to PNG", width=20, height=2, bg="blue",
                        fg="white", font=("helvetica", 12, "bold"), command=jpgToPng)

jpgToPngButton.place(x=120, y=120)

pngToJpgButton = Button(root, text="PNG to JPG", width=20, height=2, bg="blue",
                        fg="white", font=("helvetica", 12, "bold"), command=pngToJpg)

pngToJpgButton.place(x=420, y=120)

mp4Button = Button(root, text="MP4 to MP3", width=20, height=2, bg="blue",
                   fg="white", font=("helvetica", 12, "bold"), command=mp4ToMp3)
mp4Button.place(x=120, y=320)

wordPdf = Button(root, text="Word to PDF", width=20, height=2, bg="blue",
                 fg="white", font=("helvetica", 12, "bold"), command=wordToPdf)
wordPdf.place(x=420, y=320)

pdfImage = Button(root, text="PDF to JPG", width=20, height=2, bg="blue",
                  fg="white", font=("helvetica", 12, "bold"), command=pdfToJpg)
pdfImage.place(x=120, y=520)
pdfWord = Button(root, text="PDF to WORD", width=20, height=2, bg="blue",
                 fg="white", font=("helvetica", 12, "bold"), command=pdfToWord)
pdfWord.place(x=420, y=520)

imgToText = Button(root, text="Image to Text", width=20, height=2, bg="blue",
                 fg="white", font=("helvetica", 12, "bold"), command=imageToText)
imgToText.place(x=120, y=720)
root.geometry("800x800+400+200")
root.config(menu=menubar)
root.resizable(False, False)

root.mainloop()