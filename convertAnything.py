
from tkinter import *
from tkinter import filedialog as fd
from pdf2image import convert_from_path
from docx2pdf import convert
from PIL import Image
from tkinter import messagebox
from moviepy.editor import *
# import PyPDF2
from pdf2docx import Converter

root = Tk()

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


button1 = Button(root, text="JPG & JPEG to PNG", width=20, height=2, bg="green",
                 fg="white", font=("helvetica", 12, "bold"), command=jpgToPng)

button1.place(x=120, y=120)

button2 = Button(root, text="PNG to JPG", width=20, height=2, bg="green",
                 fg="white", font=("helvetica", 12, "bold"), command=pngToJpg)

button2.place(x=120, y=220)
mp4Button = Button(root, text="MP4 to MP3", width=20, height=2, bg="green",
                   fg="white", font=("helvetica", 12, "bold"), command=mp4ToMp3)
mp4Button.place(x=120, y=320)

wordPdf = Button(root, text="Word to PDF", width=20, height=2, bg="green",
                 fg="white", font=("helvetica", 12, "bold"), command=wordToPdf)
wordPdf.place(x=120, y=420)

pdfImage = Button(root, text="PDF to JPG", width=20, height=2, bg="green",
                  fg="white", font=("helvetica", 12, "bold"), command=pdfToJpg)
pdfImage.place(x=120, y=520)
pdfWord = Button(root, text="PDF to WORD", width=20, height=2, bg="green",
                 fg="white", font=("helvetica", 12, "bold"), command=pdfToWord)
pdfWord.place(x=120, y=620)
root.geometry("800x800+400+200")

root.mainloop()