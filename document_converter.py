from flask import Flask, render_template, request
from werkzeug.utils import secure_filename, send_file
from pdf2docx import Converter
from pdf2image import convert_from_path
from docx2pdf import convert
from comtypes import client
import pandas as pd
from fpdf import FPDF
import PIL
import img2pdf
import shutil
import os
import tabula
import pythoncom
import PyPDF2


directory = os.getcwd()
app = Flask(__name__)
app.config["UPLOAD_FILE"] = directory+"\\static\\uploaded file"


@app.route("/", methods=["GET", "POST"])
def home():
    title = "NecyTools | Online Document Converter"
    return render_template("home.html", title= title)


@app.route("/img2pdf", methods=["GET", "POST"])
def img_pdf():
    title = "IMAGE TO PDF"
    name = Download = ""
    if request.method == "POST":
        file = request.files["img"]
        name = file.filename
        file.save(os.path.join(app.config["UPLOAD_FILE"], secure_filename(name)))
        if " " in name:
            name = name.replace(" ", "_")
        if "(" in name:
            name = name.replace("(", "")
        if ")" in name:
            name = name.replace(")", "")
        img_path = f"{directory}\\static\\uploaded file\\" + name
        pdf_path = (
            f"{directory}\\static\\uploaded file\\"
            + name[:-4]
            + " converted.pdf"
        )
        image = PIL.Image.open(img_path)
        pdf_bytes = img2pdf.convert(image.filename)
        file = open(pdf_path, "wb")
        file.write(pdf_bytes)
        image.close()
        file.close()
        os.remove(img_path)
        Download = "Download"
    if request.method == "GET" and "/img2pdf?Download" in request.full_path:
        pdfname = request.full_path
        pdf_path = (
            f"{directory}\\static\\uploaded file\\"
            + pdfname[18:-4]
            + " converted.pdf"
        )
        return send_file(
            pdf_path, as_attachment=True, environ=request.environ
        )  # os.remove(pdf_path)#os.remove not working in vs code it's throw an erroe only works on pythonanywhere site
    return render_template("converter.html", Download=Download, name=name, title=title)


@app.route("/pdf2img", methods=["GET", "POST"])
def pdf_img():
    title="PDF TO IMAGE"
    name = Download = ""
    if request.method == "POST":
        file = request.files["pdf"]
        name = file.filename
        file.save(os.path.join(app.config["UPLOAD_FILE"], secure_filename(name)))
        if " " in name:
            name = name.replace(" ", "_")
        if "(" in name:
            name = name.replace("(", "")
        if ")" in name:
            name = name.replace(")", "")
        pdf_path = f"{directory}\\static\\uploaded file\\" + name
        img_path = (
            f"{directory}\\static\\uploaded file\\" + name[:-4] + " converted"
        )
        images = convert_from_path(
            pdf_path,
            500,
            poppler_path=r"C:\\Users\\abhis\Downloads\\poppler-0.68.0\bin",
        )
        if len(images) > 1:
            os.mkdir(f"{directory}\\static\\uploaded file\\"+name[:-4])
            img_path = f"{directory}\\static\\uploaded file\\{name[:-4]}\\{name[:-4]} converted"
        for i in range(len(images)):
            images[i].save(img_path + str(i) + ".jpg", "JPEG")
        file.close()
        os.remove(pdf_path)
        Download = "Download"
    if request.method == "GET" and "/pdf2img?Download" in request.full_path:
        imgname = request.full_path
        if os.path.isdir(f"{directory}\\static\\uploaded file\\"+imgname[18:-4]):
            shutil.make_archive(f"{directory}\\static\\uploaded file\\"+imgname[18:-4], 'zip', "{directory}\\static\\uploaded file\\"+imgname[18:-4])
            img_path = f"{directory}\\static\\uploaded file\\{imgname[18:-4]}.zip"
            shutil.rmtree(f"{directory}\\static\\uploaded file\\{imgname[18:-4]}")
        else:
            img_path = (
                f"{directory}\\static\\uploaded file\\"
                + imgname[18:-4]
                + " converted0.jpg"
            )
            print("hello")
        print(imgname)
        imgname =""
        return send_file(img_path, as_attachment=True, environ=request.environ)
    return render_template("pdf2img.html", Download=Download, name=name, title=title)


@app.route("/pdf2wrd", methods=["GET", "POST"])
def pdf_docx():
    title="PDF TO MS WORD"
    name = Download = ""
    if request.method == "POST":
        file = request.files["pdf"]
        name = file.filename
        file.save(os.path.join(app.config["UPLOAD_FILE"], secure_filename(name)))
        if " " in name:
            name = name.replace(" ", "_")
        if "(" in name:
            name = name.replace("(", "")
        if ")" in name:
            name = name.replace(")", "")
        pdf_path = f"{directory}\\static\\uploaded file\\" + name
        word_path = (
            f"{directory}\\static\\uploaded file\\"
            + name[:-4]
            + " converted.docx"
        )
        cv = Converter(pdf_path)
        cv.convert(word_path, start=0, end=None)
        cv.close()
        file.close()
        os.remove(pdf_path)
        Download = "Download"
    if request.method == "GET" and "/pdf2wrd?Download" in request.full_path:
        wordname = request.full_path
        word_path = (
            f"{directory}\\static\\uploaded file\\"
            + wordname[18:-4]
            + " converted.docx"
        )
        return send_file(
            word_path, as_attachment=True, environ=request.environ
        )  # os.remove(pdf_path)#os.remove not working in vs code it's throw an erroe only works on pythonanywhere site
    return render_template("pdf2wrd.html", Download=Download, name=name, title=title)


@app.route("/wrd2pdf", methods=["GET", "POST"])
def docx_pdf():
    title="WORD TO PDF"
    name = Download = ""
    if request.method == "POST":
        file = request.files["word"]
        name = file.filename
        file.save(os.path.join(app.config["UPLOAD_FILE"], secure_filename(name)))
        if " " in name:
            name = name.replace(" ", "_")
        if "(" in name:
            name = name.replace("(", "")
        if ")" in name:
            name = name.replace(")", "")
        word_path = f"{directory}\\static\\uploaded file\\" + name
        pdf_path = (
            f"{directory}\\static\\uploaded file\\"
            + name[:-4]
            + " converted.pdf"
        )
        pythoncom.CoInitialize()
        convert(word_path)
        file.close()
        os.remove(word_path)
        Download = "Download"
    if request.method == "GET" and "/wrd2pdf?Download" in request.full_path:
        pdfname = request.full_path
        pdf_path = (
            f"{directory}\\static\\uploaded file\\"
            + pdfname[18:-4]
            + " converted.pdf"
        )
        return send_file(
            pdf_path, as_attachment=True, environ=request.environ
        )  # os.remove(pdf_path)#os.remove not working in vs code it's throw an erroe only works on pythonanywhere site
    return render_template("wrd2pdf.html", Download=Download, name=name, title=title)


@app.route("/mergepdf", methods=["GET", "POST"])
def merge_pdf():
    title="MERGE PDF"
    Download = ""
    if request.method == "POST":
        file1 = request.files["pdf1"]
        file2 = request.files["pdf2"]
        name1 = file1.filename
        name2 = file2.filename
        file1.save(os.path.join(app.config["UPLOAD_FILE"], secure_filename(name1)))
        file2.save(os.path.join(app.config["UPLOAD_FILE"], secure_filename(name2)))
        if " " in name1:
            name1 = name1.replace(" ", "_")
        if "(" in name1:
            name1 = name1.replace("(", "")
        if ")" in name1:
            name1 = name1.replace(")", "")
        if " " in name2:
            name2 = name2.replace(" ", "_")
        if "(" in name2:
            name2 = name2.replace("(", "")
        if ")" in name2:
            name2 = name2.replace(")", "")
        pdf_file1 = f"{directory}\\static\\uploaded file\\" + name1
        pdf_file2 = f"{directory}\\static\\uploaded file\\" + name2
        Merge_pdf = f"{directory}\\static\\uploaded file\\Merge_files.pdf"
        pdf1File = open(pdf_file1, "rb")
        pdf2File = open(pdf_file2, "rb")
        pdf1Reader = PyPDF2.PdfFileReader(pdf1File)
        pdf2Reader = PyPDF2.PdfFileReader(pdf2File)
        pdfWriter = PyPDF2.PdfFileWriter()
        for pageNum in range(pdf1Reader.numPages):
            pageObj = pdf1Reader.getPage(pageNum)
            pdfWriter.addPage(pageObj)

        for pageNum in range(pdf2Reader.numPages):
            pageObj = pdf2Reader.getPage(pageNum)
            pdfWriter.addPage(pageObj)

        pdfOutputFile = open(Merge_pdf, "wb")
        pdfWriter.write(pdfOutputFile)

        pdfOutputFile.close()
        pdf1File.close()
        pdf2File.close()
        file1.close()
        file2.close()
        os.remove(pdf_file1)
        os.remove(pdf_file2)
        Download = "Download"
    if request.method == "GET" and "/mergepdf?Download" in request.full_path:
        # pdfname = request.full_path
        Merge_pdf = f"{directory}\\static\\uploaded file\\Merge_files.pdf"
        return send_file(
            Merge_pdf, as_attachment=True, environ=request.environ
        )  # os.remove(pdf_path)#os.remove not working in vs code it's throw an erroe only works on pythonanywhere site
    return render_template("merge_pdf.html", Download=Download, title=title)


@app.route("/cmpimg", methods=["GET", "POST"])
def compress():
    title="REDUSE IMAGE SIZE"
    name = Download = ""
    if request.method == "POST":
        file = request.files["img"]
        file_type = file.content_type
        range = request.values["range"]
        name = file.filename
        file.save(os.path.join(app.config["UPLOAD_FILE"], secure_filename(name)))
        if " " in name:
            name = name.replace(" ", "_")
        if "(" in name:
            name = name.replace("(", "")
        if ")" in name:
            name = name.replace(")", "")
        if "jpeg" in file_type: file_type = "jpg"
        elif "png" in file_type: file_type = "png"
        img_path = f"{directory}\\static\\uploaded file\\" + name
        cmp_path = (
            f"{directory}\\static\\uploaded file\\"
            + name[:-4]
            + " resized."+file_type
        )
        rwp = rhp = int(range)
        foo = PIL.Image.open(img_path)
        w, h = foo.size
        w, h = int(w - (w*(rwp/100))), int(h - (h*(rhp/100)))
        foo = foo.resize((w,h),PIL.Image.ANTIALIAS)
        foo.save(cmp_path,optimize=True,quality=95)
        file.close()
        os.remove(img_path)
        Download="Download"
    if request.method == "GET" and "/cmpimg?Download" in request.full_path:
        cmpname = request.full_path
        img_path = (
            f"{directory}\\static\\uploaded file\\"
            + cmpname[17:-4]
            + " resized."+cmpname[-3:]
        )
        return send_file(
            img_path, as_attachment=True, environ=request.environ
        )
    return render_template("cmpimg.html", Download=Download, name= name, title=title)

@app.route("/cmppdf", methods=["GET", "POST"])
def pdf_compress():
   pass

@app.route("/exl2pdf", methods=["GET", "POST"])
def xlsx_pdf():
    title="EXCEL TO PDF"
    name = Download = ""
    if request.method == "POST":
        file = request.files["excel"]
        file_type = file.mimetype
        name = file.filename
        file.save(os.path.join(app.config["UPLOAD_FILE"], secure_filename(name)))
        if " " in name:
            name = name.replace(" ", "_")
        if "(" in name:
            name = name.replace("(", "")
        if ")" in name:
            name = name.replace(")", "")
        xlsx_path = f"{directory}\\static\\uploaded file\\" + name
        pdf_path = (
            f"{directory}\\static\\uploaded file\\" + name[:-4] + " converted.pdf"
        )
        if 'csv' in file_type:df = pd.read_csv(xlsx_path)
        else: df = pd.read_excel(xlsx_path)
        data = df.to_string()
        f = open('excel.txt', 'w', encoding="utf-8")
        f = f.write(data)
        pdf = FPDF()      
        pdf.add_page()   
        pdf.set_font("Arial", size = 10)  
        f = open("excel.txt", "r")  
        for x in f:
            pdf.cell(200, 10, txt = x, ln = 1, align = 'L')   
        pdf.output(pdf_path)
        f.close()
        file.close()
        os.remove('excel.txt')
        os.remove(xlsx_path)
        Download="Download"
    if request.method == "GET" and "/exl2pdf?Download" in request.full_path:
        name = request.full_path
        pdf_path = (
            f"{directory}\\static\\uploaded file\\"
            + name[17:-4]
            + " converted.pdf" 
        )
        return send_file(
            pdf_path, as_attachment=True, environ=request.environ
        )
    return render_template("xlsx2pdf.html", Download=Download, name = name, title=title)

@app.route("/pdf2excel", methods=["GET", "POST"])
def pdf_excel():
    title="PDF TO EXCEL"
    name = Download = ""
    if request.method == "POST":
        file = request.files["pdf"]
        file_type = file.mimetype
        name = file.filename
        file.save(os.path.join(app.config["UPLOAD_FILE"], secure_filename(name)))
        if " " in name:
            name = name.replace(" ", "_")
        if "(" in name:
            name = name.replace("(", "")
        if ")" in name:
            name = name.replace(")", "")
        pdf_path = f"{directory}\\static\\uploaded file\\" + name
        xlsx_path = (
            f"{directory}\\static\\uploaded file\\" + name[:-4] + " converted.csv"
        )
        tabula.convert_into(pdf_path, xlsx_path, output_format="csv", pages='all')
        file.close()
        os.remove(pdf_path)
        Download="Download"
    if request.method == "GET" and "/pdf2excel?Download" in request.full_path:
        name = request.full_path
        excel_path = (
            f"{directory}\\static\\uploaded file\\"
            + name[20:-4]
            + " converted.csv" 
        )
        return send_file(
                excel_path, as_attachment=True, environ=request.environ
            )
    return render_template("pdf2xlsx.html", Download=Download, name = name, title=title)

@app.route("/ppt2pdf", methods=["GET", "POST"])
def ppt_pdf():
    title="PPT TO PDF"
    name = Download = ""
    powerpoint = client.CreateObject("Powerpoint.Application", pythoncom.CoInitializeEx(0x0))
    if request.method == "POST":
        file = request.files["ppt"]
        name = file.filename
        file.save(os.path.join(app.config["UPLOAD_FILE"], secure_filename(name)))
        if " " in name:
            name = name.replace(" ", "_")
        if "(" in name:
            name = name.replace("(", "")
        if ")" in name:
            name = name.replace(")", "")
        ppt_path = f"{directory}\\static\\uploaded file\\" + name
        pdf_path = (
            f"{directory}\\static\\uploaded file\\" + name[:-5] + " converted.pdf"
        )
        powerpoint.Visible = 1
        slides = powerpoint.Presentations.Open(ppt_path)
        slides.SaveAs(pdf_path, 32)
        slides.Close()
        file.close()
        os.remove(ppt_path)
        Download="Download"
    if request.method == "GET" and "/ppt2pdf?Download" in request.full_path:
        name = request.full_path
        pdf_path = (
            f"{directory}\\static\\uploaded file\\"
            + name[18:-5]
            + " converted.pdf" 
        )
        return send_file(
                pdf_path, as_attachment=True, environ=request.environ
            )
    return render_template("ppt2pdf.html", Download=Download, name = name, title=title)

@app.route("/pdf2ppt", methods=["GET", "POST"])
def pdf_ppt():
    pass

@app.route("/encrptpdf", methods=["GET", "POST"])
def encrypt_pdf():
    title="ENCRYPT PDF"
    name = Download = ""
    if request.method == "POST":
        file = request.files["pdf"]
        pswrd1 = request.values["pass1"]
        pswrd2 = request.values["pass2"]
        if pswrd1 != pswrd2: 
            Download = "Password is not same"
            return Download
        else:
            name = file.filename
            file.save(os.path.join(app.config["UPLOAD_FILE"], secure_filename(name)))
            if " " in name:
                name = name.replace(" ", "_")
            if "(" in name:
                name = name.replace("(", "")
            if ")" in name:
                name = name.replace(")", "")
            pdf_path = f"{directory}\\static\\uploaded file\\" + name
            enc_path = (
                f"{directory}\\static\\uploaded file\\" + name[:-4] + " encrypted.pdf"
            )
            with open(pdf_path, "rb") as in_file:
                input_pdf = PyPDF2.PdfFileReader(in_file)

                output_pdf = PyPDF2.PdfFileWriter()
                output_pdf.appendPagesFromReader(input_pdf)
                output_pdf.encrypt(pswrd2)

                with open(enc_path, "wb") as out_file:
                    output_pdf.write(out_file)
            file.close()
            os.remove(pdf_path)
            Download = "Download"
    if request.method == "GET" and "/encrptpdf?Download" in request.full_path:
        name = request.full_path
        pdf_path = (
            f"{directory}\\static\\uploaded file\\"
            + name[20:-4]
            + " encrypted.pdf" 
        )
        return send_file(
                pdf_path, as_attachment=True, environ=request.environ
            )
    return render_template("encrptpdf.html", Download=Download, name = name, title=title)

@app.route("/decrptpdf", methods=["GET", "POST"])
def decrypt_pdf():
    title="DECRYPT PDF"
    name = Download = ""
    if request.method == "POST":
        file = request.files["pdf"]
        pswrd = request.values["pass"]
        name = file.filename
        file.save(os.path.join(app.config["UPLOAD_FILE"], secure_filename(name)))
        if " " in name:
            name = name.replace(" ", "_")
        if "(" in name:
            name = name.replace("(", "")
        if ")" in name:
            name = name.replace(")", "")
        pdf_path = f"{directory}\\static\\uploaded file\\" + name
        dec_path = (
            f"{directory}\\static\\uploaded file\\" + name[:-4] + " decrypted.pdf"            )
        pagenum = 0
        result = PyPDF2.PdfFileWriter()
        file1 = PyPDF2.PdfFileReader(pdf_path)
        password = pswrd
        if file1.isEncrypted:
            file1.decrypt(password)
            try:  
                for i in range(999):
                    pagenum = i
                    file1.getPage(i)
            except IndexError:
                    for j in range(pagenum):
                        pages = file1.getPage(j)
                        result.addPage(pages)
                    with open(dec_path,'wb') as f:
                        result.write(f)
            Download = "Download"
        else:
            Download = 'File is not encrypted'        
        file.close()
        os.remove(pdf_path)
    if request.method == "GET" and "/decrptpdf?Download" in request.full_path:
        name = request.full_path
        pdf_path = (
            f"{directory}\\static\\uploaded file\\"
            + name[20:-4]
            + " decrypted.pdf" 
        )
        return send_file(
                pdf_path, as_attachment=True, environ=request.environ
            )
    return render_template("decrptpdf.html", Download=Download, name = name, title=title)

@app.route("/del_pages", methods=["GET", "POST"])
def del_pages():
    title="DELETE PAGES"
    Warning = name = Download = ""
    page_num = []
    result = True
    if request.method == "POST":
        file = request.files["pdf"]
        page_to_delete = request.values["pagenum"]
        page_to_delete = page_to_delete.split(',')
        print(page_to_delete)
        try:
            for i in page_to_delete:
                result = result and (i.isdigit())
                page_num.append(int(i)-1)
                print(result)
            if not result:
                Warning = "Please Enter Only Numeric Values"
            else:
                name = file.filename
                file.save(os.path.join(app.config["UPLOAD_FILE"], secure_filename(name)))
                if " " in name:
                    name = name.replace(" ", "_")
                if "(" in name:
                    name = name.replace("(", "")
                if ")" in name:
                    name = name.replace(")", "")
                pdf_path = f"{directory}\\static\\uploaded file\\" + name
                new_path = (
                    f"{directory}\\static\\uploaded file\\" + name[:-4] + " modified.pdf"            )
                infile = PyPDF2.PdfFileReader(pdf_path, 'rb')
                output = PyPDF2.PdfFileWriter()
                for i in range(infile.numPages):
                    if i in page_num:
                        print("condition",i)
                        continue
                    else:
                        print(i)
                        p = infile.getPage(int(i))
                        output.addPage(p)
                with open(new_path, 'wb') as f:
                    output.write(f)
                file.close()
                os.remove(pdf_path)
                Warning = ""
                Download = "Download"
        except IndexError:
            Warning = "Please Enter Valid Pages Number"
    if request.method == "GET" and "/del_pages?Download" in request.full_path:
        name = request.full_path
        pdf_path = (
            f"{directory}\\static\\uploaded file\\"
            + name[20:-4]
            + " modified.pdf" 
        )
        return send_file(
                pdf_path, as_attachment=True, environ=request.environ
            )
    return render_template('delete_pages.html', Download=Download, name = name, Warning=Warning, title=title)
app.run(debug=True)