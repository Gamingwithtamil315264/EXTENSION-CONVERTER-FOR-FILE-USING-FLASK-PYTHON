from flask import Flask,render_template,request,send_file
import os
from PIL import Image,ImageFile
from distutils.log import debug 
from PyPDF2 import PdfReader,PdfWriter
app=Flask(__name__)
ImageFile.LOAD_TRUNCATED_IMAGES = True
@app.route('/')
def index():
    return render_template('doc.html')
@app.route('/pdftoword',methods=["GET","POST"])
def pdftoword():
    if request.method=='POST':
        pdf=request.files['file']
        pdf_=os.path.splitext(pdf.filename)
        doc=pdf_[0]+".docx"
        pdf.save(doc)
        return send_file(doc, as_attachment=True)
@app.route('/wordtopdf',methods=["GET","POST"])
def wordtopdf():
    if request.method=='POST':
        pdf=request.files['file']
        pdf_=os.path.splitext(pdf.filename)
        doc=pdf_[0]+".pdf"
        pdf.save(doc)
        return send_file(doc, as_attachment=True)
@app.route('/pptxtopdf',methods=["GET","POST"])
def pptxtopdf():
    if request.method=='POST':
        pdf=request.files['file']
        pdf_=os.path.splitext(pdf.filename)
        doc=pdf_[0]+".pdf"
        pdf.save(doc)
        return send_file(doc, as_attachment=True)
@app.route('/exceltopdf',methods=["GET","POST"])
def exceltopdf():
    if request.method=='POST':
        pdf=request.files['file']
        pdf_=os.path.splitext(pdf.filename)
        doc=pdf_[0]+".pdf"
        pdf.save(doc)
        return send_file(doc, as_attachment=True)
@app.route('/compressimage',methods=["GET","POST"])
def compressimage():
    if request.method=='POST':
        pdf=request.files['file']
        pdf.save(pdf)
        image = Image.open(pdf)
        width, height = image.size
        new_size = (width//2, height//2)
        resized_image = image.resize(new_size)
        resized_image.save("compressed.jpg", optimize=True, quality=50)
        return send_file("compressed.jpg", as_attachment=True)
@app.route('/compresspdf',methods=["GET","POST"])
def compresspdf():
    if request.method=='POST':
        pdf=request.files['file']
        pdf.save(pdf.filename)
        print(pdf.filename)
        reader = PdfReader(pdf)
        writer = PdfWriter()
        for page in reader.pages:
            page.compress_content_streams()  
            writer.add_page(page)
        with open("compressed_file.pdf", "wb") as f:
            writer.write(f)
        return send_file("compressed_file.pdf", as_attachment=True)

if __name__=="__main__":
    app.run(debug=True)
#ssl_context='adhoc',
