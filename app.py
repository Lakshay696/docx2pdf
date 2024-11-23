from flask import Flask, render_template, request, send_file
from werkzeug.utils import secure_filename
from docx import Document
from docx2pdf import convert
from PyPDF2 import PdfReader, PdfWriter
import os
import pythoncom  

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
PDF_FOLDER = 'pdfs'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(PDF_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

def extract_metadata(file_path):
    try:
        doc = Document(file_path)
        metadata = doc.core_properties
        return {
            "Title": metadata.title or "Unknown",
            "Author": metadata.author or "Unknown",
            "Created": metadata.created or "Unknown",
            "Modified": metadata.modified or "Unknown",
        }
    except Exception as e:
        return {"Error": str(e)}


def add_password_to_pdf(pdf_path, output_path, password):
    try:
        writer = PdfWriter()
        reader = PdfReader(pdf_path)
        for page in reader.pages:
            writer.add_page(page)
        writer.encrypt(password)
        with open(output_path, "wb") as output_file:
            writer.write(output_file)
    except Exception as e:
        raise Exception(f"Failed to add password protection: {str(e)}")


@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        if 'file' not in request.files:
            return "No file uploaded!", 400

        file = request.files['file']
        if file.filename == '':
            return "No selected file!", 400

        if file and file.filename.endswith('.docx'):
            try:
                filename = secure_filename(file.filename)
                file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                file.save(file_path)
                metadata = extract_metadata(file_path)

                pdf_path = os.path.join(PDF_FOLDER, f"{os.path.splitext(filename)[0]}.pdf")
                pythoncom.CoInitialize()  
                convert(file_path, pdf_path)

                return render_template('result.html', metadata=metadata, pdf_filename=os.path.basename(pdf_path))
            except Exception as e:
                return f"Error during conversion: {str(e)}", 500
        return "Invalid file format! Only .docx files are allowed.", 400

    return render_template('index.html')


@app.route('/download/<filename>')
def download_file(filename):
    file_path = os.path.join(PDF_FOLDER, filename)
    if os.path.exists(file_path):
        return send_file(file_path, as_attachment=True)
    return "File not found!", 404

@app.route('/convert', methods=['POST'])
def convert_to_pdf():
    if 'file' not in request.files:
        return "No file uploaded!", 400

    file = request.files['file']
    if file.filename == '':
        return "No selected file!", 400

    if file and file.filename.endswith('.docx'):
        try:
            filename = secure_filename(file.filename)
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(file_path)

            
            pdf_path = os.path.join(PDF_FOLDER, f"{os.path.splitext(filename)[0]}.pdf")
            pythoncom.CoInitialize()  
            convert(file_path, pdf_path)

            
            password = request.form.get("password")
            if password:
                protected_pdf_path = os.path.join(PDF_FOLDER, f"protected_{os.path.basename(pdf_path)}")
                add_password_to_pdf(pdf_path, protected_pdf_path, password)
                return send_file(protected_pdf_path, as_attachment=True)

            return send_file(pdf_path, as_attachment=True)
        except Exception as e:
            return f"Error during conversion: {str(e)}", 500

    return "Invalid file format! Only .docx files are allowed.", 400


if __name__ == "__main__":
    app.run(debug=True)
