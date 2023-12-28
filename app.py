# app.py

from flask import Flask, render_template, request, send_file, flash
from pdf2docx import Converter
from PIL import Image, ImageDraw, ImageFont
import comtypes.client
from pdf2image import convert_from_path
from moviepy.editor import VideoFileClip
from waitress import serve
from PIL import ImageFilter
import gunicorn
from comtypes.client import CreateObject
import fitz  
import mimetypes
from PyPDF2 import PdfReader, PdfWriter
import pythoncom
from fpdf import FPDF
import os


app = Flask(__name__)
app.secret_key = 'H0139@ah'
app.config['UPLOAD_FOLDER'] = os.path.join(os.getcwd(), 'uploads')
app.config['RESIZED_FOLDER'] = os.path.join(os.getcwd(), 'resized')
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024
app.config['SESSION_COOKIE_SAMESITE'] = None

# Landing page
@app.route('/')
def index():
    return render_template('index.html')

# File conversion route
@app.route('/convert', methods=['POST'])
def convert_file():
    if request.method == 'POST':
        file = request.files['file']
        conversion_type = request.form['conversion_type']
        if file and conversion_type:
            filename = file.filename
            try:
                base_dir = os.path.dirname(os.path.abspath(__file__))
                input_path = os.path.join(base_dir, f"uploads/{filename}")

                if filename.lower().endswith('.docx') and conversion_type == 'docx_to_pdf':
                    # Convert DOCX to PDF
                    pdf_filename = filename.replace('.docx', '.pdf')
                    output_path = os.path.join(base_dir, f"converted/{pdf_filename}")

                    try:
                        file.save(input_path)

                        word = comtypes.client.CreateObject("Word.Application")
                        doc = word.Documents.Open(input_path)
                        doc.SaveAs(output_path, FileFormat=17)
                        doc.Close()
                        word.Quit()

                        # After a successful conversion, show success message
                        flash("File successfully converted and download completed!", 'success')
                        return send_file(output_path, as_attachment=True)
                    except FileNotFoundError:
                        return "File not found. Please check the uploaded file."
                
                elif filename.lower().endswith('.doc') and conversion_type == 'docx_to_pdf':
                    # Convert DOCX to PDF
                    pdf_filename = filename.replace('.doc', '.pdf')
                    output_path = os.path.join(base_dir, f"converted/{pdf_filename}")

                    try:
                        file.save(input_path)

                        word = comtypes.client.CreateObject("Word.Application")
                        doc = word.Documents.Open(input_path)
                        doc.SaveAs(output_path, FileFormat=17)
                        doc.Close()
                        word.Quit()

                        # After a successful conversion, show success message
                        flash("File successfully converted and download completed!", 'success')
                        return send_file(output_path, as_attachment=True)
                    except FileNotFoundError:
                        return "File not found. Please check the uploaded file."

                elif filename.lower().endswith('.jpg') and conversion_type == 'jpg_to_png':
                    # Convert JPG to PNG with transparent background
                    png_filename = filename.replace('.jpg', '.png')
                    output_path = os.path.join(base_dir, f"converted/{png_filename}")

                    file.save(input_path)

                    img = Image.open(input_path)
                    img = img.convert("RGBA")

                    # Set a threshold value for transparency (adjust as needed)
                    threshold = 200
                    img_data = list(img.getdata())

                    # Create a new image with transparent background
                    new_img_data = [
                        (r, g, b, a) if len((r, g, b, a)) == 4 and max((r, g, b)) > threshold else (r, g, b, 255)
                        for (r, g, b, a) in img_data
                    ]
                    img.putdata(new_img_data)

                    img.save(output_path, 'PNG')

                    # After a successful conversion, show success message
                    flash("File successfully converted and download completed!", 'success')
                    return send_file(output_path, as_attachment=True)
                
                elif filename.lower().endswith('.png') and conversion_type == 'png_to_jpg':
                    # Convert JPG to PNG with transparent background
                    png_filename = filename.replace('.png', '.jpg')
                    output_path = os.path.join(base_dir, f"converted/{png_filename}")

                    file.save(input_path)

                    img = Image.open(input_path)
                    img = img.convert("RGBA")

                    # Set a threshold value for transparency (adjust as needed)
                    threshold = 200
                    img_data = list(img.getdata())

                    # Create a new image with transparent background
                    new_img_data = [
                        (r, g, b, a) if len((r, g, b, a)) == 4 and max((r, g, b)) > threshold else (r, g, b, 255)
                        for (r, g, b, a) in img_data
                    ]
                    img.putdata(new_img_data)

                    img.save(output_path, 'JPG')

                    # After a successful conversion, show success message
                    flash("File successfully converted and download completed!", 'success')
                    return send_file(output_path, as_attachment=True)
                
                elif filename.lower().endswith('.jpeg') and conversion_type == 'jpeg_to_png':
                    # Convert JPG to PNG with transparent background
                    png_filename = filename.replace('.jpeg', '.png')
                    output_path = os.path.join(base_dir, f"converted/{png_filename}")

                    file.save(input_path)

                    img = Image.open(input_path)
                    img = img.convert("RGBA")

                    # Set a threshold value for transparency (adjust as needed)
                    threshold = 200
                    img_data = list(img.getdata())

                    # Create a new image with transparent background
                    new_img_data = [
                        (r, g, b, a) if len((r, g, b, a)) == 4 and max((r, g, b)) > threshold else (r, g, b, 255)
                        for (r, g, b, a) in img_data
                    ]
                    img.putdata(new_img_data)

                    img.save(output_path, 'JPEG')

                    # After a successful conversion, show success message
                    flash("File successfully converted and download completed!", 'success')
                    return send_file(output_path, as_attachment=True)

                elif filename.lower().endswith('.pdf') and conversion_type == 'pdf_to_ppt':
                    # Convert PDF to PPT
                    ppt_filename = filename.replace('.pdf', '.ppt')
                    output_path = os.path.join(base_dir, f"converted/{ppt_filename}")

                    file.save(input_path)

                    powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
                    presentation = powerpoint.Presentations.Add()
                    presentation.SaveAs(output_path, 24)
                    presentation.Close()
                    powerpoint.Quit()

                    # After a successful conversion, show success message
                    flash("File successfully converted and download completed!", 'success')
                    return send_file(output_path, as_attachment=True)

                elif filename.lower().endswith('.ppt') and conversion_type == 'ppt_to_pdf':
                    # Convert PPT to PDF
                    pdf_filename = filename.replace('.ppt', '.pdf')
                    output_path = os.path.join(base_dir, f"converted/{pdf_filename}")

                    file.save(input_path)

                    powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
                    presentation = powerpoint.Presentations.Open(input_path)
                    presentation.SaveAs(output_path, 32)
                    presentation.Close()
                    powerpoint.Quit()

                    # After a successful conversion, show success message
                    flash("File successfully converted and download completed!", 'success')
                    return send_file(output_path, as_attachment=True)
                
                elif filename.lower().endswith('.pptx') and conversion_type == 'ppt_to_pdf':
                    # Convert PPT to PDF
                    pdf_filename = filename.replace('.pptx', '.pdf')
                    output_path = os.path.join(base_dir, f"converted/{pdf_filename}")

                    file.save(input_path)

                    powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
                    presentation = powerpoint.Presentations.Open(input_path)
                    presentation.SaveAs(output_path, 32)
                    presentation.Close()
                    powerpoint.Quit()

                    # After a successful conversion, show success message
                    flash("File successfully converted and download completed!", 'success')
                    return send_file(output_path, as_attachment=True)

                elif filename.lower().endswith('.mp4') and conversion_type == 'video_to_audio':
                    # Convert video to audio using moviepy
                    audio_filename = filename.replace('.mp4', '.mp3')
                    output_path = os.path.join(base_dir, f"converted/{audio_filename}")

                    file.save(input_path)

                    video_clip = VideoFileClip(input_path)
                    audio_clip = video_clip.audio
                    audio_clip.write_audiofile(output_path)
                    audio_clip.close()

                    # After a successful conversion, show success message
                    flash("File successfully converted and download completed!", 'success')
                    return send_file(output_path, as_attachment=True)
                
                elif filename.lower().endswith('.m4v') and conversion_type == 'video_to_audio':
                    # Convert video to audio using moviepy
                    audio_filename = filename.replace('.m4v', '.mp3')
                    output_path = os.path.join(base_dir, f"converted/{audio_filename}")

                    file.save(input_path)

                    video_clip = VideoFileClip(input_path)
                    audio_clip = video_clip.audio
                    audio_clip.write_audiofile(output_path)
                    audio_clip.close()

                    # After a successful conversion, show success message
                    flash("File successfully converted and download completed!", 'success')
                    return send_file(output_path, as_attachment=True)
                
                elif filename.lower().endswith('.mov') and conversion_type == 'video_to_audio':
                    # Convert video to audio using moviepy
                    audio_filename = filename.replace('.mov', '.mp3')
                    output_path = os.path.join(base_dir, f"converted/{audio_filename}")

                    file.save(input_path)

                    video_clip = VideoFileClip(input_path)
                    audio_clip = video_clip.audio
                    audio_clip.write_audiofile(output_path)
                    audio_clip.close()

                    # After a successful conversion, show success message
                    flash("File successfully converted and download completed!", 'success')
                    return send_file(output_path, as_attachment=True)
                
                elif filename.lower().endswith('.avi') and conversion_type == 'video_to_audio':
                    # Convert video to audio using moviepy
                    audio_filename = filename.replace('.avi', '.mp3')
                    output_path = os.path.join(base_dir, f"converted/{audio_filename}")

                    file.save(input_path)

                    video_clip = VideoFileClip(input_path)
                    audio_clip = video_clip.audio
                    audio_clip.write_audiofile(output_path)
                    audio_clip.close()

                    # After a successful conversion, show success message
                    flash("File successfully converted and download completed!", 'success')
                    return send_file(output_path, as_attachment=True)

                elif filename.lower().endswith('.pdf') and conversion_type == 'pdf_to_docx':
                    # Convert DOCX to PDF
                    pdf_filename = filename.replace('.pdf', '.docx')
                    output_path = os.path.join(base_dir, f"converted/{pdf_filename}")

                    try:
                        file.save(input_path)

                        word = comtypes.client.CreateObject("Word.Application")
                        doc = word.Documents.Open(input_path)
                        doc.SaveAs(output_path, FileFormat=17)
                        doc.Close()
                        word.Quit()

                        # After a successful conversion, show success message
                        flash("File successfully converted and download completed!", 'success')
                        return send_file(output_path, as_attachment=True)
                    except FileNotFoundError:
                        return "File not found. Please check the uploaded file."
                
                elif filename.lower().endswith('.jpg') and conversion_type == 'jpg_to_pdf':
                    # Convert JPG to PDF
                    pdf_filename = filename.replace('.jpg', '.pdf')
                    output_path = os.path.join(base_dir, f"converted/{pdf_filename}")

                    try:
                        file.save(input_path)

                        # Check if the user specified a range (e.g., image1.jpg to image5.jpg)
                        start_index = request.form.get('start_index', 1)
                        end_index = request.form.get('end_index', 1)

                        # Convert the range of images to a single PDF
                        convert_images_to_pdf(input_path, output_path, int(start_index), int(end_index))

                        # After a successful conversion, show success message
                        flash("File successfully converted and download completed!", 'success')
                        return send_file(output_path, as_attachment=True)
                    except FileNotFoundError:
                        return "File not found. Please check the uploaded file."
                elif filename.lower().endswith('.pdf') and conversion_type == 'pdf_to_jpg':
                    # Convert PDF to JPG
                    jpg_filename = filename.replace('.pdf', '.jpg')
                    output_path = os.path.join(base_dir, f"converted/{jpg_filename}")

                    file.save(input_path)

                    images = convert_from_path(input_path, 500)  # Adjust the DPI (here, it's set to 500)
                    images[0].save(output_path, 'JPEG')

                    flash("File successfully converted and download completed!", 'success')
                    return send_file(output_path, as_attachment=True)
           
                else:
                    return "Selected conversion type is not supported."

            except Exception as e:
                flash(f"Error during conversion: {e}", 'error')
                return f"Error during conversion: {e}"
def convert_images_to_pdf(input_path, output_path, start_index, end_index):
    pdf = FPDF()
    for i in range(start_index, end_index + 1):
        img_path = input_path.replace('.jpg', f'{i}.jpg')  # Assuming sequential numbering
        
        # Open the image using Pillow and convert to RGB mode
        img = Image.open(img_path).convert('RGB')
        img.save(img_path.replace('.jpg', '.png'))  # Save the converted image as PNG
        
        pdf.add_page()
        pdf.image(f'"{img_path.replace(".jpg", ".png")}"', x=10, y=10, w=190)  # Use the PNG version

    pdf.output(output_path, "F")

# Landing Resize page 
@app.route('/resize_page')
def resize_page():
    return render_template('resize.html')

# Resize based on file size
@app.route('/resize', methods=['POST'])
def resize_file():
    if request.method == 'POST':
        file = request.files['file']
        
        if file:
            filename = file.filename
            try:
                base_dir = os.path.dirname(os.path.abspath(__file__))
                input_path = os.path.join(base_dir, f"uploads/{filename}")
                output_path = os.path.join(base_dir, f"resized/{filename}")

                # Resize based on file type
                if filename.lower().endswith('.jpg') or filename.lower().endswith('.png') or filename.lower().endswith('.jpeg'):
                    # Image file, use the existing image resizing function
                    resized_path = resize_based_on_file_size(file, input_path, output_path)
                elif filename.lower().endswith('.pdf'):
                    # PDF file, use the PDF resizing function
                    resized_path = reduce_pdf_size(file, input_path, output_path)
                elif filename.lower().endswith(('.doc', '.docx')):
                    # DOC or DOCX file, use the DOC resizing function
                    resized_path = reduce_doc_size(file, input_path, output_path)
                elif filename.lower().endswith(('.ppt', '.pptx')):
                    # PPT or PPTX file, use the PPT resizing function
                    resized_path = reduce_ppt_size(file, input_path, output_path)
                else:
                    # Unsupported file type
                    flash("Unsupported file type for resizing.", 'error')
                    return render_template('resize.html')

                return send_file(resized_path, as_attachment=True)
            except Exception as e:
                flash(f"Error during resizing: {e}", 'error')
            else:
                flash("File successfully resized and download completed!", 'success')

    return render_template('resize.html')

def resize_based_on_file_size(file, input_path, output_path):
    # Desired file size in bytes (adjust as needed)
    desired_file_size = 300 * 300  # 300 MB

    file.save(input_path)
    original_size = os.path.getsize(input_path)

    if original_size <= desired_file_size:
        # No need to resize, the file is already within the desired size
        return input_path

    # Calculate the resize factor
    resize_factor = (desired_file_size / original_size) ** 0.5

    # Open the image using PIL
    img = Image.open(input_path)

    # Calculate the new dimensions
    new_width = int(img.width * resize_factor)
    new_height = int(img.height * resize_factor)

    try:
        # Resize the image using ANTIALIAS
        img = img.resize((new_width, new_height), Image.ANTIALIAS)
    except AttributeError:
        # If 'ANTIALIAS' is not available, fall back to default resizing method
        img = img.resize((new_width, new_height))

    # Save the resized image
    img.save(output_path)

    return output_path

def reduce_pdf_size(file, input_path, output_path):
    # Desired file size in bytes (adjust as needed)
    desired_file_size = 3 * 1024 * 1024  # 10 MB

    file.save(input_path)
    original_size = os.path.getsize(input_path)

    if original_size <= desired_file_size:
        # No need to resize, the file is already within the desired size
        return send_file(input_path, as_attachment=True, download_name=file.filename)

    # Open the PDF file using PyMuPDF
    pdf_document = fitz.open(input_path)

    # Create a new PDF document for resized pages
    resized_document = fitz.open()

    # Iterate through each page and add a scaled version to the new document
    for page_number in range(pdf_document.page_count):
        page = pdf_document[page_number]

        # Calculate the new dimensions (adjust the scaling factor as needed)
        new_width = int(page.rect.width * 0.5)
        new_height = int(page.rect.height * 0.5)

        # Create a new page with the calculated dimensions
        new_page = resized_document.new_page(width=new_width, height=new_height)

        # Draw the content of the original page onto the new page
        pix = page.get_pixmap()
        new_page.insert_image((0, 0, new_width, new_height), pixmap=pix)

    # Save the resized PDF
    resized_document.save(output_path)
    return output_path

    # # Determine MIME type based on file extension
    # mimetype, _ = mimetypes.guess_type(output_path)

    # return send_file(output_path, as_attachment=True, download_name='resized_file.pdf', mimetype=mimetype)
def reduce_doc_size(file, input_path, output_path):
    # Desired file size in bytes (adjust as needed)
    desired_file_size = 2*1024 * 1024  # 2 MB

    file.save(input_path)
    original_size = os.path.getsize(input_path)

    if original_size <= desired_file_size:
        # No need to resize, the file is already within the desired size
        return input_path

    # Use comtypes.client to reduce DOC size
    pythoncom.CoInitialize()
    word = CreateObject("Word.Application")
    doc = word.Documents.Open(input_path)

    # Save the reduced DOC
    doc.SaveAs2(output_path)
    doc.Close()
    word.Quit()

    return output_path

def reduce_ppt_size(file, input_path, output_path):
    # Desired file size in bytes (adjust as needed)
    desired_file_size = 3 * 1024 * 1024  # 3 MB

    file.save(input_path)
    original_size = os.path.getsize(input_path)

    if original_size <= desired_file_size:
        # No need to resize, the file is already within the desired size
        return input_path

    # Use comtypes.client to reduce PPT size
    pythoncom.CoInitialize()
    powerpoint = CreateObject("PowerPoint.Application")
    presentation = powerpoint.Presentations.Open(input_path)

    # Save the reduced PPT
    presentation.SaveAs(output_path, 32)  # 32 corresponds to PDF format
    presentation.Close()
    powerpoint.Quit()

    return output_path

# Landing Resize page 
@app.route('/about_page')
def about_page():
    return render_template('about.html')

