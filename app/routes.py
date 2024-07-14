from flask import Blueprint, render_template, request, send_file
import docx
import io
import os
from docx.enum.text import WD_COLOR_INDEX
from docx.shared import Inches

main = Blueprint('main', __name__)

def replace_highlighted_words(doc, replacements, image_paths):
    for para in doc.paragraphs:
        for run in para.runs:
            if run.font.highlight_color == WD_COLOR_INDEX.YELLOW:
                for old_word, new_word in replacements.items():
                    if old_word in run.text:
                        run.text = run.text.replace(old_word, new_word)
                        run.font.highlight_color = None
                if "gbremonev" in run.text and image_paths.get("gbremonev"):
                    run.text = run.text.replace("gbremonev", "")
                    run.add_picture(image_paths["gbremonev"], width=Inches(2))
                    run.font.highlight_color = None
                if "gmbrsmart" in run.text and image_paths.get("gmbrsmart"):
                    run.text = run.text.replace("gmbrsmart", "")
                    run.add_picture(image_paths["gmbrsmart"], width=Inches(2))
                    run.font.highlight_color = None
    return doc

@main.route('/')
def home():
    return render_template('form.html')

@main.route('/replace', methods=['POST'])
def replace_word():
    replacements = {
        'JUDUL': request.form['judul'],
        'Alamat': request.form['alamat'],
        'Telpon': request.form['telpon'],
        'Laman': request.form['laman'],
        'Surel': request.form['surel'],
        'NAMSAT': request.form['namsat'],
        'satker': request.form['satker'],
        'triwulan': request.form['triwulan'],
        'tahun': request.form['tahun'],
        'output': request.form['output'],
        'wulantri': request.form['wulantri'],
        'hunta': request.form['hunta'],
        'putout': request.form['putout'],
        'hambatan1': request.form['hambatan1'],
        'hambatan2': request.form['hambatan2'],
        'rencana1': request.form['rencana1'],
        'rencana2': request.form['rencana2']
    }
    
    image_paths = {}
    if 'gambar_gbremonev' in request.files:
        image_file = request.files['gambar_gbremonev']
        image_path = os.path.join(r'C:\Users\user\Documents\PKL\KEMENKUMHAM\program website\app\static\images\uploads', image_file.filename)
        image_file.save(image_path)
        image_paths['gbremonev'] = image_path

    if 'gambar_gmbrsmart' in request.files:
        image_file = request.files['gambar_gmbrsmart']
        image_path = os.path.join(r'C:\Users\user\Documents\PKL\KEMENKUMHAM\program website\app\static\images\uploads', image_file.filename)
        image_file.save(image_path)
        image_paths['gmbrsmart'] = image_path
    
    doc_path = r'C:\Users\user\Documents\PKL\KEMENKUMHAM\program website\file download\laporanevaluasi.docx'
    if not os.path.exists(doc_path):
        return f"Dokumen tidak ditemukan di path yang diberikan: {doc_path}"
    
    doc = docx.Document(doc_path)
    
    doc = replace_highlighted_words(doc, replacements, image_paths)
    
    new_doc_path = r'C:\Users\user\Documents\PKL\KEMENKUMHAM\program website\file download\update_laporanevaluasi.docx'
    doc.save(new_doc_path)
    
    return send_file(new_doc_path, as_attachment=True)
