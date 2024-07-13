from flask import Blueprint, render_template, request, send_file
import docx
import io
import os
from docx.enum.text import WD_COLOR_INDEX

main = Blueprint('main', __name__)

def replace_highlighted_words(doc, replacements):
    for para in doc.paragraphs:
        for run in para.runs:
            if run.font.highlight_color == WD_COLOR_INDEX.YELLOW:
                for old_word, new_word in replacements.items():
                    if old_word in run.text:
                        run.text = run.text.replace(old_word, new_word)
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
    
    doc_path = r'C:\Users\user\Documents\PKL\KEMENKUMHAM\program website\file download\laporanevaluasi.docx'
    if not os.path.exists(doc_path):
        return f"Dokumen tidak ditemukan di path yang diberikan: {doc_path}"
    
    doc = docx.Document(doc_path)
    
    doc = replace_highlighted_words(doc, replacements)
    
    new_doc_path = r'C:\Users\user\Documents\PKL\KEMENKUMHAM\program website\file download\update_laporanevaluasi.docx'
    doc.save(new_doc_path)
    
    return send_file(new_doc_path, as_attachment=True)
