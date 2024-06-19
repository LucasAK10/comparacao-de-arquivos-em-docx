import os
import time
from flask import Flask, request, render_template, send_from_directory, flash, redirect, url_for
from werkzeug.utils import secure_filename
from apscheduler.schedulers.background import BackgroundScheduler
from docx import Document
from docx.shared import RGBColor, Pt
from docx.enum.section import WD_SECTION, WD_ORIENT


UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'docx'}
MAX_CONTENT_LENGTH = 16 * 1024 * 1024  # 16 MB
FILE_LIFETIME = 60 * 60 * 24  # 24 horas

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = MAX_CONTENT_LENGTH
app.secret_key = 'supersecretkey'  # Alterar para uma chave secreta adequada

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def read_docx(file_path):
    doc = Document(file_path)
    full_text = []
    for para in doc.paragraphs:
        full_text.append(para.text)
    return '\n'.join(full_text)

def add_paragraph(doc, text, bold=False, italic=False, font_size=12, alignment=WD_ORIENT.PORTRAIT.JUSTIFY):
    paragraph = doc.add_paragraph()
    run = paragraph.add_run(text)
    run.bold = bold
    run.italic = italic
    run.font.name = 'Calibri'
    run.font.size = Pt(font_size)
    paragraph.alignment = alignment

    # Define parágrafo com espaçamento e recuos especificados
    paragraph_format = paragraph.paragraph_format
    paragraph_format.left_indent = Pt(0)
    paragraph_format.right_indent = Pt(0)
    paragraph_format.first_line_indent = Pt(28.3)  # 2 cm
    paragraph_format.space_before = Pt(0)
    paragraph_format.space_after = Pt(6)
    paragraph_format.line_spacing = 1

    return paragraph

def compare_documents(file1_path, file2_path, output_file):
    doc1 = Document(file1_path)
    doc2 = Document(file2_path)
    output_doc = Document()

    # Configurações de página
    sections = output_doc.sections
    for section in sections:
        section.page_height = Pt(29.7 * 28.35)
        section.page_width = Pt(21 * 28.35)
        section.left_margin = Pt(2.5 * 28.35)
        section.right_margin = Pt(2 * 28.35)
        section.top_margin = Pt(2 * 28.35)
        section.bottom_margin = Pt(2 * 28.35)
        section.gutter = Pt(0)
        section.orientation = WD_ORIENT.PORTRAIT

    def mark_difference(doc, text, bold=False, italic=False, color=RGBColor(0, 0, 0)):
        p = add_paragraph(doc, text, bold=bold, italic=italic)
        run = p.runs[0]
        run.font.color.rgb = color

    def compare_paragraphs(paras1, paras2):
        for p1, p2 in zip(paras1, paras2):
            if p1.text != p2.text:
                mark_difference(output_doc, f'Documento 1: {p1.text}', bold=True, color=RGBColor(255, 0, 0))
                mark_difference(output_doc, f'Documento 2: {p2.text}', italic=True, color=RGBColor(0, 255, 0))
            else:
                add_paragraph(output_doc, p1.text)

    compare_paragraphs(doc1.paragraphs, doc2.paragraphs)
    output_doc.save(output_file)

def clean_upload_folder():
    now = time.time()
    for filename in os.listdir(UPLOAD_FOLDER):
        file_path = os.path.join(UPLOAD_FOLDER, filename)
        if os.path.isfile(file_path):
            file_age = now - os.path.getmtime(file_path)
            if file_age > FILE_LIFETIME:
                os.remove(file_path)
                print(f'Removido: {file_path}')

@app.route('/', methods=['GET', 'POST'])
def upload_files():
    if request.method == 'POST':
        if 'file1' not in request.files or 'file2' not in request.files:
            flash('Nenhum arquivo foi enviado.')
            return render_template('upload.html')
        file1 = request.files['file1']
        file2 = request.files['file2']

        if file1.filename == '' or file2.filename == '':
            flash('Nenhum arquivo foi selecionado.')
            return render_template('upload.html')

        if file1 and allowed_file(file1.filename) and file2 and allowed_file(file2.filename):
            filename1 = secure_filename(file1.filename)
            filename2 = secure_filename(file2.filename)
            file1_path = os.path.join(app.config['UPLOAD_FOLDER'], filename1)
            file2_path = os.path.join(app.config['UPLOAD_FOLDER'], filename2)
            output_file = os.path.join(app.config['UPLOAD_FOLDER'], 'relatorio_diferencas.docx')

            file1.save(file1_path)
            file2.save(file2_path)
            
            compare_documents(file1_path, file2_path, output_file)
            
            return redirect(url_for('result', filename='relatorio_diferencas.docx'))
        else:
            flash('Apenas arquivos DOCX são permitidos.')
            return render_template('upload.html')

    return render_template('upload.html')

@app.route('/result', methods=['GET', 'POST'])
def result():
    filename = request.args.get('filename')
    if request.method == 'POST':
        if 'new_upload' in request.form:
            return redirect(url_for('upload_files'))
        elif 'shutdown' in request.form:
            shutdown_server()
            return 'Servidor encerrado.'
    return render_template('result.html', filename=filename)

@app.route('/uploads/<filename>')
def uploaded_file(filename):
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename)

def shutdown_server():
    func = request.environ.get('werkzeug.server.shutdown')
    if func is None:
        raise RuntimeError('Não é possível encerrar o servidor.')
    func()

if __name__ == '__main__':
    if not os.path.exists(UPLOAD_FOLDER):
        os.makedirs(UPLOAD_FOLDER)

    scheduler = BackgroundScheduler()
    scheduler.add_job(clean_upload_folder, 'interval', hours=24)
    scheduler.start()

    try:
        app.run(debug=True)
    except (KeyboardInterrupt, SystemExit):
        scheduler.shutdown()
