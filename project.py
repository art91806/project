from flask import Flask, render_template, request, send_file, redirect
from docx2pdf import convert
import os

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

@app.route('/', methods=['GET', 'POST'])
def home():
    if request.method == 'POST':
        if 'file' not in request.files:
            return redirect(request.url)
        
        file = request.files['file']
        
        if file.filename == '':
            return redirect(request.url)
        
        if file and file.filename.endswith('.docx'):
            word_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
            file.save(word_path)
            
            pdf_filename = file.filename.replace('.docx', '.pdf')
            pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], pdf_filename)
            
            try:
                convert(word_path, pdf_path)
                os.remove(word_path)
                return send_file(pdf_path, as_attachment=True)
            
            except Exception as e:
                return f"Ошибка при конвертации: {e}"
    
    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)