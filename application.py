from flask import Flask, render_template
from email_flask import email_extractor_bp
from pdf_flask import pdf_extractor_bp
from pptx_flask import pptx_extractor_bp
from word_flask import word_extractor_bp

app = Flask(__name__)
extract_email = app.register_blueprint(email_extractor_bp)
extract_pdf = app.register_blueprint(pdf_extractor_bp)
extract_pptx = app.register_blueprint(pptx_extractor_bp)
extract_word = app.register_blueprint(word_extractor_bp)
@app.route('/')
def index():
    return render_template('pop_up.html')

@app.route('/email')
def email():
    extract_email
    
@app.route('/pdf')
def pdf():
    extract_pdf

@app.route('/ppt')
def ppt():
    extract_pptx

@app.route('/word')
def word():
    extract_word
    
if __name__ == '__main__':
    app.run(debug=True, port=5000)

