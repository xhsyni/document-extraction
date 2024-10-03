from flask import Flask, render_template
from email_flask import email_extractor_bp
from pdf_flask import pdf_extractor_bp

app = Flask(__name__)
extract_email = app.register_blueprint(email_extractor_bp)
extract_pdf = app.register_blueprint(pdf_extractor_bp)
@app.route('/')
def index():
    return render_template('pop_up.html')

@app.route('/email')
def email():
    extract_email
    
@app.route('/pdf')
def pdf():
    extract_pdf
    
if __name__ == '__main__':
    app.run(debug=True, port=5000)

