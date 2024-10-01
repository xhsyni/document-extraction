import fitz
import re
import numpy as np
import pandas as pd
import comtypes.client
import os
from collections import Counter
from nltk.tokenize import word_tokenize
from nltk.corpus import stopwords
from nltk.stem import WordNetLemmatizer
import nltk
import time
import sys
import platform
nltk.download('punkt')
nltk.download('stopwords')
nltk.download('wordnet')


def convert_to_pdf(input_file, output_file):
    try:
        current_os = platform.system()

        if current_os == "Windows":
            import comtypes.client
            comtypes.CoInitialize()
            
            input_file = os.path.abspath(input_file)
            output_file = os.path.abspath(output_file)
            
            # Initialize PowerPoint application
            powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
            powerpoint.Visible = 1
            
            # Open the presentation
            presentation = powerpoint.Presentations.Open(input_file)
            time.sleep(2)
            
            # Save as PDF
            presentation.SaveAs(output_file, FileFormat=32)  # 32 is the constant for PDF format
            presentation.Close()
            
            # Quit PowerPoint application
            powerpoint.Quit()
            
            return output_file

        elif current_os in ["Linux", "Darwin"]:  # Darwin is macOS
            # Use LibreOffice or OpenOffice for Linux and macOS
            input_file = os.path.abspath(input_file)
            output_file = os.path.abspath(output_file)
            
            # Make sure LibreOffice is installed
            # You can use unoconv for this conversion process
            command = f"libreoffice --headless --convert-to pdf {input_file} --outdir {os.path.dirname(output_file)}"
            os.system(command)
            
            return output_file if os.path.exists(output_file) else None

        else:
            print(f"Unsupported operating system: {current_os}")
            return None

    except Exception as e:
        print(f"Error during PowerPoint conversion: {e}")
        return None

    finally:
        if current_os == "Windows":
            comtypes.CoUninitialize()

def fonts(doc):
    data = []

    for page_num in range(len(doc)):
        # Load the current page
        page = doc.load_page(page_num)
        # Extract text blocks from the page
        blocks = page.get_text("dict")["blocks"]

        # Iterate through each block of text on the page
        for b in blocks:
            # Check if the block contains text
            if b['type'] == 0:
                # Iterate through each line in the block
                for l in b["lines"]:
                    # Iterate through each text span in the line
                    for s in l["spans"]:
                        # Round size, block-left, and block-top values for consistency
                        rounded_size = round(s['size'], 1)
                        rounded_block_top = round(s['bbox'][1] / 100, 0)

                        # Append the extracted information to the data list
                        data.append({
                            "Page numbers": page_num + 1,    # Page number (1-based index)
                            "Content": s['text'],            # Text content of the span
                            "Font": s['font'],               # Font used in the text span
                            "Size": rounded_size,            # Font size (rounded)
                            "Blocks": rounded_block_top      # Vertical position of the block (rounded)
                        })

    fontdf = pd.DataFrame(data, columns=["Page numbers", "Blocks", "Content", "Font", "Size"])

    return fontdf

def groupFontSize(df):
    df = df.groupby(["Page numbers","Blocks","Size","Font"]).agg({
        "Content": lambda x: " ".join(x)
    }).reset_index()
    df = df.sort_values(by=['Page numbers', 'Size'], ascending=[True, False]).reset_index(drop=True)

    return df

def clean_text(text):
    # Convert text to lowercase and whitespaces
    text = re.sub(r'\s+', ' ', text.lower())

    # Replace newlines with spaces
    text = re.sub(r'\n', ' ', text)

    # Replace tabs with spaces
    text = re.sub(r'\t', ' ', text)

    # Remove control characters
    text = re.sub(r'[\x00-\x1F\x7F]', ' ', text)

    # Remove URLs
    text = re.sub(r'https?://\S+', '', text)
    text = re.sub(r'www\.\S+', '', text)

    # Remove non-alphabetic characters
    text = re.sub(r'[^a-zA-Z0-9\s]', " ", text)

    # Remove names of months from the text
    text = re.sub(r'\b(?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)\b', " ", text)

    # Remove common abbreviations and not applicable
    text = re.sub(r'\b(fy|na|n a|yr)\b', '', text)

    return text.strip()

def identify_header(df):
    header_combinations = []

    # Iterate over each page's data
    for page_num, page_data in df.groupby("Page numbers"):
        header_row = page_data.iloc[0]
        header_combinations.append((header_row[1], header_row[2], header_row[3])) # block, size, font

    # Find the most common header combination
    most_common_combination = max(Counter(header_combinations).items(), key=lambda x: x[1])[0]

    # Split the combination
    block,size,font = most_common_combination

    return block,size,font

def identify_footer(df):
    footer_combinations = []

    # Iterate over each page's data
    for page_num, page_data in df.groupby("Page numbers"):
        header_row = page_data.iloc[-1]
        footer_combinations.append((header_row[1], header_row[2], header_row[3])) # block, size, font

    # Find the most common header combination
    most_common_combination = max(Counter(footer_combinations).items(), key=lambda x: x[1])[0]

    # Split the combination
    block,size,font = most_common_combination

    return block,size,font

def setHeaderFooter(header,footer):
    header_footer = {}
    header_footer['Subject'] = {
        'Blocks': header[0],
        'Size': header[1],
        'Font': header[2]
    }
    header_footer['Footer'] = {
        'Blocks': footer[0],
        'Size': footer[1],
        'Font': footer[2]
    }

    return header_footer

def remove_header_footer(df, header_footer, footerexist):
    # Filter rows that match header characteristics
    header_rows = df[((df['Size'] == header_footer['Subject']['Size']) &
                      (df['Font'] == header_footer['Subject']['Font'])) |
                     (df['Blocks'] == header_footer['Subject']['Blocks'])]


    df = df.drop(header_rows.index)
    header_rows = header_rows.groupby('Page numbers').agg({'Content': ' '.join}).reset_index()
    header_rows["Subject"] = header_rows["Content"]
    header_rows = header_rows.drop(columns=["Content"])

    header_rows['Subject'] = header_rows['Subject'].apply(clean_text)

    if footerexist:
      # Filter rows that match footer characteristics
      footer_rows = df[(df['Blocks'] == header_footer['Footer']['Blocks']) &
                      (df['Size'] == header_footer['Footer']['Size']) &
                      (df['Font'] == header_footer['Footer']['Font'])]

      df = df.drop(footer_rows.index)
      footer_rows = footer_rows.groupby('Page numbers').agg({'Content': ''.join}).reset_index()
      footer_rows["Footer"] = footer_rows["Content"]
      footer_rows = footer_rows.drop(columns=["Content"])
    else:
      footer_rows = pd.DataFrame(columns=["Page numbers", "Footer"])

    footer_rows['Footer'] = footer_rows['Footer'].apply(clean_text)

    return df, header_rows, footer_rows

def normalize_text(text):
    # Tokenize the text into words
    tokens = word_tokenize(text)

    # Keep only alphabetic words, remove numbers, punctuation
    tokens = [word for word in tokens if word.isalpha()]

    # Remove common English stopwords
    tokens = [word for word in tokens if word not in stopwords.words('english')]

    # Initialize the lemmatizer
    lemmatizer = WordNetLemmatizer()

    # Lemmatize each word in the tokens list
    tokens = [lemmatizer.lemmatize(word) for word in tokens]

    return ' '.join(tokens)

def runAllProcess(doc,footer_exist):
    all = fonts(doc)
    df = groupFontSize(all)
    df["Content"] = df["Content"].apply(clean_text)
    df = df[df['Content'].str.strip() != '']
    df.reset_index(drop=True, inplace=True)
    header = identify_header(df)
    footer = identify_footer(df)
    header_footer = setHeaderFooter(header, footer)
    df, header_rows, footer_rows = remove_header_footer(df, header_footer, footer_exist)
    df.drop(columns=["Blocks", "Size", "Font"], inplace=True)
    df = df.groupby('Page numbers').agg({'Content': ' '.join}).reset_index()
    df = pd.merge(df, header_rows, on='Page numbers', how='left')
    df = pd.merge(df, footer_rows, on='Page numbers', how='left')
    df['Content'] = df['Content'].apply(normalize_text)

    return df

from flask import Blueprint, render_template, render_template_string, send_file,request
import threading
from werkzeug.utils import secure_filename
import time
import os
import fitz

pptx_extractor_bp = Blueprint('pptx_extractor_bp',__name__)

UPLOAD_FOLDER = 'uploads'
TEMP_FOLDER = 'temps'
CSV_FOLDER = 'csv_files'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(TEMP_FOLDER, exist_ok=True)
os.makedirs(CSV_FOLDER, exist_ok=True)

ALLOWED_EXTENSIONS = {'pptx','ppt'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit(".",1)[1].lower() in ALLOWED_EXTENSIONS

@pptx_extractor_bp.route('/ppt')
def hello():
    return render_template("ppt.html")

@pptx_extractor_bp.route('/extract-pptx',methods=['POST'])
def extract_ppt():
    if 'pptxFile' in request.files:
        file = request.files['pptxFile']
        if file.filename == '':
            return render_template_string("""
                <html>
                    <head>
                        <title>No file uploaded</title>
                        <link rel="stylesheet" type="text/css" href="/static/doc.css">
                    </head>
                    <body>
                        <h2>No files uploaded!</h2>
                        <form method="GET" action="/ppt">
                            <button type="submit">Back to extract</button>
                        </form>
                    </body>
                </html>
            """)
        
        if not allowed_file(file.filename):
                return render_template_string("""
                    <html>
                        <head>
                            <title>Invalid File Type</title>
                            <link rel="stylesheet" type="text/css" href="/static/doc.css">
                        </head>
                        <body>
                            <h2>Please select a valid file type (PPT).</h2>
                            <form method="GET" action="/ppt">
                                <button type="submit">Back to extract</button>
                            </form>
                        </body>
                    </html>
                """)
        
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            file_path = os.path.join(UPLOAD_FOLDER, filename)
            file.save(file_path)

            start_time = time.time()
            output_pdf = os.path.join(TEMP_FOLDER, filename.split(".")[0] + ".pdf")
            output = convert_to_pdf(file_path, output_pdf)
            
            if output is None:
                return "Error converting file to PDF", 500

            doc = fitz.open(output_pdf)
            final_df = runAllProcess(doc,request.form.get('footerexists-file'))
            end_time = time.time()
            execution_time = end_time-start_time
            doc.close()
            
            csv_filename = f"{filename.split('.')[0]}.csv"
            csv_path = os.path.join(CSV_FOLDER, csv_filename)
            final_df.to_csv(csv_path, index=False)

            os.remove(file_path)
            os.remove(output_pdf)   

            return render_template_string("""
                <html>
                    <head>
                        <title>Download CSV</title>
                        <link rel="stylesheet" type="text/css" href="/static/doc.css">
                    </head>
                    <body>
                        <div class="complete-container">
                            <h1>Processing Complete!!!</h1>
                            <p>Your PPTs have been processed. Click the buttons below to download the CSV files.</p>
                            <div class="download-card">
                                <p>The executed time for <strong>{{ csv_filename }}</strong> is <strong>{{ execution_time }}</strong> seconds!</p>
                                <form method="GET" action="/download_csv">
                                    <input type="hidden" name="filename" value="{{ csv_filename }}">
                                    <button type="submit">Download CSV</button>
                                </form>
                            </div>
                            <form method="GET" action="/ppt">
                                <button type="submit">Back to Upload</button>
                            </form>
                        </div>
                    </body>
                </html>
            """, csv_filename=csv_filename,execution_time=round(execution_time,2))
        
    elif 'pptxFolder' in request.files:
        files = request.files.getlist('pptxFolder')
        if not files or len(files) == 0:
            return "No selected files in folder",400
        
        file_path = []
        all_dfs = []
        maindf = pd.DataFrame()
        merged_df = ''

        for file in files:
            if file.filename == '':
                return render_template_string("""
                <html>
                    <head>
                        <title>No file uploaded</title>
                        <link rel="stylesheet" type="text/css" href="/static/doc.css">
                    </head>
                    <body>
                        <h2>No files uploaded!</h2>
                        <form method="GET" action="/ppt">
                            <button type="submit">Back to extract</button>
                        </form>
                    </body>
                </html>
            """)

            if not allowed_file(file.filename):
                return render_template_string("""
                    <html>
                        <head>
                            <title>Invalid File Type</title>
                            <link rel="stylesheet" type="text/css" href="/static/doc.css">
                        </head>
                        <body>
                            <h2>Please select a valid file type (PPT).</h2>
                            <form method="GET" action="/ppt">
                                <button type="submit">Back to extract</button>
                            </form>
                        </body>
                    </html>
                """)
            
            if file and allowed_file(file.filename):
                filename = secure_filename(file.filename)
                file_path = os.path.join(UPLOAD_FOLDER, filename)
                file.save(file_path)

                start_time = time.time()
                output_pdf = os.path.join(TEMP_FOLDER, filename.split(".")[0] + ".pdf")
                output = convert_to_pdf(file_path, output_pdf)
                
                if output is None:
                    return "Error converting file to PDF", 500

                doc = fitz.open(output_pdf)
                final_df = runAllProcess(doc,request.form.get('footerexists-folder'))
                end_time = time.time()
                execution_time = end_time-start_time
                doc.close()

                csv_filename = f"{filename.split('.')[0]}.csv"
                csv_path = os.path.join(CSV_FOLDER, csv_filename)
                final_df.to_csv(csv_path, index=False)

                all_dfs.append((csv_filename, round(execution_time,2)))

                os.remove(file_path)
                os.remove(output_pdf)

                merged_file = f"Merged CSV.csv"
                csv_path = os.path.join(CSV_FOLDER, merged_file)
                maindf = pd.concat([maindf,final_df])
                maindf.reset_index(inplace=True,drop=True)
                maindf.to_csv(csv_path, index=False)

                merged_df = merged_file

        return render_template_string("""
        <html>
            <head>
                <title>Download CSVs</title>
                <link rel="stylesheet" type="text/css" href="/static/doc.css">
            </head>
            <body>
                <div class="complete-container">
                    <h1>Processing Complete!!!</h1>
                    <p>Your PPTs have been processed. Click the buttons below to download the CSV files.</p>
                    {% for csv_filename, execution_time in all_dfs %}
                        <div class="download-card">                                
                            <p>The execution time for <strong>{{ csv_filename }}</strong> was <strong>{{ execution_time }}</strong> seconds.</p>
                            <form method="GET" action="/download_csv">
                                <input type="hidden" name="filename" value="{{ csv_filename }}">
                                <button type="submit">Download CSV</button>
                            </form>
                        </div>
                    {% endfor %}
                    <div class="merged-container">
                        <h3> This is the merged CSV file. </h3>
                        <form method="GET" action="/download_csv">
                            <input type="hidden" name="filename" value="{{ merged_df }}">
                            <button type="submit">Download Merged CSV</button>
                        </form>
                    </div>
                    <form method="GET" action="/ppt">
                        <button type="submit">Back to extract</button>
                    </form>
                </div>
            </body>
        </html>
    """, all_dfs=all_dfs, merged_df=merged_df)

@pptx_extractor_bp.route('/download_csv',methods=['GET'])
def download_csv():
    csv_filename = request.args.get('filename')
    csv_path = os.path.join(CSV_FOLDER, csv_filename)

    if os.path.exists(csv_path):
        return send_file(csv_path, as_attachment=True)
    else:
        return "File not found",400
