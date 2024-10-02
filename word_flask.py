import fitz
import re
import numpy as np
import pandas as pd
import os
from collections import Counter
import comtypes.client
import time
from nltk.tokenize import word_tokenize
from nltk.corpus import stopwords
from nltk.stem import WordNetLemmatizer
import nltk
nltk.data.path.append('./nltk_data')

def convert_to_pdf(input_file, output_file):
    try:
        comtypes.CoInitialize()
        input_file = os.path.abspath(input_file)
        output_file = os.path.abspath(output_file)

        # Initialize Word application
        word = comtypes.client.CreateObject("Word.Application")
        word.Visible = False
        
        # Open the document
        doc = word.Documents.Open(input_file)
        time.sleep(2)

        # Save as PDF
        doc.SaveAs(output_file, FileFormat=17)  # 17 is the constant for PDF format
        doc.Close()

        # Quit Word application
        word.Quit()

        return output_file
    except Exception as e:
        print(f"Error during Word conversion: {e}")
        return None
    finally:
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
                            "Block-Top": rounded_block_top   # Vertical position of the block (rounded)
                        })

    fontdf = pd.DataFrame(data, columns=["Page numbers", "Block-Top", "Content", "Font", "Size"])

    return fontdf

def group(test):
    # Group the DataFrame by 'Page numbers', 'Block-Top', 'Font', and 'Size'
    test = test.groupby(['Page numbers', 'Block-Top', 'Font',"Size"]).agg({
        'Content': lambda x: ' '.join(x)
    }).reset_index()

    return test


def getfileyear(test, pdf_name):
    # Get the File Year
    month_year = []
    month_pattern = r"\b(?:January|February|March|April|May|June|July|August|September|October|November|December) \d{4}\b"
    year_pattern = r"\b\d{4}\b"

    # Applying the pattern to find month and year in the 'Content' column
    month_list = test['Content'].apply(lambda x: re.findall(month_pattern, x)).sum()
    year_list = test['Content'].apply(lambda x: re.findall(year_pattern, x)).sum()

    month_count = Counter(month_list)
    year_count = Counter(year_list)

    unique_month_list = list(month_count.items())
    unique_year_list = list(year_count.items())

    year_key, year_value = max(unique_year_list, key=lambda item: item[1])

    if unique_month_list:
        for month, count in unique_month_list:
            if year_key in month:
                month_year.append(month)
                break
    else:
        month_year.append(year_key)

    if any(char.isdigit() for char in pdf_name):
        fileName = pdf_name
    else:
        fileName = pdf_name+"_"+year_key

    final_df = pd.DataFrame({'Date': month_year, 'File Name': fileName})

    return final_df,fileName

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
    text = re.sub(r'[^a-zA-Z\s]', " ", text)

    # Remove names of months from the text
    text = re.sub(r'\b(?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)\b', " ", text)

    # Remove common abbreviations and not applicable
    text = re.sub(r'\b(fy|na|n a|yr)\b', '', text)

    return text.strip()

def isBold(font):
    # Check if 'bold' or 'black' is present in the font name
    if 'bold' in font.lower() or 'black' in font.lower():
        return 'True'
    else:
        return 'False'

def cleanText(text):
    text = re.sub(r'\s+', ' ', text.lower())
    text = re.sub(r'\n', ' ', text)
    text = re.sub(r'\t', ' ', text)

    return text.strip()

def identify_header(df):
    header_combinations = []

    # Iterate over each page's data
    for page_num, page_data in df.groupby("Page numbers"):
        if len(page_data) > 2:
            header_row = page_data.loc[page_data['Block-Top'].idxmin()]
            header_combinations.append((header_row[2], header_row[3], header_row[1]))  # (font, size, block)

    # Find the most common combination of font, size, and block
    most_common_combination = max(Counter(header_combinations).items(), key=lambda x: x[1])[0]

    # Split the combination
    font_header, size_header, block_header = most_common_combination

    return font_header, size_header, block_header

def identify_footer(df):
    footer_combinations = []

    # Iterate over each page's data
    for page_num, page_data in df.groupby("Page numbers"):
        footer_row = page_data.loc[page_data['Block-Top'].idxmax()]
        footer_combinations.append((footer_row[2], footer_row[3], footer_row[1]))  # (font, size, block)

    # Find the most common combination of font, size, and block
    most_common_combination = max(Counter(footer_combinations).items(), key=lambda x: x[1])[0]

    # Split the combination
    font_footer, size_footer, block_footer = most_common_combination

    return font_footer, size_footer, block_footer

def setHeaderFooter(test1,header_exists,footer_exists):
    # Store header and footer's styles into a dictionary
    header_footer = {}
    if header_exists:
        header = identify_header(test1)
        header_footer['Header'] = {
            'Font': header[0],
            'Size': header[1],
            'Blocks': header[2]
        }
    if footer_exists:
        footer = identify_footer(test1)
        header_footer['Footer'] = {
            'Font': footer[0],
            'Size': footer[1],
            'Blocks': footer[2]
        }

    return header_footer

def getHeaderFooter(test1,header_footer):
    if 'Header' in header_footer:
        # Get the text from the dataframe which have the same styles with header
        headerdf = test1[(test1['Block-Top'] == header_footer['Header']['Blocks']) &
                        (test1['Size'] == header_footer['Header']['Size']) &
                        (test1['Font'] == header_footer['Header']['Font'])]

        # Remove header
        test1 = test1.drop(headerdf.index)
        headerdf = headerdf.groupby('Page numbers').agg({'Content': ' '.join}).reset_index()
        headerdf['Header'] = headerdf['Content'].apply(cleanText)
        headerdf.drop(columns=['Content'], inplace=True)
        headerdf = headerdf[headerdf['Header'].str.strip() != ""]
        headerdf.reset_index(drop=True, inplace=True)
    else:
        headerdf = None

    if 'Footer' in header_footer:
        # Get the text from the dataframe which have the same styles with footer
        footerdf= test1[(test1['Block-Top'] == header_footer['Footer']['Blocks']) &
                        (test1['Size'] == header_footer['Footer']['Size']) &
                        (test1['Font'] == header_footer['Footer']['Font'])]

        # Remove footer
        test1 = test1.drop(footerdf.index)
        footerdf = footerdf.groupby('Page numbers').agg({'Content': ' '.join}).reset_index()
        footerdf['Footer'] = footerdf['Content'].apply(cleanText)
        footerdf.drop(columns=['Content'], inplace=True)
        footerdf = footerdf[footerdf['Footer'].str.strip() != ""]
        footerdf.reset_index(drop=True, inplace=True)
    else:
        footerdf = None

    return test1, headerdf, footerdf

def groupFontSize(test):
    # Group by Page numbers, Font, Size, and isBold
    test = test.groupby(['Page numbers',"Font","Size","isBold"]).agg({
        'Block-Top': lambda x: ' '.join(map(str,x)),
        "Content": lambda x: ' '.join(x)
    }).reset_index()

    return test

def mean_block(block):
    # Check if the input 'block' is already a float
    if isinstance(block, float):
        return block
    else:
        # Split the string 'block'
        block = block.split(" ")

        # Convert each substring in the list to a float
        for i in range(len(block)):
            block[i] = float(block[i])

        return round(np.mean(block), 0)

def mean_size(sizes):
    # Check if the input 'sizes' is already a float
    if isinstance(sizes, float):
        return sizes

    else:
        # Split the string 'sizes'
        size = sizes.split(" ")
        for i in range(len(size)):
            size[i] = float(size[i])

        return round(np.mean(size), 1)

def getHeading(test1):
    heading_list = []

    for page_num, page_data in test1.groupby('Page numbers'):
        # Check if the heading Content length is not more than 25 words
        if (page_data.iloc[0]['isBold'] and
            page_data.iloc[0]['Size'] >= np.mean(page_data.iloc[0]['Size']) and
            len(page_data.iloc[0]['Content'].split()) <= 25) and len(page_data)>1:

            heading_list.append(page_data.iloc[0])
        else:
            heading_list.append(None)

    # Filter out None entries
    heading_list = [row for row in heading_list if row is not None]

    # Convert the list of heading rows to a DataFrame if any valid headings were found
    if heading_list:
        headingsdf = pd.DataFrame(heading_list).reset_index(drop=True)
        headingsdf['Heading'] = headingsdf['Content']
        headingsdf.drop(columns=['Block-Top', 'Size', 'Content', "isBold","Font"], inplace=True)
    else:
        headingsdf = None

    headingsdf['Heading'] = headingsdf['Heading'].apply(cleanText)

    return headingsdf

def removeHeading(test1,headingsdf):

    heading_texts = set(headingsdf['Heading'].tolist()) if headingsdf is not None else set()

    test1['Content'] = test1['Content'].str.lower().str.strip()

    # Filter test1 to remove rows where the text is in the heading
    test1 = test1[~test1['Content'].isin(heading_texts)]
    test1 = test1.drop(columns=["Block-Top","Size","isBold","Font"])
    test1.reset_index(drop=True, inplace=True)

    return test1

def mergePage(test1):
    # Merge the text without the headings, headers, footer by the page numbers
    test1 = test1.groupby(['Page numbers']).agg({
        'Content': lambda x: ' '.join(x)
    }).reset_index()

    return test1

def mergeAll(test1,footerdf,headerdf,headingsdf):
    # Merging footerdf
    if footerdf is not None:
        test1 = test1.merge(footerdf, on='Page numbers', how='left')
    else:
        test1['Footer'] = pd.NA

    # Merging headerdf
    if headerdf is not None:
        test1 = test1.merge(headerdf, on='Page numbers', how='left')
    else:
        test1['Header'] = pd.NA

    # Merging headingsdf
    if headingsdf is not None:
        test1 = test1.merge(headingsdf, on='Page numbers', how='left')
    else:
        test1['Heading'] = pd.NA

    # Drop the row that content, footer, header, and heading is NA
    df = test1.dropna(subset=['Content','Footer', 'Header', 'Heading'],how='all')
    df.reset_index(drop=True, inplace=True)

    return df

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

def concatAll(final_df,df):
    # Merge the DataFrames on the 'Date' column
    concatenated_df = pd.concat([final_df, df], axis=1)

    # Fill up the null value
    concatenated_df['Date'].fillna(method='ffill', inplace=True)
    concatenated_df['File Name'].fillna(method='ffill', inplace=True)

    return concatenated_df

# call function
def runAllProcess(doc,file_name,headerexists,footerexists):
    # extract all the text from pdf
    allfontdf = fonts(doc)
    test = group(allfontdf)

    # get file year and file name
    finaldf,storeName = getfileyear(test,file_name)

    # Clean Content
    test['Content'] = test['Content'].apply(clean_text)

    # Remove rows where the 'Content' column is empty or contains only whitespace
    test = test[test['Content'].str.strip() != '']

    # Check if it is Bold
    test['isBold'] = test['Font'].apply(isBold)
    test.reset_index(drop=True, inplace=True)

    # Sort page numbers ascending and size false
    test['Content'] = test['Content'].apply(cleanText)
    test = test.sort_values(by=['Page numbers', 'Size','Block-Top'], ascending=[True, False,True]).reset_index(drop=True)

    # Identify the style of header and footer
    header_footer = setHeaderFooter(test,headerexists,footerexists)

    # Get the header and footer from the dataframe
    test,headerdf,footerdf = getHeaderFooter(test,header_footer)

    test = groupFontSize(test)
    test['Block-Top'] = test['Block-Top'].apply(mean_block)
    test['Size'] = test['Size'].apply(mean_size)
    test = test.sort_values(by=['Page numbers', 'Size','Block-Top'], ascending=[True, False,True]).reset_index(drop=True)

    # Get and remove the heading from the dataframe
    headingsdf = getHeading(test)
    test = removeHeading(test, headingsdf)

    # Group the page number and merge the content with header, footer, and headings
    test = mergePage(test)
    df = mergeAll(test,footerdf,headerdf,headingsdf)

    # Normalize it
    df['Content'] = df['Content'].apply(normalize_text)
    final_df = concatAll(finaldf,df)

    return final_df,storeName

from flask import Blueprint, request, send_file, render_template_string, render_template
import threading
from werkzeug.utils import secure_filename
import time
import os
import fitz

word_extractor_bp = Blueprint('word_extractor_bp',__name__)

ALLOWED_EXTENSIONS = {'docx','doct','docm','doc','dot','xml'}


UPLOAD_FOLDER = '/tmp/uploads'
TEMP_FOLDER = '/tmp/temps'
CSV_FOLDER = '/tmp/csv_files'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(TEMP_FOLDER, exist_ok=True)
os.makedirs(CSV_FOLDER, exist_ok=True)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@word_extractor_bp.route('/word')
def hello():
    return render_template('word.html')

@word_extractor_bp.route('/extract-word',methods=['POST'])
def extract_doc():
    if 'wordFile' in request.files:
        file = request.files['wordFile']

        if file.filename == '':
            return render_template_string("""
                <html>
                    <head>
                        <title>No file uploaded</title>
                        <link rel="stylesheet" type="text/css" href="/static/doc.css">
                    </head>
                    <body>
                        <h2>No files uploaded!</h2>
                        <form method="GET" action="/word">
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
                        <h2>Please select a valid file type (Word).</h2>
                        <form method="GET" action="/word">
                            <button type="submit">Back to extract</button>
                        </form>
                    </body>
                </html>
            """)

        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            file_path = os.path.join(UPLOAD_FOLDER, filename)
            file.save(file_path)

            header_exists = request.form.get('headerexists-file')
            footer_exists = request.form.get('footerexists-file')

            start_time = time.time()
            
            output_pdf = os.path.join(TEMP_FOLDER, filename.split(".")[0] + ".pdf")
            output = convert_to_pdf(file_path, output_pdf)
            
            if output is None:
                return "Error converting file to PDF", 500

            doc = fitz.open(output_pdf)
            final_df,store_name = runAllProcess(doc,filename.split(".")[0],header_exists,footer_exists)
            end_time = time.time()
            execution_time = end_time-start_time
            doc.close()

            csv_filename = f"{store_name}.csv"
            csv_path = os.path.join(CSV_FOLDER, csv_filename)
            final_df.to_csv(csv_path, index=False)

            os.remove(file_path)
            os.remove(output)

            return render_template_string("""
                <html>
                    <head>
                        <title>Download CSV</title>
                        <link rel="stylesheet" type="text/css" href="/static/doc.css">
                    </head>
                    <body>
                        <div class="complete-container">
                            <h1>Processing Complete!!!</h1>
                            <p>Your Docs have been processed. Click the buttons below to download the CSV files.</p>
                            <div class="download-card">
                                <p>The executed time for <strong>{{ csv_filename }}</strong> is <strong>{{ execution_time }}</strong> seconds!</p>
                                <form method="GET" action="/download_csv">
                                    <input type="hidden" name="filename" value="{{ csv_filename }}">
                                    <button type="submit">Download CSV</button>
                                </form>
                            </div>
                            <form method="GET" action="/word">
                                <button type="submit">Back to Upload</button>
                            </form>
                        </div>
                    </body>
                </html>
            """, csv_filename=csv_filename,execution_time=round(execution_time,2))
        
    elif 'wordFolder' in request.files:
        files = request.files.getlist('wordFolder')
        if not files or len(files) == 0:
            return render_template_string("""
                <html>
                    <head>
                        <title>No file uploaded</title>
                        <link rel="stylesheet" type="text/css" href="/static/doc.css">
                    </head>
                    <body>
                        <h2>No files uploaded!</h2>
                        <form method="GET" action="/word">
                            <button type="submit">Back to extract</button>
                        </form>
                    </body>
                </html>
            """)

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
                        <form method="GET" action="/word">
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
                            <h2>Please select a valid file type (Word).</h2>
                            <form method="GET" action="/word">
                                <button type="submit">Back to extract</button>
                            </form>
                        </body>
                    </html>
                """)

            if file and allowed_file(file.filename):
                filename = secure_filename(file.filename)
                file_path = os.path.join(UPLOAD_FOLDER, filename)
                file.save(file_path)

                header_exists = request.form.get('headerexists-folder')
                footer_exists = request.form.get('footerexists-folder')

                start_time = time.time()
                output_pdf = os.path.join(TEMP_FOLDER, filename.split(".")[0] + ".pdf")
                output = convert_to_pdf(file_path, output_pdf)
                
                if output is None:
                    return "Error converting file to PDF", 500

                doc = fitz.open(output_pdf)
                final_df,store_name = runAllProcess(doc,filename.split(".")[0],header_exists,footer_exists)
                end_time = time.time()
                execution_time = end_time-start_time
                doc.close()

                csv_filename = f"{store_name}.csv"
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
                
            else:
                return "Invalid file type",400 

        return render_template_string("""
            <html>
                <head>
                    <title>Download CSV</title>
                    <link rel="stylesheet" type="text/css" href="/static/doc.css">
                </head>
                <body>
                    <div class="complete-container">
                        <h1>Processing Complete!!!</h1>
                        <p>Your Docs have been processed. Click the buttons below to download the CSV files.</p>
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
                        <form method="GET" action="/word">
                            <button type="submit">Back to Upload</button>
                        </form>
                    </div>
                </body>
            </html>
        """,all_dfs=all_dfs, merged_df=merged_df)

@word_extractor_bp.route('/download_csv', methods=['GET'])
def download_csv():
    csv_filename = request.args.get('filename')
    csv_path = os.path.join(CSV_FOLDER, csv_filename)

    if os.path.exists(csv_path):
        return send_file(csv_path, as_attachment=True)
    else:
        return "File not found", 404
