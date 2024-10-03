import imaplib
import email as em
from email.utils import parsedate_to_datetime
from email.header import decode_header
import pandas as pd
import re
from bs4 import BeautifulSoup
from nltk.tokenize import word_tokenize
from nltk.corpus import stopwords
from nltk.stem import WordNetLemmatizer
import nltk
import os
nltk.data.path.append('./nltk_data')
nltk.download('punkt')
nltk.download('stopwords')
nltk.download('wordnet')

def html_to_text(html_content):
    soup = BeautifulSoup(html_content, 'html.parser')
    return soup.get_text()

def decode_email_content(bodyContent):
    try:
        return bodyContent.decode('utf-8')
    except UnicodeDecodeError:
        try:
            return bodyContent.decode('latin-1')
        except UnicodeDecodeError:
            return bodyContent.decode('iso-8859-1')

def decode_header_value(header_value):
    decoded_parts = decode_header(header_value)
    decoded_str = ""
    for part, encoding in decoded_parts:
        if encoding is not None:
            decoded_str += part.decode(encoding)
        else:
            decoded_str += part if isinstance(part, str) else part.decode('utf-8')
    return decoded_str

def validate_passkey(user, passkey, domain):
    try:
        if domain == "gmail":
            mail = imaplib.IMAP4_SSL('imap.gmail.com')
            mail.login(user, passkey)
            return True

        elif domain == "outlook":
            mail = imaplib.IMAP4_SSL('outlook.office365.com')
            mail.login(user, passkey)
            return True

    except imaplib.IMAP4.error:
        return False

    return False

def extract_email_content(user,passkey,key,value,domain,selection):
    if domain == "gmail":
        mail = imaplib.IMAP4_SSL('imap.gmail.com')
        mail.login(user, passkey)

        if selection == "inbox":
            mail.select('inbox')
        elif selection == "sent":
            mail.select('"[Gmail]/Sent Mail"')

    elif domain == "outlook":
        mail = imaplib.IMAP4_SSL('outlook.office365.com')
        mail.login(user, passkey)

        if selection == "inbox":
            mail.select('inbox')
        elif selection == "sent":
            mail.select('"Sent Items"')

    if key == "all":
        search_criteria = 'ALL'
    else:
        search_criteria = f'({key} "{value}")'

    status, response = mail.search(None, search_criteria)

    if status != "OK":
        raise Exception("Error searching emails")

    email_ids = response[0].split()
    if not email_ids:
        return pd.DataFrame()  # Return empty DataFrame if no emails found

    # Fetch emails
    msgs = []
    for email_id in email_ids:
        status, data = mail.fetch(email_id, '(RFC822)')
        msgs.append(data)

    # Dictionary to store extracted information
    all_content = {"Datetime": [], "Sender": [], "Recipient": [], "Subject": [], "Content": [], "CC": [], "BCC": [], "Message-ID": [], "In-Reply-To": [], "References": []}

    for msg in msgs:
        for parts in msg:
            if isinstance(parts, tuple):
                myMsg = em.message_from_bytes(parts[1])
                sender = myMsg.get('From', "")
                recipient = myMsg.get('To', "")
                cc = myMsg.get('CC', "")
                bcc = myMsg.get('BCC', "")
                message_id = myMsg.get('Message-ID', "")
                in_reply_to = myMsg.get('In-Reply-To', "")
                references = myMsg.get('References', "")


                # Decode and clean up the subject line
                subject = myMsg.get('Subject', "")
                subject_cleaned = decode_header_value(subject)

                date_str = myMsg.get('Date', "")
                datetime_obj = parsedate_to_datetime(date_str) if date_str else None

                email_contents = []
                content_found = False

                for part in myMsg.walk():
                    content_type = part.get_content_type()

                    # If content is already processed, skip further parts
                    if content_found:
                        continue

                    # Handle text/html content
                    if content_type == "text/html":
                        bodyContent = part.get_payload(decode=True)
                        decoded_content = decode_email_content(bodyContent)
                        email_contents.append(html_to_text(decoded_content))  # Convert HTML to text
                        content_found = True
                        break

                    # Handle text/plain only if no HTML has been processed
                    elif content_type == "text/plain":
                        bodyContent = part.get_payload(decode=True)
                        decoded_content = decode_email_content(bodyContent)
                        email_contents.append(decoded_content)
                        content_found = True

                # Join the email content into one
                full_email_content = "\n".join(email_contents)

                # Split references into a list
                references_list = [ref.strip() for ref in references.split() if ref.strip()]

                # Append to the final dictionary
                all_content["Datetime"].append(datetime_obj)
                all_content["Sender"].append(sender)
                all_content["Recipient"].append(recipient)
                all_content["Subject"].append(subject_cleaned)
                all_content["Content"].append(full_email_content.split("\n"))
                all_content["CC"].append(cc)
                all_content["BCC"].append(bcc)
                all_content["Message-ID"].append(message_id)
                all_content["In-Reply-To"].append(in_reply_to)
                all_content["References"].append(references_list)

    return pd.DataFrame(all_content)

def get_identifier(df):
    for i, row in df.iterrows():
        reference = row["References"]
        if reference:
            df.at[i, "References"] = reference[0]
        else:
            df.at[i, "References"] = None
    return df

def group_email(df):
    df['References'].fillna(df['Message-ID'], inplace=True)
    df = df.groupby(['References']).agg(
        {'Datetime': lambda x: x.tolist(),
        'Sender': lambda x: x.tolist(),
        'Recipient': lambda x: x.tolist(),
        'Subject': lambda x: x.tolist(),
        'Content': lambda x: x.tolist(),
        'CC': lambda x: x.tolist(),
        'BCC': lambda x: x.tolist()
        })
    df.reset_index(inplace=True,drop=True)
    return df

def get_last_content(df):
    for i, row in df.iterrows():
        datetime = row["Datetime"]
        sender = row["Sender"]
        recipient = row["Recipient"]
        subject = row["Subject"]
        content = row["Content"]
        cc = row["CC"]
        bcc = row["BCC"]
        if content :
            df.at[i, "Content"] = content[-1]
        else:
            df.at[i, "Content"] = None

        if datetime:
            df.at[i, "Datetime"] = datetime[-1]
        else:
            df.at[i, "Datetime"] = None

        if sender:
            df.at[i, "Sender"] = sender[-1]
        else:
            df.at[i, "Sender"] = None

        if recipient:
            df.at[i, "Recipient"] = recipient[-1]
        else:
            df.at[i, "Recipient"] = None

        if subject:
            df.at[i, "Subject"] = subject[-1]
        else:
            df.at[i, "Subject"] = None

        if cc:
            df.at[i, "CC"] = cc[-1]
        else:
            df.at[i, "CC"] = None

        if bcc:
            df.at[i, "BCC"] = bcc[-1]
        else:
            df.at[i, "BCC"] = None

    return df

def remove_utc(df, column_name):
    # Convert the column to datetime if it's not already
    df[column_name] = pd.to_datetime(df[column_name], utc=True)

    # Remove timezone information (convert to naive datetime)
    df[column_name] = df[column_name].dt.tz_localize(None)

    return df

def clean_text(text):
  # Remove the line breaks
  text = re.sub(r"\n"," ",text.lower())

  # Remove carriage returns
  text = re.sub(r"\r"," ",text)

  # Remove emoji
  text = re.sub(r'[^\x00-\x7F]+', ' ', text)

  # Remove whitespaces
  text = re.sub(r"\s+"," ",text)

  # Remove tabs
  text = re.sub(r"\t"," ",text)

  # Remove the forward message
  text = re.sub(r'---------- forwarded message ---------'," ",text)

  return text.strip()

def dataCleaning(df):
    df['Subject'] = df['Subject'].apply(clean_text)
    df['Content'] = df['Content'].apply(clean_text)
    df = remove_utc(df, 'Datetime')
    df = df[(df['Content'].str.strip() != "") & (df['Subject'].str.strip() != "")]
    df['Index'] = df.index

    return df

def cleanReply(text):
    cleaned_text = re.sub(r'>', '', text)
    cleaned_text = re.sub(r'<', '', cleaned_text)
    cleaned_text = re.sub(r'\n+', '\n', cleaned_text).strip()

    return cleaned_text

def cleanReplyContent(test1):
    replydf = test1[test1['Subject'].str.startswith('re:')]
    replydf['Content'] = replydf['Content'].apply(cleanReply)
    replydf = replydf[(replydf['Content'].str.strip() != "") & (replydf['Subject'].str.strip() != "")]
    replydf = replydf.groupby(['Index','Datetime','Sender','Recipient','Subject','CC','BCC']).agg(
        {'Content': ' '.join}
        ).reset_index()

    return replydf

def extract_reply_email(df):

    # Initialize a list to hold the extracted sections
    extracted_sections = {"Index":[],"Datetime": [],'Sender': [],  "Recipient": [], "Subject": [],"Content":[],"CC":[],"BCC":[]}

    for i, row in df.iterrows():
        index = row["Index"]
        email_content = row["Content"]
        subject = row["Subject"]
        sender = row["Sender"]
        recipient = row["Recipient"]
        datetime = row["Datetime"]
        cc = row["CC"]
        bcc = row["BCC"]

        pattern = r"(on\s(.+?)\sat\s(\d{1,2}:\d{2}\s(?:am|pm))\s(.+?)\swrote:)"  # Matches 'on [Day], [Date] at [Time am/pm] [Sender] wrote:'

        matches = re.findall(pattern, email_content, flags=re.IGNORECASE)

        # Iterate over the matches to extract datetime, sender, and message
        for i in range(len(matches)):
            match = matches[i]
            full_match, date, time, sender = match
            datetime_str = f"{date} at {time}"

            # For the current match, slice content after this match
            if i < len(matches) - 1:
                # Content between current match and next match
                next_full_match = matches[i + 1][0]
                content = email_content.split(full_match, 1)[1].split(next_full_match, 1)[0].strip()
            else:
                # For the last match, get all remaining content
                content = email_content.split(full_match, 1)[1].strip()

            # Append extracted data to the list
            extracted_sections['Index'].append(index)
            extracted_sections['Datetime'].append(datetime_str)
            extracted_sections['Sender'].append(sender)
            extracted_sections['Content'].append(content)
            extracted_sections['Recipient'].append(recipient)
            extracted_sections['Subject'].append(subject)
            extracted_sections['CC'].append(cc)
            extracted_sections['BCC'].append(bcc)

    return pd.DataFrame(extracted_sections)

def getReplyEmail(replydf):
    reply_df = extract_reply_email(replydf)
    reply_df['Datetime'] = pd.to_datetime(reply_df['Datetime'], errors='coerce')
    reply_df = reply_df.sort_values(by=['Index', 'Datetime'], ascending=[True, True]).reset_index(drop=True)

    return reply_df

def remove_content_before_from(df):
    pattern = r'from:'

    current_subject = None
    current_content_before_from = []
    text_to_remove = []

    for i, row in df.iterrows():
        subject = row["Subject"]
        content = row["Content"]

        # If this is the first row or the subject has changed, reset the content accumulator
        if current_subject is None or current_subject != subject:
            current_subject = subject
            current_content_before_from = []

        # Check if "from:" exists in the current row
        if re.search(pattern, content, re.IGNORECASE):
            text_to_remove.append(current_content_before_from)
            current_content_before_from = []
        else:
            # Accumulate content in the rows until "from:" is found, for the same subject
            current_content_before_from.append(content)

    return df, text_to_remove

def extract_forward_email(test1):
    extracted_info = {"Index":[],"Datetime": [],'Sender': [],  "Recipient": [], "Subject": [],"Content":[],"CC":[],"BCC":[]}

    current_subject = None
    current_content = []
    current_cc = ""
    current_bcc = ""
    current_sender = ""
    current_recipient = ""
    current_datetime = ""
    current_index = None

    regex_patterns = {
        'Sender': r'^from:\s*(.+)$',
        'Datetime': r'^(sent|date):\s*(.+)$',
        'Recipient': r'^to:\s*(.+)$',
        'Subject': r'^subject:\s*(.+)$',
        'CC': r'^cc:\s*(.+)$',
        'BCC': r'^bcc:\s*(.+)$'
    }

    for i, row in test1.iterrows():
        subject = row["Subject"].split(":")[1].strip().lower()
        content = row["Content"].strip()
        matched = False

        # Check if this is the first row (current_subject is None)
        if current_subject is None:
            current_subject = subject
            current_index = row.name

        # Check if a new subject starts, if so, save the last subject's content
        if current_subject != subject:
            extracted_info['Index'].append(current_index)
            extracted_info['Datetime'].append(current_datetime)
            extracted_info['Sender'].append(current_sender)
            extracted_info['Recipient'].append(current_recipient)
            extracted_info['Subject'].append(current_subject)
            extracted_info['Content'].append(" ".join(current_content).strip())
            extracted_info['CC'].append(current_cc.strip())
            extracted_info['BCC'].append(current_bcc.strip())

            # Reset for the new subject
            current_content = []
            current_cc = ""
            current_bcc = ""
            current_sender = ""
            current_recipient = ""
            current_datetime = ""

            # Update current subject to the new subject
            current_subject = subject
            current_index = row.name

        # Match against regex patterns
        for key, pattern in regex_patterns.items():
            match = re.match(pattern, content, re.IGNORECASE)
            if match:
                if key == 'Sender':
                    current_sender = match.group(1)
                elif key == 'Recipient':
                    current_recipient = match.group(1)
                elif key == 'Datetime':
                    current_datetime = match.group(2)  # Using group 2 for the combined 'sent' or 'date'
                elif key == 'CC':
                    current_cc = match.group(1)
                elif key == 'BCC':
                    current_bcc = match.group(1)
                matched = True
                break

        # If no pattern matches, add content to the current subject
        if not matched:
            current_content.append(content)

    # Append the last subject's content if there's any remaining
    if current_subject:
        extracted_info['Index'].append(current_index)
        extracted_info['Datetime'].append(current_datetime)
        extracted_info['Sender'].append(current_sender)
        extracted_info['Recipient'].append(current_recipient)
        extracted_info['Subject'].append(current_subject)
        extracted_info['Content'].append(" ".join(current_content).strip())
        extracted_info['CC'].append(current_cc.strip())
        extracted_info['BCC'].append(current_bcc.strip())

    # Ensure all lists are the same length
    max_length = max(len(extracted_info[key]) for key in extracted_info)
    for key in extracted_info:
        while len(extracted_info[key]) < max_length:
            extracted_info[key].append("")

    return pd.DataFrame(extracted_info)

def getForwardEmail(df):
    fw_emails_df = df[df['Subject'].str.startswith('fw:') | df['Subject'].str.startswith('fwd:')]

    # Get and remove the content before the 'from:'
    forward_df, current_content_before_from = remove_content_before_from(fw_emails_df)
    forward_df = forward_df[~forward_df['Content'].isin([item for sublist in current_content_before_from for item in sublist])]

    # Extract forward email fields from the cleaned content
    forward_df = extract_forward_email(forward_df)

    # Further cleanup if necessary
    if not forward_df.empty:
        forward_df = forward_df[(forward_df['Content'].str.strip() != "") & (forward_df['Subject'].str.strip() != "")]
        forward_df['Datetime'] = pd.to_datetime(forward_df['Datetime'], errors='coerce')
        remove_utc(forward_df, 'Datetime')

    return forward_df

# Group the normal inbox
def groupMainEmail(df):
    df = df.groupby(['Index']).agg({
        'Datetime': 'first',
        'Sender': 'first',
        'Recipient': 'first',
        'Subject': 'first',
        'CC': 'first',
        'BCC': 'first',
        'Content': ' '.join
        }).reset_index()

    return df

# Get the content of normal inbox
def getMainContent(df):
    pattern = r'^(.*?)(?=from:)'
    pattern_reply = r'^(.*?)(?=on)'
    for i, row in df.iterrows():
        subject = row["Subject"].strip().lower()
        content = row["Content"].strip().lower()
        if subject.startswith('fw:') or subject.startswith('fwd:'):
            match = re.match(pattern, content)
            if match:
                new_content = match.group(1).strip()
                df.at[i, 'Content'] = new_content
        if subject.startswith('re:') or subject.startswith('rep:'):
            match = re.match(pattern_reply, content)
            if match:
                new_content = match.group(1).strip()
                df.at[i, 'Content'] = new_content
    return df

def concatAllData(df,forward_df,reply_df):
    # Combine DataFrames
    updated_df = pd.concat([df, forward_df, reply_df], ignore_index=True)

    # Convert 'Datetime' column to datetime format
    updated_df['Datetime'] = pd.to_datetime(updated_df['Datetime'], errors='coerce')

    # Check if 'Index' is a column and sort by 'Index' and 'Datetime'
    if 'Index' in updated_df.columns:
        updated_df = updated_df.sort_values(by=['Index', 'Datetime'], ascending=[True, True]).reset_index(drop=True)
    else:
        updated_df = updated_df.sort_values(by=['Datetime'], ascending=True).reset_index(drop=True)

    # Group by 'Index' and aggregate columns into lists
    updated_df = updated_df.groupby('Index').agg({
        'Datetime': lambda x: x.tolist(),
        'Sender': lambda x: x.tolist(),
        'Recipient': lambda x: x.tolist(),
        'Subject': lambda x: x.tolist(),
        'Content': lambda x: x.tolist()
    }).reset_index()

    return updated_df

def clean_text2(text):
    # Convert text to lowercase and whitespaces
    text = re.sub(r'\s+', ' ', text.lower())

    # Replace newlines with spaces
    text = re.sub(r'\r', ' ', text)

    # Replace newlines with spaces
    text = re.sub(r'\n', ' ', text)

    # Replace tabs with spaces
    text = re.sub(r'\t', ' ', text)

    # Remove emoji
    text = re.sub(r'[^\x00-\x7F]+', '', text)

    # Remove control characters
    text = re.sub(r'[\x00-\x1F\x7F]', ' ', text)

    # Remove URLs
    text = re.sub(r'https?://\S+', ' ', text)
    text = re.sub(r'www\.\S+', ' ', text)

    # Remove non-alphabetic characters
    text = re.sub(r'[^a-zA-Z0-9\s]', " ", text)

    # Remove names of months from the text
    text = re.sub(r'\b(?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)\b', " ", text)

    # Remove common abbreviations and not applicable
    text = re.sub(r'\b(fy|na|n a|yr)\b', ' ', text)

    # Remove leading and trailing whitespace
    return text.strip()

def cleanText2(updated_df):
    updated_df['Content'] = updated_df['Content'].apply(lambda x: [clean_text(i) for i in x])
    updated_df = updated_df[(updated_df['Content'].str.strip() != "") & (updated_df['Subject'].str.strip() != "")]
    updated_df.drop("Index", axis=1, inplace=True)
    updated_df.reset_index(drop=True, inplace=True)

    return updated_df

def normalize_text(text,option):
    # Tokenize the text into words
    tokens = word_tokenize(text)

    # Keep only alphabetic words, remove numbers, punctuation
    tokens = [word for word in tokens if word.isalpha()]

    if option:
        # Remove common English stopwords
        tokens = [word for word in tokens if word not in stopwords.words('english')]

    # Initialize the lemmatizer
    lemmatizer = WordNetLemmatizer()

    # Lemmatize each word in the tokens list
    tokens = [lemmatizer.lemmatize(word) for word in tokens]

    # Join the tokens back into a single string and return
    return ' '.join(tokens)

def normalizeText(df,option):
    df['Content'] = df['Content'].apply(lambda x: [normalize_text(i,option) for i in x])
    df = df[(df['Content'].str.strip() != "") & (df['Subject'].str.strip() != "")]
    df['Subject'] = df['Subject'].apply(lambda x: ';'.join(x))
    df['Content'] = df['Content'].apply(lambda x: ';'.join(x))

    return df

def runAllProcess(user,passkey,key,value,domain,selection,option):
    df = extract_email_content(user,passkey,key,value,domain,selection)
    if df.empty:
        return df
    df = get_identifier(df)
    df = group_email(df)
    df = get_last_content(df)
    df = df.explode("Content")
    df = dataCleaning(df)
    df1 = df.copy()

    # Manage reply email
    replydf = cleanReplyContent(df1)
    reply_df = getReplyEmail(replydf)

    # Manage forward email
    forward_df = getForwardEmail(df)
    df = groupMainEmail(df)
    df = getMainContent(df)
    updated_df = concatAllData(df,forward_df,reply_df)
    updated_df = cleanText2(updated_df)
    updated_df = normalizeText(updated_df,option)

    return updated_df

from flask import Blueprint, request, send_file, render_template_string, render_template
import time
import os

email_extractor_bp = Blueprint('email-extractor',__name__)

CSV_FOLDER = '/tmp/csv_files'
os.makedirs(CSV_FOLDER, exist_ok=True)

@email_extractor_bp.route('/email')
def hello():
    return render_template("email.html")

@email_extractor_bp.route('/extract-email', methods=['POST'])
def extract():
    user = request.form.get('email_address')
    passkey = request.form.get('passkey')
    key = request.form.get('key')
    value = request.form.get('value')
    domain = request.form.get('domain')
    selection = request.form.get('selection')
    option = request.form.get('remove-stop')

    is_valid = validate_passkey(user, passkey, domain)
    if not is_valid:
        return render_template_string("""
            <html>
                <head>
                    <title>Email Extractor</title>
                    <link rel="stylesheet" type="text/css" href="/static/email.css">
                </head>
                <body>
                    <p>Invalid passkey. Please try again.</p>
                    <form method="GET" action="/email">
                        <button type="submit">Back to Extract</button>
                    </form>
                </body>
            </html>
        """)

    start_time = time.time()
    final_df = runAllProcess(user,passkey,key,value,domain,selection,option)
    if final_df.empty:
        return render_template_string("""
            <html>
                <head>
                    <title>Email Extractor</title>
                    <link rel="stylesheet" type="text/css" href="/static/email.css">
                </head>
                <body>
                    <p>No content to extract.</p>
                    <form method="GET" action="/email">
                        <button type="submit">Back to Extract</button>
                    </form>
                </body>
            </html>
        """)
    end_time = time.time()
    execution_time = end_time-start_time

    csv_filename = f"{key} {value}.csv"
    csv_path = os.path.join(CSV_FOLDER, csv_filename)
    final_df.to_csv(csv_path, index=False)

    return render_template_string("""
        <html>
            <head>
                <title>Download CSV</title>
                <link rel="stylesheet" type="text/css" href="/static/email.css">
            </head>
            <body>
                <div class="complete-container">
                    <h1>Processing Complete!!!</h1>
                    <p>Your Email has been processed. Click the button below to download the CSV file.</p>
                    <div class="download-card">
                        <p>The executed time is <strong>{{ execution_time }}</strong> seconds!</p>
                        <form method="GET" action="/download_csv">
                            <input type="hidden" name="filename" value="{{ csv_filename }}">
                            <button type="submit">Download CSV</button>
                        </form>
                    </div>
                    <form method="GET" action="/email">
                        <button type="submit">Back to Extract</button>
                    </form>
                </div>
            </body>
        </html>
    """, csv_filename=csv_filename,execution_time=execution_time)

@email_extractor_bp.route('/download_csv', methods=['GET'])
def download_csv():
    csv_filename = request.args.get('filename')
    csv_path = os.path.join(CSV_FOLDER, csv_filename)

    if os.path.exists(csv_path):
        return send_file(csv_path, as_attachment=True)
    else:
        return "File not found", 404
    
