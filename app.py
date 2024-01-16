from flask import Flask, render_template, request, send_file
import os
import pandas as pd
import re
from newspaper import Article
from datetime import datetime, timedelta
import xlsxwriter
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'static/files'



def extract_links(message):
    link_pattern = r'https?://[^\s]+'
    links = re.findall(link_pattern, message)
    return links

def extract_title(url):
    try:
        article = Article(url)
        ## Download HTML yourself and insert into Newspaper3k
            ## response = requests.get(url)
            ##article.download(input_html=response.text)
        article.download()
        article.parse()
        title = article.title.strip()
        return title
    except Exception as e:
        print(f"Error extracting title for {url}: {e}")
        return ''

def extract_left_members(message):
    leave_pattern = r'(.+?) left$'
    match = re.match(leave_pattern, message)
    return match.group(1).strip() if match else None

def extract_added_members(message):
    added_pattern = r'(.+?) added (.+)$'
    match = re.match(added_pattern, message)
    return match.groups() if match else None

def extract_removed_members(message):
    removed_pattern = r'(.+?) removed (.+)$'
    match = re.match(removed_pattern, message)
    return match.groups() if match else None

def parse_whatsapp_chat(file_path, start_time, end_time):
    with open(file_path, 'r', encoding='utf-8') as file:
        lines = file.readlines()

    link_data = {'Sender name': [], 'Link': [], 'Title': []}
    left_data = {'Left Member': []}
    added_data = {'Added Member': [], 'Added By': []}
    removed_data = {'Removed Member': [], 'Removed By': []}
    current_sender = None

    for line in lines:
        message_match = re.match(r'\[([^,]+), ([^]]+)\] (.+?): (.+)', line)
        if message_match:
            date, time, sender, message_content = message_match.groups()
            current_sender = sender.lstrip("~ ").strip()

            try:
                message_datetime = datetime.strptime(f'{date} {time}', '%y/%m/%d %H:%M:%S')
            except ValueError:
                message_datetime = datetime.strptime(f'{date} {time}', '%m/%d/%y %H:%M:%S')

            if start_time <= message_datetime <= end_time:
                left_member = extract_left_members(message_content)
                if left_member:
                    left_data['Left Member'].append(left_member)

                added_member = extract_added_members(message_content)
                if added_member:
                    added_data['Added Member'].append(added_member[0].strip())
                    added_data['Added By'].append(current_sender)

                removed_member = extract_removed_members(message_content)
                if removed_member:
                    removed_data['Removed Member'].append(removed_member[0].strip())
                    removed_data['Removed By'].append(current_sender)

                links = extract_links(message_content)
                for link in links:
                    title = extract_title(link)
                    link_data['Sender name'].append(current_sender)
                    link_data['Link'].append(link)
                    link_data['Title'].append(title)

    return pd.DataFrame(link_data), pd.DataFrame(left_data), pd.DataFrame(added_data), pd.DataFrame(removed_data)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/', methods=['POST'])
def upload_file():
    if request.method == 'POST':
        uploaded_file = request.files['file']
        selected_date = request.form['date']
        file = uploaded_file

        file_contents = uploaded_file.stream.read().decode("utf-8")


        if uploaded_file.filename != '':
            # Convert selected date to datetime object
            selected_date = datetime.strptime(selected_date, '%Y-%m-%d')

            # Set start_time to 12:00 PM and end_time to 12:00 PM on the selected date
            start_time = (selected_date - timedelta(days=1)).replace(hour=12, minute=0, second=0, microsecond=0)
            end_time = selected_date.replace(hour=12, minute=0, second=0, microsecond=0)

            #file_path = os.path.join('uploads', uploaded_file.filename)
            #file_path2 = os.path.join(os.path.abspath(os.path.dirname(__file__)),app.config['UPLOAD_FOLDER'],secure_filename(file.filename))
            file_path = os.path.join(app.root_path, app.config['UPLOAD_FOLDER'], uploaded_file.filename)
            uploaded_file.stream.seek(0)
            uploaded_file.save(file_path)
            #uploaded_file.save(os.path.join(app.root_path, app.config['UPLOAD_FOLDER'], uploaded_file.filename))
            #file.save(file_path2) # Then save the file
           # uploaded_file.save(file_path)
          #  uploaded_file.save(file_path2)

        ## current_date_time = datetime.now()
        ## start_time = current_date_time - timedelta(days=1)
        ## start_time = start_time.replace(hour=12, minute=0, second=0, microsecond=0)
        ## end_time = current_date_time.replace(hour=12, minute=0, second=0, microsecond=0)

        links_df, left_df, added_df, removed_df = parse_whatsapp_chat(file_path, start_time, end_time)

        excel_output_file = os.path.join(app.root_path, f'whatsapp_data_{selected_date.strftime("%Y-%m-%d")}.xlsx')
        with pd.ExcelWriter(excel_output_file, engine='xlsxwriter') as writer:
            links_df.to_excel(writer, sheet_name='Links', index=False)
            left_df.to_excel(writer, sheet_name='Left Members', index=False)
            added_df.to_excel(writer, sheet_name='Added Members', index=False)
            removed_df.to_excel(writer, sheet_name='Removed Members', index=False)

            workbook = writer.book
            # Inside the Excel writing loop for all tabs
            for sheet_name in writer.sheets:
                    sheet = writer.sheets[sheet_name]
                    data_df = None  # Initialize the DataFrame to None

                    if sheet_name == 'Links':
                        data_df = links_df
                    elif sheet_name == 'Left Members':
                        data_df = left_df
                    elif sheet_name == 'Added Members':
                        data_df = added_df
                    elif sheet_name == 'Removed Members':
                        data_df = removed_df

                    if data_df is not None:
                        for column_name in data_df.columns:
                            column_index = data_df.columns.get_loc(column_name)
                            if sheet_name == 'Links' and column_name == 'Link':
                                # Check if the current column is the 'Link' column in the 'Links' tab
                                for idx, link in enumerate(data_df[column_name], start=2):
                                    sheet.write_url(idx - 1, column_index, link)
                            else:
                                # For other columns or tabs, simply write the data without hyperlinks
                                sheet.write_column(1, column_index, data_df[column_name])

        # Debugging statements
        print("File created successfully:", excel_output_file)
        print("Is file exists?", os.path.exists(excel_output_file))

        return send_file(excel_output_file, as_attachment=True)

    return render_template('index.html', error='Please upload a file.')


if __name__ == '__main__':
    os.makedirs('uploads', exist_ok=True)
    app.run(debug=True)
