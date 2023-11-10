from io import BytesIO
import pandas as pd
from flask import Flask, request, send_file, make_response
from lxml import html
import re

app = Flask(__name__)

def parse_xml_with_lxml(xml_data):
    try:
        xml_data_cleaned = re.sub(r'</?xml>', '', xml_data)
        tree = html.fromstring(xml_data_cleaned)

        qa_pairs = {}
        dts = tree.xpath('//dt')
        for dt in dts:
            question = dt.text.strip()
            dd = dt.getnext()
            answer = dd.text_content().strip() if dd is not None else ''
            qa_pairs[question] = answer
        return qa_pairs
    except Exception as e:
        return {}

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        file = request.files.get('file')
        if file:
            df = pd.read_excel(file)

            if 'body/en' in df.columns:
                df['parsed_xml_en'] = df['body/en'].apply(parse_xml_with_lxml)

            if 'body/fr' in df.columns:
                df['parsed_xml_fr'] = df['body/fr'].apply(parse_xml_with_lxml)

            unique_questions = set()
            for col in ['parsed_xml_en', 'parsed_xml_fr']:
                if col in df.columns:
                    for parsed_data in df[col]:
                        unique_questions.update(parsed_data.keys())

            for question in unique_questions:
                df[question] = ''

            if 'parsed_xml_en' in df.columns:
                for index, row in df.iterrows():
                    for question, answer in row['parsed_xml_en'].items():
                        if question in df.columns:
                            df.at[index, question] = answer

            if 'parsed_xml_fr' in df.columns:
                for index, row in df.iterrows():
                    for question, answer in row['parsed_xml_fr'].items():
                        if question in df.columns:
                            df.at[index, question] = answer

            df.drop(columns=['parsed_xml_en', 'parsed_xml_fr', 'body/en', 'body/fr'], inplace=True)

            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, sheet_name='Sheet1', index=False)

            output.seek(0)
            response = make_response(send_file(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', as_attachment=True, download_name='processed_file.xlsx'))
            return response

    return '''
    <!doctype html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <title>Upload Excel File</title>
        <!-- Styles -->
    </head>
    <body>
        <div class="container">
            <h1>Upload Excel File</h1>
            <form method="post" enctype="multipart/form-data">
                <div class="upload-btn-wrapper">
                    <button class="btn">Choose a file</button>
                    <input type="file" name="file" />
                </div>
                <br><br>
                <input type="submit" value="Upload">
            </form>
        </div>
    </body>
    </html>
    '''

if __name__ == '__main__':
    app.run(debug=True)
