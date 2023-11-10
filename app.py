from io import BytesIO
import pandas as pd
from flask import Flask, request, send_file, make_response
import pandas as pd
from io import BytesIO
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
        if 'file' not in request.files:
            return 'No file part'
        file = request.files['file']
        if file.filename == '':
            return 'No selected file'
        if file:
            df = pd.read_excel(file)

            # Process the XML data in the 'body/en' column
            df['parsed_xml'] = df['body/en'].apply(parse_xml_with_lxml)

            # Identifying all unique questions
            unique_questions = set()
            for parsed_data in df['parsed_xml']:
                unique_questions.update(parsed_data.keys())

            # Creating new columns for each unique question
            for question in unique_questions:
                df[question] = ''

            # Populating the new columns with answers
            for index, row in df.iterrows():
                for question, answer in row['parsed_xml'].items():
                    if question in df.columns:
                        df.at[index, question] = answer

            # Dropping the original and parsed_xml columns
            df.drop(columns=['parsed_xml', 'body/en'], inplace=True)

            # Write the processed DataFrame to a BytesIO buffer
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, sheet_name='Sheet1', index=False)

            output.seek(0)
            
            # Sending the processed file as a response
            response = make_response(send_file(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', as_attachment=True, download_name='processed_file.xlsx'))
            return response

    return '''
    <!doctype html>
    <title>Upload Excel File</title>
    <h1>Upload Excel File</h1>
    <form method=post enctype=multipart/form-data>
      <input type=file name=file>
      <input type=submit value=Upload>
    </form>
    '''

if __name__ == '__main__':
    app.run(debug=True)