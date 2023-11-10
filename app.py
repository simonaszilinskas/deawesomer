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
        # ... [existing code]
        if file:
            df = pd.read_excel(file)

            # Process 'body/en' 
            if 'body/en' in df.columns:
                df['parsed_xml_en'] = df['body/en'].apply(parse_xml_with_lxml)

                # Add logic to process df['parsed_xml_en'] 
                # and extract unique questions and answers

            # Process 'body/fr'
            if 'body/fr' in df.columns:
                df['parsed_xml_fr'] = df['body/fr'].apply(parse_xml_with_lxml)

                # Add logic to process df['parsed_xml_fr'] 
                # and extract unique questions and answers

            # Example: Combine or keep separate - depends on your requirement
            # Here's a simple example to combine the unique questions from both
            unique_questions_en = set()
            unique_questions_fr = set()

            if 'parsed_xml_en' in df.columns:
                for parsed_data in df['parsed_xml_en']:
                    unique_questions_en.update(parsed_data.keys())

            if 'parsed_xml_fr' in df.columns:
                for parsed_data in df['parsed_xml_fr']:
                    unique_questions_fr.update(parsed_data.keys())

            # Combine unique questions from both languages (if this is your requirement)
            unique_questions = unique_questions_en.union(unique_questions_fr)

            # Creating new columns for each unique question
            for question in unique_questions:
                df[question] = ''

            # Populating the new columns with answers for both languages
            # Adjust this logic based on how you want to handle the languages
            # This is an example where we fill data from both columns into the same DataFrame
            if 'parsed_xml_en' in df.columns:
                for index, row in df.iterrows():
                    for question, answer in row['parsed_xml_en'].items():
                        if question in df.columns:
                            df.at[index, question] = answer

            if 'parsed_xml_fr' in df.columns:
                for index, row in df.iterrows():
                    for question, answer in row['parsed_xml_fr'].items():
                        if question in df.columns:
                            # Example logic - you can adjust it based on your needs
                            # This will overwrite the English data if French data is available
                            df.at[index, question] = answer

            # Dropping the original and parsed_xml columns
            df.drop(columns=['parsed_xml_en', 'parsed_xml_fr', 'body/en', 'body/fr'], inplace=True)

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
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>Upload Excel File</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            background-color: #f4f4f4;
            display: flex;
            justify-content: center;
            align-items: center;
            height: 100vh;
            margin: 0;
        }
        .container {
            background-color: white;
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
        }
        h1 {
            color: #333;
        }
        .upload-btn-wrapper {
            position: relative;
            overflow: hidden;
            display: inline-block;
        }
        .btn {
            border: 2px solid gray;
            color: gray;
            background-color: white;
            padding: 8px 20px;
            border-radius: 8px;
            font-size: 16px;
            font-weight: bold;
        }
        .upload-btn-wrapper input[type=file] {
            font-size: 100px;
            position: absolute;
            left: 0;
            top: 0;
            opacity: 0;
        }
    </style>
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