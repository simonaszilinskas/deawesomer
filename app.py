from io import BytesIO
import pandas as pd
from flask import Flask, request, send_file, make_response
from lxml import html
import re
from bs4 import BeautifulSoup

app = Flask(__name__)

def contains_html_tags(text):
    """Check if a string contains HTML tags."""
    if not isinstance(text, str):
        return False
    # Look for anything that looks like an HTML tag
    return bool(re.search(r'<[^>]+>', text))

def clean_html_content(text):
    """Clean HTML content thoroughly."""
    if not isinstance(text, str):
        return ''
    
    # Convert to BeautifulSoup object
    soup = BeautifulSoup(text, 'html.parser')
    
    # Get text content and clean up whitespace
    cleaned_text = soup.get_text(separator=' ')
    cleaned_text = re.sub(r'\s+', ' ', cleaned_text)
    cleaned_text = cleaned_text.strip()
    
    # Additional cleaning to catch any remaining tags
    cleaned_text = re.sub(r'<[^>]+>', '', cleaned_text)
    
    return cleaned_text

def verify_and_clean_dataframe(df):
    """Verify all columns for HTML tags and clean if necessary."""
    for column in df.columns:
        # Check each cell in the column
        html_found = df[column].astype(str).apply(contains_html_tags).any()
        
        if html_found:
            print(f"Found HTML tags in column: {column}")
            # Clean the column
            df[column] = df[column].astype(str).apply(clean_html_content)
    
    return df

def parse_xml_with_lxml(xml_data):
    try:
        if not isinstance(xml_data, str):
            return {}
            
        xml_data_cleaned = re.sub(r'</?xml>', '', xml_data)
        tree = html.fromstring(xml_data_cleaned)

        qa_pairs = {}
        dts = tree.xpath('//dt')
        for dt in dts:
            question = clean_html_content(dt.text) if dt.text else ''
            
            dd = dt.getnext()
            if dd is not None:
                answer_html = html.tostring(dd, encoding='unicode')
                answer = clean_html_content(answer_html)
            else:
                answer = ''
                
            if question:
                qa_pairs[question] = answer
                
        return qa_pairs
    except Exception as e:
        print(f"Error parsing XML: {str(e)}")
        return {}

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        file = request.files.get('file')
        if file:
            try:
                df = pd.read_excel(file)

                # Process XML columns
                if 'body/en' in df.columns:
                    df['parsed_xml_en'] = df['body/en'].apply(parse_xml_with_lxml)

                if 'body/fr' in df.columns:
                    df['parsed_xml_fr'] = df['body/fr'].apply(parse_xml_with_lxml)

                # Extract unique questions
                unique_questions = set()
                for col in ['parsed_xml_en', 'parsed_xml_fr']:
                    if col in df.columns:
                        for parsed_data in df[col]:
                            unique_questions.update(parsed_data.keys())

                # Create and fill new columns
                sorted_questions = sorted(unique_questions)
                for question in sorted_questions:
                    df[question] = ''

                # Fill in answers
                for col_source, parsed_col in [('parsed_xml_en', 'parsed_xml_en'), 
                                            ('parsed_xml_fr', 'parsed_xml_fr')]:
                    if parsed_col in df.columns:
                        for index, row in df.iterrows():
                            for question, answer in row[parsed_col].items():
                                if question in df.columns:
                                    df.at[index, question] = answer

                # Drop processing columns
                columns_to_drop = ['parsed_xml_en', 'parsed_xml_fr', 'body/en', 'body/fr']
                for col in columns_to_drop:
                    if col in df.columns:
                        df.drop(columns=[col], inplace=True)

                # Verify and clean any remaining HTML tags
                df = verify_and_clean_dataframe(df)

                # Export to Excel
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df.to_excel(writer, sheet_name='Sheet1', index=False)

                output.seek(0)
                response = make_response(send_file(
                    output,
                    mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    as_attachment=True,
                    download_name='processed_file.xlsx'
                ))
                return response
                
            except Exception as e:
                return f'Error processing file: {str(e)}', 400

    return '''
        <!doctype html>
        <html lang="fr">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>Deawesomer - Traitement des exports Decidim</title>
            <style>
                body {
                    font-family: Arial, sans-serif;
                    background-color: #f4f4f4;
                    display: flex;
                    justify-content: center;
                    align-items: center;
                    min-height: 100vh;
                    margin: 0;
                    padding: 20px;
                }
                .container {
                    background-color: white;
                    padding: 2rem;
                    border-radius: 12px;
                    box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
                    max-width: 800px;
                    width: 100%;
                }
                .header {
                    text-align: center;
                    margin-bottom: 2rem;
                }
                h1 {
                    color: #2c3e50;
                    font-size: 2rem;
                    margin-bottom: 0.5rem;
                }
                .description {
                    color: #34495e;
                    line-height: 1.6;
                    margin-bottom: 2rem;
                    text-align: left;
                }
                .upload-section {
                    background-color: #f8f9fa;
                    padding: 2rem;
                    border-radius: 8px;
                    text-align: center;
                }
                .upload-btn-wrapper {
                    position: relative;
                    overflow: hidden;
                    display: inline-block;
                    margin-bottom: 1rem;
                }
                .btn {
                    border: 2px solid #3498db;
                    color: #3498db;
                    background-color: white;
                    padding: 10px 24px;
                    border-radius: 8px;
                    font-size: 16px;
                    font-weight: bold;
                    cursor: pointer;
                    transition: all 0.3s ease;
                }
                .btn:hover {
                    background-color: #3498db;
                    color: white;
                }
                .upload-btn-wrapper input[type=file] {
                    font-size: 100px;
                    position: absolute;
                    left: 0;
                    top: 0;
                    opacity: 0;
                    cursor: pointer;
                }
                .submit-btn {
                    background-color: #2ecc71;
                    color: white;
                    border: none;
                    padding: 10px 24px;
                    border-radius: 8px;
                    font-size: 16px;
                    font-weight: bold;
                    cursor: pointer;
                    transition: all 0.3s ease;
                }
                .submit-btn:hover {
                    background-color: #27ae60;
                }
                .footer {
                    text-align: center;
                    margin-top: 2rem;
                    color: #7f8c8d;
                    font-size: 0.9rem;
                }
                .footer a {
                    color: #3498db;
                    text-decoration: none;
                }
                .footer a:hover {
                    text-decoration: underline;
                }
                .selected-file {
                    margin-top: 1rem;
                    color: #7f8c8d;
                    font-size: 0.9rem;
                }
            </style>
            <script>
                function updateFileName() {
                    const input = document.querySelector('input[type=file]');
                    const fileNameDisplay = document.getElementById('fileName');
                    if (input.files.length > 0) {
                        fileNameDisplay.textContent = 'Fichier sélectionné : ' + input.files[0].name;
                    }
                }
            </script>
        </head>
        <body>
            <div class="container">
                <div class="header">
                    <h1>Deawesomer</h1>
                    <div class="description">
                        <p>
                            <strong>Description :</strong> Cet outil transforme les exports de propositions utilisant les champs personnalisés 
                            de Decidim Awesome en un format plus lisible et exploitable.
                        </p>
                        <p>
                            Pour utiliser l'outil :
                            <ol>
                                <li>Exportez vos propositions depuis Decidim au format Excel (.xlsx)</li>
                                <li>Téléversez le fichier via le bouton ci-dessous</li>
                                <li>Un nouveau fichier Excel traité sera automatiquement téléchargé</li>
                            </ol>
                        </p>
                    </div>
                </div>
                
                <div class="upload-section">
                    <form method="post" enctype="multipart/form-data">
                        <div class="upload-btn-wrapper">
                            <button class="btn">Choisir un fichier</button>
                            <input type="file" name="file" accept=".xlsx,.xls" onchange="updateFileName()" />
                        </div>
                        <div id="fileName" class="selected-file"></div>
                        <br>
                        <input type="submit" value="Traiter le fichier" class="submit-btn">
                    </form>
                </div>

                <div class="footer">
                    <p>
                        Créé par l'équipe <a href="https://opensourcepolitics.eu" target="_blank">Open Source Politics</a>
                        <br>
                        <a href="https://github.com/simonaszilinskas/deawesomer" target="_blank">Code source disponible sur GitHub</a>
                    </p>
                </div>
            </div>
        </body>
        </html>
        '''

if __name__ == '__main__':
    app.run(debug=True)