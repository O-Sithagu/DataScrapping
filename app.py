
from flask import Flask, request, render_template
import pandas as pd
from openpyxl import load_workbook

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('upload.html')

@app.route('/upload', methods=['POST'])
def upload():
    file = request.files['excelFile']
    df = pd.read_excel(file)
    
    template = 'template.xlsx'
    wb = load_workbook(template)
    ws = wb.active

    for row in df.itertuples(index=False):
        ws.append(row)

    wb.save('updated_template.xlsx')
    return "File processed successfully!"

if __name__ == '__main__':
    app.run()
