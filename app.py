import openpyxl
from flask import Flask, render_template, request, redirect, url_for

file_path = 'C:\\Users\\Chauh\\Documents\\ThaneFilteredEngg.xlsx'
workbook = openpyxl.load_workbook(file_path)
sheet = workbook.active

app = Flask(__name__)

def find_info(query):
    if query.isdigit() and len(query) == 5:
        roll_no = 'P0' + query
        for row in sheet.iter_rows(values_only=True):
            if row[0] == roll_no:
                return [{"name": row[1], "roll": row[0]}]
    else:
        matches = []
        for row in sheet.iter_rows(values_only=True):
            if query.lower() in row[1].lower():
                matches.append({"name": row[1], "roll": row[0]})
        return matches

    return []

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/search', methods=['POST'])
def search():
    query = request.form['query']
    info = find_info(query)
    if info:
        return redirect(url_for('results', query=query))
    else:
        return "No matching records found"

@app.route('/results')
def results():
    query = request.args.get('query')
    info = find_info(query)
    return render_template('search_results.html', info=info)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8080)
