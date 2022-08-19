import os
import json
from openpyxl import load_workbook
from flask import Flask, request
from flask_cors import CORS, cross_origin

from openpyxl import Workbook
from openpyxl.utils import get_column_letter

app = Flask(__name__)
cors = CORS(app)
app.config['CORS_HEADERS'] = 'Content-Type'


def excel(source, ignore_blank_row=False, count=-1):
    wb = load_workbook(source)
    sheets = wb.sheetnames
    json_obj = []
    for sheet in sheets:
        limit = count
        worksheet = wb[sheet]
        iterable = worksheet.iter_rows()
        current = []
        while True and limit != 0:
            try:
                row = next(iterable)
                row_list = []
                for cell in row:
                    row_list.append(cell.value)
            except StopIteration:
                break
            if ignore_blank_row is True:
                check = set(row_list)
                if len(check) == 1 and None in check:
                    continue
            limit -= 1
            current.append(row_list)
        json_obj.append({sheet: current})

    return json_obj


@app.route('/', methods=['GET', 'POST'])
@cross_origin()
def hello():
    if request.method == 'POST':
        print(request.form['ignore_blank_rows'])
        ignore = request.form['ignore_blank_rows']

        source = request.files['source']
        result = excel(source, False, 5)
        # output = json.dumps(result, indent=4)
        return result

    return 'Hello there.'


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    print(port)
    app.run()
