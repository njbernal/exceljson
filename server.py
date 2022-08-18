import os
import openpyxl
from flask import Flask, request
from flask_cors import CORS, cross_origin

from openpyxl import Workbook
from openpyxl.utils import get_column_letter

app = Flask(__name__)
cors = CORS(app)
app.config['CORS_HEADERS'] = 'Content-Type'


@app.route('/')
def hello():
    return 'Hello there.'


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=True)
