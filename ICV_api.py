from pymongo import MongoClient
import pymongo
import flask
from flask import jsonify, request
import openpyxl
from datetime import datetime

XLS_CELL_ADDRESS = [
# COMPANY GENERAL INFORMATION
					'C11', # Name (as per License)
					'G11', # License No.
					'C14', # Address (as per License) - line1
					'C15', # line2
					'C16', # line3
					'C17', # line4
					'E14', # Contact Person Details: Name
					'E15', # E-mail
					'E16', # Office
					'E17', # Mobile
					'D18', # Company Type: drop down list, selected value
					'H18', # Company based in drop down list:[], selected value
					'D21', #Financial Year End date (mm/dd/yyyy)
					'G21', #Audited Financial Statement Issue date (mm/dd/yyyy)
					'D24'#Company Business drop down list:[], selected value
]

def create_dict():
    wb = openpyxl.load_workbook("ICV-certificate-filled.xlsx", data_only=True)
    sheet = wb["Page 1 - ICV Summary"]
    page = {}
    page['label'] = 'Page 1 - ICV Summary'
    page['Company general information'] = {}
    
    page['Company general information']['Name (as per License)'] = {}
    page['Company general information']['Name (as per License)']['value'] = sheet[XLS_CELL_ADDRESS[0]].value
    page['Company general information']['Name (as per License)']['type'] = 'string'
    page['Company general information']['Name (as per License)']['validated'] = 'boolean'
    page['Company general information']['Name (as per License)']['comments'] = 'string'


    page['Company general information']['License No'] = {}
    page['Company general information']['License No']['value'] = sheet[XLS_CELL_ADDRESS[1]].value
    page['Company general information']['License No']['type'] = 'string'
    page['Company general information']['License No']['validated'] = 'boolean'
    page['Company general information']['License No']['comments'] = 'string'


    page['Company general information']['Address (as per License)'] = {}
    page['Company general information']['Address (as per License)']['line1'] = {}
    page['Company general information']['Address (as per License)']['line1']['value'] = sheet[XLS_CELL_ADDRESS[2]].value
    page['Company general information']['Address (as per License)']['line1']['type'] = 'string'

    page['Company general information']['Address (as per License)']['line2'] = {}
    page['Company general information']['Address (as per License)']['line2']['value'] = sheet[XLS_CELL_ADDRESS[3]].value
    page['Company general information']['Address (as per License)']['line2']['type'] = 'string'
    
    page['Company general information']['Address (as per License)']['line3'] = {}
    page['Company general information']['Address (as per License)']['line3']['value'] = sheet[XLS_CELL_ADDRESS[4]].value
    page['Company general information']['Address (as per License)']['line3']['type'] = 'string'
    
    page['Company general information']['Address (as per License)']['line4'] = {}
    page['Company general information']['Address (as per License)']['line4']['value'] = sheet[XLS_CELL_ADDRESS[5]].value
    page['Company general information']['Address (as per License)']['line4']['type'] = 'string'

    page['Company general information']['Address (as per License)']['validated'] = 'boolean'
    page['Company general information']['Address (as per License)']['comments'] = 'string'

    page['Company general information']['Contact Person Details'] = {}
    page['Company general information']['Contact Person Details']['Name'] = {}
    page['Company general information']['Contact Person Details']['Name']['value'] = sheet[XLS_CELL_ADDRESS[6]].value
    page['Company general information']['Contact Person Details']['Name']['type'] = 'string'
    page['Company general information']['Contact Person Details']['Name']['validated'] = 'boolean'
    page['Company general information']['Contact Person Details']['Name']['comments'] = 'string'

    page['Company general information']['Contact Person Details']['E-mail'] = {}
    page['Company general information']['Contact Person Details']['E-mail']['value'] = sheet[XLS_CELL_ADDRESS[7]].value
    page['Company general information']['Contact Person Details']['E-mail']['type'] = 'string'
    page['Company general information']['Contact Person Details']['E-mail']['validated'] = 'boolean'
    page['Company general information']['Contact Person Details']['E-mail']['comments'] = 'string'

    page['Company general information']['Contact Person Details']['Office'] = {}
    page['Company general information']['Contact Person Details']['Office']['value'] = sheet[XLS_CELL_ADDRESS[8]].value
    page['Company general information']['Contact Person Details']['Office']['type'] = 'string'
    page['Company general information']['Contact Person Details']['Office']['validated'] = 'boolean'
    page['Company general information']['Contact Person Details']['Office']['comments'] = 'string'

    page['Company general information']['Contact Person Details']['Mobile'] = {}
    page['Company general information']['Contact Person Details']['Mobile']['value'] = sheet[XLS_CELL_ADDRESS[9]].value
    page['Company general information']['Contact Person Details']['Mobile']['type'] = 'Number'
    page['Company general information']['Contact Person Details']['Mobile']['validated'] = 'boolean'
    page['Company general information']['Contact Person Details']['Mobile']['comments'] = 'string'

    page['Company general information']['Company Type'] = {}
    page['Company general information']['Company Type']['value'] = sheet[XLS_CELL_ADDRESS[10]].value
    page['Company general information']['Company Type']['Dropdown list'] = ['SME in UAE', 'Non SME in UAE', 'International Company']
    page['Company general information']['Company Type']['type'] = 'Dropdown'
    page['Company general information']['Company Type']['validated'] = 'boolean'
    page['Company general information']['Company Type']['comments'] = 'string'

    
    page['Company general information']['Company based in'] = {}
    page['Company general information']['Company Type']['value'] = sheet[XLS_CELL_ADDRESS[11]].value
    page['Company general information']['Company Type']['Dropdown list'] = ['Outside UAE', 'Within UAE']
    page['Company general information']['Company Type']['type'] = 'dropdown'
    page['Company general information']['Company Type']['validated'] = 'boolean'
    page['Company general information']['Company Type']['comments'] = 'string'


    page['Company general information']['Financial Year End date (mm/dd/yyyy)'] = {}
    page['Company general information']['Financial Year End date (mm/dd/yyyy)']['value'] = (sheet[XLS_CELL_ADDRESS[12]].value).isoformat()
    page['Company general information']['Financial Year End date (mm/dd/yyyy)']['type'] = 'datetime'
    page['Company general information']['Financial Year End date (mm/dd/yyyy)']['validated'] = 'boolean'
    page['Company general information']['Financial Year End date (mm/dd/yyyy)']['comments'] = 'string'


    page['Company general information']['Audited Financial Statement Issue date (mm/dd/yyyy)'] = {}
    page['Company general information']['Audited Financial Statement Issue date (mm/dd/yyyy)']['value'] = (sheet[XLS_CELL_ADDRESS[12]].value).isoformat()
    page['Company general information']['Audited Financial Statement Issue date (mm/dd/yyyy)']['type'] = 'datetime'
    page['Company general information']['Audited Financial Statement Issue date (mm/dd/yyyy)']['validated'] = 'boolean'
    page['Company general information']['Audited Financial Statement Issue date (mm/dd/yyyy)']['comments'] = 'string'

    page['Company general information']['Company Business'] = {}
    page['Company general information']['Company Business']['value'] = sheet[XLS_CELL_ADDRESS[11]].value
    page['Company general information']['Company Business']['Dropdown list'] = ['GOOD MANUFACTURER', 'SERVICE PROVIDER']
    page['Company general information']['Company Business']['type'] = 'dropdown'
    page['Company general information']['Company Business']['validated'] = 'boolean'
    page['Company general information']['Company Business']['comments'] = 'string'

    
    
    return page


client = MongoClient('mongodb://localhost:27017/')
db = client.icv_db #db
icv = db.icv # collection

app = flask.Flask(__name__)

@app.route('/', methods=['GET'])
def home():

    post = create_dict()
    result = icv.insert_one(post)
    id = result.inserted_id
    
    posts = list(icv.find({'_id': id}))
    for record in posts:
        record['_id'] = str(record.pop('_id'))
    return jsonify(posts), 200

app.run(host="0.0.0.0")

