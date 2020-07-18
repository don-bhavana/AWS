from pymongo import MongoClient
import flask
from flask import jsonify, request
from bson.json_util import dumps
import openpyxl


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
#IN-COUNTRY VALUE DETAILS
]

wb = openpyxl.load_workbook("ICV-certificate-filled.xlsx", data_only=True)
sheet = wb["Page 1 - ICV Summary"]
x = "C10"
for i in XLS_CELL_ADDRESS:
	print(sheet[i].value)


client = MongoClient('mongodb://localhost:27017/')
db = client.pymongo_test #db
posts = db.posts # collection
app = flask.Flask(__name__)

@app.route('/', methods=['GET'])
def home():
    scotts_posts = posts.find({'author': 'Scott'})
    post = dumps(scotts_posts)
   
    return jsonify(post)

app.run()

