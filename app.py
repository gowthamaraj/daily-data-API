from flask import Flask,jsonify,request,send_from_directory
import json
import requests
from datetime import datetime as dt
import xlsxwriter
import os

# getting the data from the machstatz API
data = json.loads(requests.get("https://assignment-machstatz.herokuapp.com/excel").text) 

#calling the constructor to create the app
app = Flask(__name__) 

#endpoint for GET / 
@app.route('/',methods=['GET'])
def index():
    return jsonify(data)

#endpoint for GET /total
@app.route('/total',methods=['GET'])
def res():
    if 'day' in request.args:
        try:
            day = request.args['day']
            day = dt.strptime(day, '%d-%m-%Y')
            #conv. the string given by user to datetime object
            if len(data) > 0:
                #check if the data is received from the machstatz API
                totalWeight = 0
                totalLength = 0
                totalQuantity = 0
                for item in data:
                    item_copy = item.copy()
                    #make a copy of the item to not spoil the original data
                    item_date = dt.fromisoformat(item_copy["DateTime"].replace("Z", ""))
                    if day.date() == item_date.date():
                        #check if it is the same day
                        totalWeight += item_copy["Weight"]
                        totalLength += item_copy["Length"]
                        totalQuantity += item_copy["Quantity"]
                        #making sum of the required fields
                data_total = {"totalWeight":totalWeight,"totalLength":totalLength,"totalQuantity":totalQuantity}
                return jsonify(data_total)
            else:
                return jsonify({"Error":"No response from machstatz API"})
        except:
            return jsonify({"Error":"Please provide a valide date"})
    else:
        return jsonify({"Error":"Please specify a day"})

@app.route('/excelreport',methods=['GET'])
def excel():
    try:
        workbook = xlsxwriter.Workbook('report.xlsx')
        #create an excel file
        prev_date = None
        #for checking whether the next data in the array is same data as the previous one
        row =0
        col = 0
        #initalize to the start of the file
        for item in data:
            item_copy = item.copy()
            #make a copy of the item to not spoil the original data
            item_date = dt.fromisoformat(item_copy["DateTime"].replace("Z", ""))
            current_date = item_date.date()
            if current_date != prev_date:
                #if not same date, make a new sheet
                worksheet = workbook.add_worksheet(str(current_date))
                row = 0
                col = 0
                #initalize to the start of the file
                worksheet.write(row, col, 'DateTime') 
                worksheet.write(row, col + 1, 'Length') 
                worksheet.write(row, col + 2, 'Weight') 
                worksheet.write(row, col + 3, 'Quantity')
                row+=1
                col=0
                #next row
                worksheet.write(row, col, item_copy['DateTime']) 
                worksheet.write(row, col + 1, item_copy['Length']) 
                worksheet.write(row, col + 2, item_copy['Weight']) 
                worksheet.write(row, col + 3, item_copy['Quantity'])
                row+=1
                col=0
                #next row
                prev_date = current_date 
                #make current data as previous data
            else:
                worksheet.write(row, col, item_copy['DateTime']) 
                worksheet.write(row, col + 1, item_copy['Length']) 
                worksheet.write(row, col + 2, item_copy['Weight']) 
                worksheet.write(row, col + 3, item_copy['Quantity']) 
                row+=1
                col=0
                #next row
                prev_date = current_date  
                #make current data as previous data
        workbook.close()
        return send_from_directory(os.getcwd(),'report.xlsx', as_attachment=True) 
    except:
        return jsonify({"Error":"File cannot be generated at the moment"})

if __name__ == "__main__":
    app.config.from_pyfile('config.py')
    app.run()