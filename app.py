from flask import Flask,jsonify,request,send_from_directory
import json
import requests
from datetime import datetime
import xlsxwriter
import os

data = json.loads(requests.get("https://assignment-machstatz.herokuapp.com/excel").text) 

app = Flask(__name__) 

@app.route('/total',methods=['GET'])
def res():
    data_filtered = []
    if 'day' in request.args:
        # try:
            day = request.args['day']
            day = datetime.strptime(day, '%d-%m-%Y')
            print(day.date())
            if len(data) > 0:
                for item in data:
                    item_copy = item.copy()
                    item_date = datetime.fromisoformat(item_copy["DateTime"].replace("Z", ""))
                    if day.date() == item_date.date():
                        item_copy.pop("DateTime")
                        data_filtered.append(item_copy)
                return jsonify(data_filtered)
            else:
                return "No response from machstatz API"
        # except:
        #     return "Error: Please provide a valide date"
    else:
        return "Error: Please specify a day"

@app.route('/excelreport',methods=['GET'])
def excel():
    try:
        workbook = xlsxwriter.Workbook('report.xlsx')
        first_date = None
        sheet =1
        row =0
        col = 0
        for item in data:
            item_copy = item.copy()
            item_date = datetime.fromisoformat(item_copy["DateTime"].replace("Z", ""))
            current_date = item_date.date()
            if current_date != first_date:
                worksheet = workbook.add_worksheet(str(current_date))
                row = 0
                col = 0
                worksheet.write(row, col, 'DateTime') 
                worksheet.write(row, col + 1, 'Length') 
                worksheet.write(row, col + 2, 'Weight') 
                worksheet.write(row, col + 3, 'Quantity')
                row+=1
                col=0
                worksheet.write(row, col, item_copy['DateTime']) 
                worksheet.write(row, col + 1, item_copy['Length']) 
                worksheet.write(row, col + 2, item_copy['Weight']) 
                worksheet.write(row, col + 3, item_copy['Quantity'])
                row+=1
                col=0
                sheet+=1
                first_date = current_date 
            else:
                worksheet.write(row, col, item_copy['DateTime']) 
                worksheet.write(row, col + 1, item_copy['Length']) 
                worksheet.write(row, col + 2, item_copy['Weight']) 
                worksheet.write(row, col + 3, item_copy['Quantity']) 
                row+=1
                col=0
                first_date = current_date  
        workbook.close()
        return send_from_directory(os.getcwd(),'report.xlsx', as_attachment=True) 
    except:
        return "Error: File cannot be generated at the moment"  

if __name__ == "__main__":
    app.run(threaded=True, port=5000,debug=True)