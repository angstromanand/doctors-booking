from flask import Flask, render_template, request, redirect
import openpyxl
from openpyxl import Workbook
import os

app = Flask(__name__)

# Check if the Excel file exists, if not create one
if not os.path.exists('appointments.xlsx'):
    wb = Workbook()
    ws = wb.active
    ws.append(["Name", "Email", "Phone", "Reason", "Preferred Date"])
    wb.save('appointments.xlsx')

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # Get form data
        name = request.form['name']
        email = request.form['email']
        phone = request.form['phone']
        reason = request.form['reason']
        preferred_date = request.form['preferred_date']

        # Save data to Excel
        wb = openpyxl.load_workbook('appointments.xlsx')
        ws = wb.active
        ws.append([name, email, phone, reason, preferred_date])
        wb.save('appointments.xlsx')

        return redirect('/')

    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)
