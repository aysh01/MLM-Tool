from flask import Flask, request, render_template
import openpyxl
from openpyxl import Workbook
import os

app = Flask(__name__)
FILE_PATH = "customer_details.xlsx"

# Initialize Excel
def initialize_excel():
    if not os.path.exists(FILE_PATH):
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Customer Details"
        sheet.append(["Name", "Age", "Mobile Number", "Email", "Bank KYC"])
        workbook.save(FILE_PATH)

initialize_excel()

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        name = request.form["name"]
        age = request.form["age"]
        mobile = request.form["mobile"]
        email = request.form["email"]
        kyc = request.form["kyc"]

        workbook = openpyxl.load_workbook(FILE_PATH)
        sheet = workbook.active
        sheet.append([name, age, mobile, email, kyc])
        workbook.save(FILE_PATH)

        return "Customer details added successfully!"
    return render_template("index.html")

if __name__ == "__main__":
    app.run(debug=True)
