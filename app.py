from flask import Flask, request, render_template, send_file
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
    download_ready = False
    if request.method == "POST":
        # Get form data
        name = request.form["name"]
        age = request.form["age"]
        mobile = request.form["mobile"]
        email = request.form["email"]
        kyc = request.form["kyc"]

        # Save data to Excel
        workbook = openpyxl.load_workbook(FILE_PATH)
        sheet = workbook.active
        sheet.append([name, age, mobile, email, kyc])
        workbook.save(FILE_PATH)

        # Set flag for download link
        download_ready = True

    return render_template("index.html", download_ready=download_ready)

@app.route("/download")
def download():
    # Serve the Excel file for download
    return send_file(FILE_PATH, as_attachment=True)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
