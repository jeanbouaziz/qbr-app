from flask import request, send_file
import os

@app.route("/generate_qbr", methods=["POST"])
def generate_qbr():

    client = request.form.get("client")
    excel_file = request.files["excel"]
    ppt_file = request.files["ppt"]

    # Save uploaded files temporarily
    excel_path = os.path.join("temp_excel.xlsx")
    ppt_path = os.path.join("temp_template.pptx")

    excel_file.save(excel_path)
    ppt_file.save(ppt_path)

    print("Client:", client)
    print("Excel received:", excel_path)
    print("PPT received:", ppt_path)

    return "Files received successfully"


