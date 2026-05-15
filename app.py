from flask import Flask, render_template, request, send_file

from generate_qbr import create_qbr

app = Flask(__name__)

@app.route("/")
def home():
    return render_template("index.html")

@app.route("/generate")
def generate():

    customer = request.args.get("customer")

    filename = create_qbr(customer)

    return send_file(
        filename,
        as_attachment=True
    )

if __name__ == "__main__":
    app.run(debug=True)
    
