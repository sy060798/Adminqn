from flask import Flask, request, send_file
import os
from processor import process_boQ

app = Flask(__name__)

@app.route("/process", methods=["POST"])
def process():

    boq = request.files["boq"]
    lms = request.files.getlist("lms")

    boq_path = "boq.xlsx"
    boq.save(boq_path)

    lms_paths = []
    for i,f in enumerate(lms):
        path = f"lms_{i}.xlsx"
        f.save(path)
        lms_paths.append(path)

    output = "hasil.xlsx"

    process_boQ(boq_path, lms_paths, output)

    return send_file(output, as_attachment=True)

app.run(debug=True)
