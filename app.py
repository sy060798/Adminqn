from flask import Flask, request, send_file
import os
from processor import process_boQ

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

@app.route("/process", methods=["POST"])
def process():

    boq = request.files["boq"]
    lms_files = request.files.getlist("lms")

    boq_path = os.path.join(UPLOAD_FOLDER, boq.filename)
    boq.save(boq_path)

    lms_paths = []
    for f in lms_files:
        path = os.path.join(UPLOAD_FOLDER, f.filename)
        f.save(path)
        lms_paths.append(path)

    output_path = os.path.join(UPLOAD_FOLDER, "hasil.xlsx")

    process_boQ(boq_path, lms_paths, output_path)

    return send_file(output_path, as_attachment=True)


if __name__ == "__main__":
    app.run(debug=True)
