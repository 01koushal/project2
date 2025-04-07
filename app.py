from flask import Flask, render_template, request, redirect, url_for, flash, send_file
import fitz  # PyMuPDF
import cv2
import numpy as np
import pytesseract
import json
import os
import re
from pyzbar.pyzbar import decode
from datetime import datetime
from werkzeug.utils import secure_filename
import pandas as pd

# Optional: set tesseract path for local (on Render it's unnecessary)
# pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

app = Flask(__name__)
app.secret_key = "supersecretkey"

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER

students = []

def extract_qr_from_pdf(pdf_path):
    doc = fitz.open(pdf_path)
    pix = doc[0].get_pixmap()
    img = np.frombuffer(pix.samples, dtype=np.uint8).reshape((pix.h, pix.w, pix.n))
    if img.shape[-1] == 4:
        cv_img = cv2.cvtColor(img, cv2.COLOR_RGBA2RGB)
    else:
        cv_img = cv2.cvtColor(img, cv2.COLOR_RGB2BGR)
    return extract_qr_from_image_array(cv_img)

def extract_qr_from_image_array(image):
    image = cv2.resize(image, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC)
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    gray = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 11, 2)
    qr_codes = decode(gray)
    for qr in qr_codes:
        return qr.data.decode('utf-8')
    return None

def extract_text_from_certificate(file_path):
    text = ""
    if file_path.lower().endswith(".pdf"):
        doc = fitz.open(file_path)
        for page in doc:
            text += page.get_text("text")
    return text.lower()

def normalize_date(date_text):
    try:
        return datetime.strptime(date_text, "%B %d, %Y").strftime("%Y-%m-%d")
    except ValueError:
        return date_text

@app.route("/")
def home():
    return render_template("upload.html")

@app.route("/verify", methods=["POST"])
def verify():
    name = request.form.get("name").strip().lower()
    email = request.form.get("email")
    file = request.files.get("certificate")

    if not file or not name or not email:
        flash("All fields are required.")
        return redirect(url_for("home"))

    filename = secure_filename(file.filename)
    file_path = os.path.join(app.config["UPLOAD_FOLDER"], filename)
    file.save(file_path)

    qr_data = extract_qr_from_pdf(file_path)
    status = "Fake"
    course = "Unknown"
    date_completed = "N/A"

    if qr_data:
        try:
            qr_json = json.loads(qr_data)
            issued_to = qr_json["credentialSubject"]["issuedTo"].strip().lower()
            course = qr_json["credentialSubject"]["course"].strip().lower()
            date_completed = qr_json["credentialSubject"]["completedOn"][:10]
        except Exception as e:
            issued_to = ""

        extracted_text = extract_text_from_certificate(file_path)
        match = re.search(r"on (\w+ \d{1,2}, \d{4})", extracted_text)
        ocr_date = normalize_date(match.group(1)) if match else None

        if (
            issued_to in extracted_text and
            course in extracted_text and
            (ocr_date == date_completed)
        ):
            status = "Real"

    students.append({
        "Name": name.title(),
        "Email": email,
        "Platform": course.title(),
        "Status": status,
        "Date": date_completed
    })

    return redirect(url_for("show_verified"))

@app.route("/verified")
def show_verified():
    return render_template("verified.html", students=students)

@app.route("/download-excel")
def download_excel():
    if not students:
        flash("No student data available to download.")
        return redirect(url_for("show_verified"))

    try:
        df = pd.DataFrame(students)
        excel_path = os.path.join("/tmp", "verified_students.xlsx")
        df.to_excel(excel_path, index=False, engine='openpyxl')
        return send_file(excel_path, as_attachment=True)
    except Exception as e:
        print("Excel generation failed:", e)
        flash("Something went wrong while generating the Excel file.")
        return redirect(url_for("show_verified"))
