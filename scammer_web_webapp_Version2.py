import os
import io
import secrets
from flask import Flask, render_template, request, redirect, url_for, session, send_file, flash, abort
from werkzeug.utils import secure_filename
from functools import wraps
import csv
import json
import xlsxwriter
from rapidfuzz import process, fuzz

ALLOWED_EXTENSIONS = {'csv', 'json'}
EXPORT_FORMATS = {'csv', 'xlsx', 'html'}

app = Flask(__name__)
app.secret_key = secrets.token_hex(32)
USERS = {"admin": "changeme123"}

def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if "user" not in session:
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return decorated

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

class Contact:
    def __init__(self, name, email, phone):
        self.name = name.strip().lower()
        self.email = email.strip().lower()
        self.phone = self._normalize_phone(phone)
    def _normalize_phone(self, phone):
        import re
        digits = re.sub(r'\D', '', phone)
        if len(digits) == 10:
            return "+1" + digits
        elif len(digits) == 11 and digits.startswith('1'):
            return "+" + digits
        elif digits.startswith('00'):
            return "+" + digits[2:]
        elif digits.startswith('+'):
            return digits
        return digits

def load_contacts_from_csv(fileobj):
    contacts = {}
    fileobj.seek(0)
    reader = csv.DictReader(io.StringIO(fileobj.read().decode('utf-8')))
    for row in reader:
        name = row.get('name', '') or row.get('Name', '')
        email = row.get('email', '') or row.get('Email', '')
        phone = row.get('phone', '') or row.get('Phone', '')
        contact = Contact(name, email, phone)
        if contact.name:
            contacts[contact.name] = contact
    return contacts

def load_contacts_from_json(fileobj):
    contacts = {}
    fileobj.seek(0)
    data = json.load(io.TextIOWrapper(fileobj, 'utf-8'))
    for entry in data:
        name = entry.get('name', '')
        email = entry.get('email', '')
        phone = entry.get('phone', '')
        contact = Contact(name, email, phone)
        if contact.name:
            contacts[contact.name] = contact
    return contacts

def find_suspect_contacts(official_contacts, suspect_contacts, fuzzy=False, fuzzy_threshold=90):
    flagged = []
    for name, suspect in suspect_contacts.items():
        official = official_contacts.get(name)
        if not official and fuzzy:
            matches = process.extract(name, list(official_contacts.keys()), scorer=fuzz.ratio, limit=1)
            if matches and matches[0][1] >= fuzzy_threshold:
                official = official_contacts[matches[0][0]]
        reason = []
        if not official:
            flagged.append({
                "name": suspect.name,
                "suspect_email": suspect.email,
                "suspect_phone": suspect.phone,
                "official_email": "",
                "official_phone": "",
                "reason": "Name not found in official records"
            })
            continue
        email_match = suspect.email == official.email
        phone_match = suspect.phone == official.phone
        if not email_match:
            reason.append("Email mismatch")
        if not phone_match:
            reason.append("Phone mismatch")
        if reason:
            flagged.append({
                "name": suspect.name,
                "suspect_email": suspect.email,
                "suspect_phone": suspect.phone,
                "official_email": official.email,
                "official_phone": official.phone,
                "reason": "; ".join(reason)
            })
    return flagged

def save_report_excel(report, fileobj):
    workbook = xlsxwriter.Workbook(fileobj, {'in_memory': True})
    worksheet = workbook.add_worksheet()
    headers = list(report[0].keys())
    for col, h in enumerate(headers):
        worksheet.write(0, col, h)
    for row, entry in enumerate(report, start=1):
        for col, h in enumerate(headers):
            worksheet.write(row, col, entry[h])
    workbook.close()

def save_report_html(report, fileobj):
    headers = list(report[0].keys())
    html = "<html><body><table border='1'><tr>"
    for h in headers:
        html += f"<th>{h}</th>"
    html += "</tr>"
    for entry in report:
        html += "<tr>"
        for h in headers:
            html += f"<td>{entry[h]}</td>"
        html += "</tr>"
    html += "</table></body></html>"
    fileobj.write(html)

@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        user = request.form.get("username", "")
        pw = request.form.get("password", "")
        if USERS.get(user) == pw:
            session["user"] = user
            return redirect(url_for("index"))
        flash("Invalid credentials")
    return render_template("login.html")

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))

@app.route("/", methods=["GET", "POST"])
@login_required
def index():
    if request.method == "POST":
        official_file = request.files.get("official")
        suspect_file = request.files.get("suspect")
        if not official_file or not allowed_file(official_file.filename):
            flash("Invalid or missing official contacts file.")
            return redirect(request.url)
        if not suspect_file or not allowed_file(suspect_file.filename):
            flash("Invalid or missing suspect contacts file.")
            return redirect(request.url)
        official_ext = official_file.filename.rsplit('.', 1)[1].lower()
        suspect_ext = suspect_file.filename.rsplit('.', 1)[1].lower()
        if official_ext == "json":
            official_contacts = load_contacts_from_json(official_file)
        else:
            official_contacts = load_contacts_from_csv(official_file)
        if suspect_ext == "json":
            suspect_contacts = load_contacts_from_json(suspect_file)
        else:
            suspect_contacts = load_contacts_from_csv(suspect_file)
        fuzzy = bool(request.form.get("fuzzy"))
        fuzzy_threshold = int(request.form.get("fuzzy_threshold") or 90)
        flagged = find_suspect_contacts(official_contacts, suspect_contacts, fuzzy=fuzzy, fuzzy_threshold=fuzzy_threshold)
        session["results"] = flagged
        return render_template("results.html", flagged=flagged)
    return render_template("index.html")

@app.route("/export/<fmt>")
@login_required
def export(fmt):
    if fmt not in EXPORT_FORMATS:
        abort(400)
    flagged = session.get("results")
    if not flagged:
        abort(404)
    buf = io.BytesIO()
    if fmt == "csv":
        output = io.StringIO()
        writer = csv.DictWriter(output, fieldnames=flagged[0].keys())
        writer.writeheader()
        for row in flagged:
            writer.writerow(row)
        buf.write(output.getvalue().encode("utf-8"))
        buf.seek(0)
        return send_file(buf, mimetype="text/csv", as_attachment=True, download_name="report.csv")
    elif fmt == "xlsx":
        save_report_excel(flagged, buf)
        buf.seek(0)
        return send_file(buf, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                         as_attachment=True, download_name="report.xlsx")
    elif fmt == "html":
        output = io.StringIO()
        save_report_html(flagged, output)
        buf.write(output.getvalue().encode("utf-8"))
        buf.seek(0)
        return send_file(buf, mimetype="text/html", as_attachment=True, download_name="report.html")
    else:
        abort(400)

if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=False)