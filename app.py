import json
import os
import pickle
from flask import Flask, render_template, request, redirect
from datetime import datetime

# قمنا بتعطيل مكتبات الإكسل لأنها تسبب تعارض مع سيرفر Render في النسخة المجانية
# from openpyxl import Workbook, load_workbook
# from openpyxl.chart import PieChart, Reference
# from openpyxl.chart.series import DataPoint

app = Flask(__name__)

history_data = []
EXCEL_FILE = "history.xlsx"

# تحميل الموديل والـ vectorizer
with open("model.pkl", "rb") as f:
    model = pickle.load(f)

with open("vectorizer.pkl", "rb") as f:
    vectorizer = pickle.load(f)

# تعطيل دالة حفظ الإكسل بالكامل لتجنب أخطاء الصلاحيات (Permission Errors)
def save_to_excel(message, prediction, time):
    pass 
    # تم تعطيل الكود الداخلي للدالة لضمان استقرار الرفع
    """
    if os.path.exists(EXCEL_FILE):
        wb = load_workbook(EXCEL_FILE)
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "History"
        ws.append(["Message", "Prediction", "Time"])
    ws.append([message, prediction, time])
    wb.save(EXCEL_FILE)
    """

@app.route('/', methods=['GET', 'POST'])
def check():
    prediction = None
    message = request.form.get('message', '') or request.args.get('message', '')
    action = request.form.get('action', '')

    if action == "model_info":
        return redirect("/model")
    if action == "sample_message":
        return redirect("/samples")

    if request.method == 'POST' and action == "check":
        if message.strip() == "":
            prediction = "Neutral 🔵"
            history_data.append({
                "message": "(Empty Message)",
                "prediction": prediction,
                "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            })
            # save_to_excel("(Empty Message)", prediction, datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        else:
            X = vectorizer.transform([message])
            y_pred = model.predict(X)[0]

            if y_pred == 1:
                prediction = "Spam ❌"
            else:
                prediction = "Not Spam ✅"

            history_data.append({
                "message": message,
                "prediction": prediction,
                "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            })
            # تم تعطيل الاستدعاء هنا أيضاً
            # save_to_excel(message, prediction, datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

    return render_template('check.html', prediction=prediction, message=message)

@app.route('/history')
def history():
    spam_count = sum(1 for item in history_data if item["prediction"] == "Spam ❌")
    notspam_count = sum(1 for item in history_data if item["prediction"] == "Not Spam ✅")
    neutral_count = sum(1 for item in history_data if item["prediction"] == "Neutral 🔵")

    return render_template(
        'history.html',
        history=history_data,
        spam_count=spam_count,
        notspam_count=notspam_count,
        neutral_count=neutral_count
    )

@app.route('/reset')
def reset_history():
    global history_data
    history_data = []
    # تم تعطيل حذف ملف الإكسل لتجنب المشاكل
    return redirect('/history')

@app.route('/model')
def model_info():
    model_details = {
        "name": "Naive Bayes Classifier",
        "description": "This model uses TF-IDF Vectorizer to transform text data and Naive Bayes for classification.",
        "implementation": "Project implementation by: Mohammad Hayajneh and Mothana Mufleh",
        "supervised": "Supervised by: Dr. Ra’ed Alkhateeb"
    }
    return render_template('model.html', model=model_details)

@app.route('/samples')
def samples():
    try:
        with open("messages.json", "r", encoding="utf-8") as f:
            data = json.load(f)
        return render_template('samples.html', spam=data["spam"], not_spam=data["not_spam"])
    except FileNotFoundError:
        return "messages.json file not found."

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)