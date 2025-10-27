from flask import Flask, render_template, request, jsonify
import pandas as pd
from datetime import datetime
import os

app = Flask(__name__)
EXCEL_FILE = 'events.xlsx'

# 初始化 Excel 檔案
if not os.path.exists(EXCEL_FILE):
    df = pd.DataFrame(columns=['date', 'title', 'description'])
    df.to_excel(EXCEL_FILE, index=False)

@app.route('/')
def calendar_page():
    return render_template('calendar.html')

@app.route('/events')
def get_events():
    df = pd.read_excel(EXCEL_FILE)
    events = []
    for _, row in df.iterrows():
        events.append({
            "title": row['title'],
            "start": row['date'].strftime('%Y-%m-%d'),
            "description": row.get('description', '')
        })
    return jsonify(events)

@app.route('/add_event', methods=['POST'])
def add_event():
    title = request.form.get('title')
    date = request.form.get('date')
    description = request.form.get('description', '')

    df = pd.read_excel(EXCEL_FILE)
    df = pd.concat([df, pd.DataFrame([{
        'date': pd.to_datetime(date),
        'title': title,
        'description': description
    }])], ignore_index=True)
    df.to_excel(EXCEL_FILE, index=False)

    return jsonify({"status": "success"})

if __name__ == '__main__':
    app.run(debug=True)
