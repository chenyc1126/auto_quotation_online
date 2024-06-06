from flask import Flask, request, render_template, send_from_directory, jsonify
import io
import json
import pandas as pd
import shutil
import os
import docx  # Ensure python-docx is installed
import auto_quotation

app = Flask(__name__)

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        input_data = request.form['quotationData']
        quotation = []
        lines = ""
        try:
            for line in input_data.splitlines():
                if '分隔線' not in line:
                    lines += line + "\n"
                else:
                    buf = io.StringIO(lines)
                    lines = ""
                    key = []
                    value = []
                    
                    for line in buf.readlines():
                        line = line.strip()
                        if '：' in line:
                            index = line.index('：') + 1
                            key.append(line[:index-1])
                            value.append(line[index:])
                    
                    if key and value:  # Ensure key and value are not empty
                        quotation.append(dict(zip(key, value)))
                    buf.close()
            
            # Assuming process_quotation function returns the name of the saved file
            file_name = auto_quotation.process_quotation(quotation)
            
            with open(f"output/quotation.json", "w", encoding='utf-8') as f:
                f.write(json.dumps(quotation, indent=4, ensure_ascii=False))
            
            return render_template("index.html", file_name=file_name)
        except Exception as e:
            return jsonify({"error": str(e)}), 500

    return render_template("index.html")

@app.route('/download/<filename>')
def download_file(filename):
    directory = os.path.join(app.root_path, 'output')
    return send_from_directory(directory, filename, as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)
