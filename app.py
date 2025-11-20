from flask import Flask, render_template, request, send_file
import pandas as pd
import matplotlib.pyplot as plt
import re, os, zipfile
from datetime import datetime

app = Flask(__name__)

UPLOAD_FOLDER = "static/output"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def extract_number(val):
    if isinstance(val, str):
        match = re.findall(r"-?\d+\.?\d*", val)
        if match:
            return float(match[0])
    return None

@app.route("/", methods=["GET", "POST"])
def upload_file():
    if request.method == "POST":
        file = request.files["excel_file"]
        if not file:
            return render_template("upload.html", download_link=None)

        # Save uploaded file temporarily
        filepath = os.path.join(UPLOAD_FOLDER, "uploaded.xlsx")
        file.save(filepath)

        # Process Excel
        df = pd.read_excel(filepath)
        data = df.iloc[4:].copy()
        data.columns = ["Group", "Code", "Value", "Start", "End"]

        # Clean numeric values
        data["Numeric"] = data["Value"].apply(extract_number)
        data["Start"] = pd.to_datetime(data["Start"], errors="coerce")

        metrics = [
            "UP Speed", "Down Speed",
            "UP Proportion", "Down Proportion",
            "Tx_power", "Rx_power"
        ]

        timestamp_folder = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_dir = os.path.join(UPLOAD_FOLDER, timestamp_folder)
        os.makedirs(output_dir)

        # Generate plots
        for metric in metrics:
            subset = data[data["Code"] == metric].sort_values("Start")

            plt.figure(figsize=(12, 5))
            plt.plot(subset["Start"], subset["Numeric"])
            plt.title(metric)
            plt.xlabel("Time")
            plt.ylabel(metric)

            ticks = subset["Start"][::10]
            plt.xticks(ticks, ticks.dt.strftime('%Y-%m-%d %H:%M'), rotation=90)

            plt.tight_layout()
            save_path = os.path.join(output_dir, metric.replace(" ", "_") + ".jpeg")
            plt.savefig(save_path, format="jpeg")
            plt.close()

        # Create ZIP file
        zip_path = os.path.join(UPLOAD_FOLDER, f"{timestamp_folder}.zip")
        with zipfile.ZipFile(zip_path, 'w') as zipf:
            for file_name in os.listdir(output_dir):
                zipf.write(os.path.join(output_dir, file_name),
                           arcname=file_name)

        return render_template("upload.html",
                               download_link=f"/download/{timestamp_folder}")

    return render_template("upload.html", download_link=None)

@app.route("/download/<folder>")
def download_zip(folder):
    zip_file = os.path.join(UPLOAD_FOLDER, f"{folder}.zip")
    return send_file(zip_file, as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)
