import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import zipfile
import io
import re
import os
from datetime import datetime

st.set_page_config(page_title="Excel Graph Generator", layout="centered")

# ------------------------
# Helper Function
# ------------------------
def extract_number(val):
    if isinstance(val, str):
        match = re.findall(r"-?\d+\.?\d*", val)
        if match:
            return float(match[0])
    return None


# ------------------------
# Streamlit UI
# ------------------------
st.title("📊 Excel Graph Generator")
st.write("Upload your Excel file and automatically generate graphs.")

uploaded_file = st.file_uploader("Choose Excel File", type=["xlsx"])

if uploaded_file:
    st.success(f"📁 Selected File: **{uploaded_file.name}**")

if uploaded_file and st.button("Generate Graphs"):
    with st.spinner("Processing file..."):
        df = pd.read_excel(uploaded_file)

        # Process dataframe similar to Flask
        data = df.iloc[4:].copy()
        data.columns = ["Group", "Code", "Value", "Start", "End"]

        data["Numeric"] = data["Value"].apply(extract_number)
        data["Start"] = pd.to_datetime(data["Start"], errors="coerce")

        metrics = [
            "UP Speed", "Down Speed",
            "UP Proportion", "Down Proportion",
            "Tx_power", "Rx_power"
        ]

        # Temporary folder inside memory
        zip_buffer = io.BytesIO()

        # Start ZIP creation
        with zipfile.ZipFile(zip_buffer, "w") as zipf:
            progress = st.progress(0)

            for i, metric in enumerate(metrics):
                subset = data[data["Code"] == metric].sort_values("Start")

                # Create matplotlib plot
                plt.figure(figsize=(12, 5))
                plt.plot(subset["Start"], subset["Numeric"])
                plt.title(metric)
                plt.xlabel("Time")
                plt.ylabel(metric)

                ticks = subset["Start"][::10]
                plt.xticks(ticks, ticks.dt.strftime('%Y-%m-%d %H:%M'), rotation=90)
                plt.tight_layout()

                # Save image to memory
                img_bytes = io.BytesIO()
                plt.savefig(img_bytes, format="jpeg")
                img_bytes.seek(0)
                plt.close()

                # Add to ZIP
                zipf.writestr(metric.replace(" ", "_") + ".jpeg", img_bytes.read())

                # Update progress bar
                progress.progress((i + 1) / len(metrics))

    st.success("✔ Graphs generated successfully!")

    # Download ZIP button
    st.download_button(
        label="⬇ Download ZIP File",
        data=zip_buffer.getvalue(),
        file_name=f"graphs_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
        mime="application/zip"
    )
