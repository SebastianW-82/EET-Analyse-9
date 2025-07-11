import streamlit as st
import matplotlib.pyplot as plt
import pandas as pd
import numpy as np
import os
import io
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from PIL import Image

st.set_page_config(layout="wide")
st.title("Analyse Tool EET-Daten")

# Excel-Datei laden (muss im Projektverzeichnis liegen)
excel_path = "EET_Beispieldaten_100_ISINs_variiert.xlsx"

if not os.path.exists(excel_path):
    st.error(f"Die Datei '{excel_path}' wurde nicht gefunden. Bitte stelle sicher, dass sie im Projektverzeichnis liegt.")
else:
    data = pd.read_excel(excel_path)

    spalten = [
        'Mindestanteil nachhaltiger Investionen (in %)',
        'Tatsächlicher Anteil nachhaltiger Investitionen (in %)',
        'Mindestanteil taxonomiekonformer Investitionen (in %)',
        'Tatsächlicher Anteil taxonomiekonformer Investitionen (in %)',
        'Scope 1 Emissionen (in MT)',
        'Scope 2 Emissionen (in MT)',
        'Scope 3 Emissionen (in MT)'
    ]

    st.markdown("""
        <style>
        div[data-testid=\"stTextInput\"] input {
            width: 300px;
        }
        </style>
    """, unsafe_allow_html=True)

    user_isin = st.text_input("Geben Sie bitte eine ISIN ein:").strip().upper()

    if st.button("Analyse starten"):
        if user_isin not in data['ISIN'].values:
            st.error("ISIN nicht gefunden. Bitte überprüfen Sie Ihre Eingabe.")
        else:
            user_row = data[data['ISIN'] == user_isin].iloc[0]
            user_klassifikation = user_row['Klassifikation']
            subset = data[data['Klassifikation'] == user_klassifikation]

            klassifikation_label = f"Art. {int(user_klassifikation)}"

            st.markdown("<div style='height: 30px'></div>", unsafe_allow_html=True)

            st.markdown(f"""
                <div style='background-color:#00a0de; padding: 15px; border-radius: 0px; width: fit-content;'>
                    <h4 style='color: white;'>Daten zur ISIN {user_isin}</h4>
                    <p style='color: white;'>Klassifikation: {klassifikation_label}</p>
                    <ul style='color: white;'>
                        {''.join([f'<li>{column}: {user_row[column]}</li>' for column in spalten])}
                    </ul>
                </div>
            """, unsafe_allow_html=True)

            st.markdown("<div style='height: 80px'></div>", unsafe_allow_html=True)

            pdf_buffer = io.BytesIO()
            c = canvas.Canvas(pdf_buffer, pagesize=A4)

            for idx, column in enumerate(spalten):
                user_value = user_row[column]
                mean_val = subset[column].mean()
                median_val = subset[column].median()
                percentile = (subset[column] < user_value).mean() * 100
                num_values = subset[column].count()

                col1, col2 = st.columns([3, 1], gap='large')

                with col1:
                    st.markdown(f"<h4 style='margin-bottom: 5px;'>{column}</h4>", unsafe_allow_html=True)
                    fig, ax = plt.subplots(figsize=(8, 3.5))
                    plt.xticks(fontsize=7)
                    plt.yticks(fontsize=7)
                    ax.hist(subset[column], bins=10, edgecolor='black', alpha=0.3, label='Verteilung')

                    ax.axvline(user_value, color='red', linestyle='--', linewidth=2, label='Wert zur ISIN')
                    ax.axvline(mean_val, color='green', linestyle=':', linewidth=2, label='Mittelwert')
                    ax.axvline(median_val, color='blue', linestyle='-.', linewidth=2, label='Median')

                    ax.set_xlabel(column, fontsize=8)
                    ax.set_ylabel('Häufigkeit', fontsize=8)
                    ax.legend(prop={'size': 6})
                    ax.grid(True)

                    st.pyplot(fig)

                    buf = io.BytesIO()
                    fig.savefig(buf, format='png', bbox_inches='tight')
                    buf.seek(0)
                    image = Image.open(buf)
                    image_path = f"{column}.png"
                    image.save(image_path)
                    c.drawString(2 * cm, 27 * cm, f"Analyse zur ISIN: {user_isin}")
                    c.drawString(2 * cm, 26.5 * cm, f"{column}:")
                    c.drawImage(image_path, 2 * cm, 15 * cm, width=15.9 * cm, height=7.9 * cm)

                    # Grauer Kasten rechts neben dem Diagramm (PDF)
                    info_x = 15.9 * cm
                    info_y = 23.2 * cm
                    c.setFillColorRGB(0.97, 0.97, 0.97)
                    c.setFont("Helvetica", 8)
                    text_lines = [
                        f"Wert zur ISIN: {user_value}",
                        f"Anzahl ISINs Peergroup: {num_values}",
                        f"Mittelwert: {mean_val:.2f}",
                        f"Median: {median_val:.2f}",
                        f"{percentile:.1f}% der Werte sind kleiner"
                    ]
                    box_height = len(text_lines) * 0.55 * cm + 0.5 * cm
                    c.rect(info_x, info_y, 5.5 * cm, box_height, fill=1, stroke=0)
                    c.setFillColorRGB(0, 0, 0)
                    line_y = info_y + box_height - 0.5 * cm
                    for line in text_lines:
                        c.drawString(info_x + 0.3 * cm, line_y, line)
                        line_y -= 0.55 * cm

                    c.showPage()

                with col2:
                    st.markdown(f"""
                    <div style='background-color: #f2f2f2; padding: 10px; border-radius: 5px;'>
                        <strong>ISIN:</strong> {user_isin}<br>
                        <strong>{column}:</strong> {user_value}<br>
                        <strong>Anzahl ISINs Peergroup:</strong> {num_values}<br>
                        <strong>Mittelwert:</strong> {mean_val:.1f}<br>
                        <strong>Median:</strong> {median_val:.1f}<br>
                        <strong>{percentile:.1f}%</strong> der Werte sind kleiner als der ISIN-Wert
                    </div>
                    """, unsafe_allow_html=True)

            c.save()
            pdf_buffer.seek(0)
            st.download_button(
                label="📄 Gesamte Analyse als PDF herunterladen",
                data=pdf_buffer,
                file_name=f"{user_isin}_analyse.pdf",
                mime="application/pdf"
            )
