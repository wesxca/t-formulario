from flask import Flask, render_template, request
import pandas as pd
from datetime import datetime
import openpyxl

app = Flask(__name__)

# Leer base de datos de números
df_base = pd.read_excel("datos_base.xlsx")
df_base.columns = [col.upper() for col in df_base.columns]  # Ajustar nombres
dropdown_data = df_base.groupby('NUMEROECONOMICO')['NUMEROSERIE'].apply(list).to_dict()

@app.route("/", methods=["GET", "POST"])
def form():
    if request.method == "POST":
        numero_economico = request.form.get("numero_economico")
        numero_serie = request.form.get("numero_serie")
        reporte = request.form.get("reporte")
        fecha = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # Guardar en Excel
        try:
            libro = openpyxl.load_workbook("respuestas.xlsx")
        except FileNotFoundError:
            libro = openpyxl.Workbook()
            hoja = libro.active
            hoja.append(["NumeroEconomico", "NumeroSerie", "Reporte", "Fecha"])
        hoja = libro.active
        hoja.append([numero_economico, numero_serie, reporte, fecha])
        libro.save("respuestas.xlsx")

        return "¡Formulario enviado!"
    
    return render_template("form.html", opciones=dropdown_data)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
