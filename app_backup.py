from flask import Flask, render_template, request
import pandas as pd
from datetime import datetime
import os

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'

# Crear carpeta uploads si no existe
if not os.path.exists(app.config['UPLOAD_FOLDER']):
    os.makedirs(app.config['UPLOAD_FOLDER'])

# Cargar Excel y normalizar columnas a minúsculas
df_base = pd.read_excel("datos_base.xlsx")
df_base.columns = df_base.columns.str.lower()  # <- importante

# Crear diccionario para dropdown (Número Económico -> lista de Números de Serie)
dropdown_data = df_base.groupby('numeroeconomico')['numeroserie'].apply(list).to_dict()

@app.route('/', methods=['GET', 'POST'])
def formulario():
    if request.method == 'POST':
        numero_economico = request.form['numero_economico']
        numero_serie = request.form['numero_serie']
        reporte = request.form['reporte']
        fecha = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # Manejar archivo subido
        evidencia_file = request.files.get('evidencia')
        evidencia_filename = ""
        if evidencia_file and evidencia_file.filename != "":
            evidencia_filename = datetime.now().strftime("%Y%m%d%H%M%S_") + evidencia_file.filename
            evidencia_file.save(os.path.join(app.config['UPLOAD_FOLDER'], evidencia_filename))

        # Guardar en Excel
        datos_excel = "datos.xlsx"
        if os.path.exists(datos_excel):
            df = pd.read_excel(datos_excel)
        else:
            df = pd.DataFrame(columns=['numeroeconomico', 'numeroserie', 'reporte', 'fecha', 'evidencia'])

        # Agregar nueva fila
        df = pd.concat([df, pd.DataFrame([{
            'numeroeconomico': numero_economico,
            'numeroserie': numero_serie,
            'reporte': reporte,
            'fecha': fecha,
            'evidencia': evidencia_filename
        }])], ignore_index=True)

        df.to_excel(datos_excel, index=False)

        return f"¡Formulario enviado! {numero_economico} - {numero_serie}"

    return render_template('form.html', dropdown_data=dropdown_data)

if __name__ == '__main__':
    app.run(host='0.0.0.0', debug=True)

