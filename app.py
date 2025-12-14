from flask import Flask, render_template, request, redirect, session, jsonify, send_file
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side
from datetime import datetime, date
import os
import smtplib
from email.message import EmailMessage
import json
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "super_secret_key_123")

# =======================================
# ARCHIVO DINÁMICO POR USUARIO + RUTA
# =======================================
def archivo_excel():
    usuario = session.get("usuario", "user").replace(" ", "_")
    ruta = session.get("ruta", "ruta").replace(" ", "_")
    return f"registro_{usuario}_{ruta}.xlsx"

# =======================================
# CATEGORÍAS BASE
# =======================================
categorias_base = [
    "Autos", "Camionetas", "Micro Bus", "Mini Bus", "Bus", "Omnibus",
    "Camiones de 1 Eje", "Camiones de 2 Ejes", "Camiones 3 Ejes o más",
    "Motos", "Jeeps", "Bicicletas", "Peatones", "V.A", "V.C",
    "Tracción Animal", "Rickshaw"
]

# =======================================
# ESTILOS EXCEL
# =======================================
def estilizar(ws):
    header_font = Font(bold=True)
    header_fill = PatternFill("solid", fgColor="D9D9D9")
    border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin")
    )

    for c in ws[1]:
        c.font = header_font
        c.fill = header_fill
        c.border = border

    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.border = border

# =======================================
# CREAR EXCEL
# =======================================
def inicializar_excel():
    archivo = archivo_excel()

    if not os.path.exists(archivo):
        wb = openpyxl.Workbook()

        c = wb.active
        c.title = "Conteo"
        c.append(["Categoría", "Conteo", "Fecha", "Ruta"])

        for cat in categorias_base:
            c.append([cat, 0, "", ""])

        h = wb.create_sheet("Historial")
        h.append(["Fecha", "Hora", "Ruta", "Usuario", "Categoría", "Cantidad", "N° Vehículos"])

        estilizar(c)
        estilizar(h)
        wb.save(archivo)

# =======================================
# LOGIN
# =======================================
@app.route("/", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        usuario = request.form["usuario"].strip()
        ruta = request.form["ruta"].strip()

        if not usuario or not ruta:
            return render_template("login.html", error="Debes completar Usuario y Ruta")

        session["usuario"] = usuario
        session["ruta"] = ruta
        session["conteos"] = {c: 0 for c in categorias_base}
        session["nuevas"] = {}
        session["vehiculos_hoy"] = 0

        inicializar_excel()
        return redirect("/contador")

    return render_template("login.html")

# =======================================
# CONTADOR
# =======================================
@app.route("/contador")
def contador():
    if "usuario" not in session:
        return redirect("/")

    return render_template(
        "contador.html",
        categorias_base=categorias_base,
        conteos=session["conteos"],
        nuevas=session["nuevas"]
    )

# =======================================
# MODIFICAR
# =======================================
@app.route("/modificar", methods=["POST"])
def modificar():
    cat = request.form["categoria"]
    accion = request.form["accion"]

    if cat in session["conteos"]:
        if accion == "sumar":
            session["conteos"][cat] += 1
            session["vehiculos_hoy"] += 1
        elif accion == "restar" and session["conteos"][cat] > 0:
            session["conteos"][cat] -= 1
            session["vehiculos_hoy"] -= 1
    else:
        if accion == "sumar":
            session["nuevas"][cat] += 1
            session["vehiculos_hoy"] += 1
        elif accion == "restar" and session["nuevas"][cat] > 0:
            session["nuevas"][cat] -= 1
            session["vehiculos_hoy"] -= 1

    session.modified = True
    return jsonify(ok=True)

# =======================================
# NUEVA CATEGORÍA
# =======================================
@app.route("/nueva_categoria", methods=["POST"])
def nueva_categoria():
    nombre = request.form["nombre"].strip()

    if not nombre:
        return jsonify(error="Debe escribir un nombre")

    if nombre in session["conteos"] or nombre in session["nuevas"]:
        return jsonify(error="La categoría ya existe")

    session["nuevas"][nombre] = 0
    session.modified = True
    return jsonify(ok=True)

# =======================================
# GUARDAR EXCEL
# =======================================
@app.route("/guardar", methods=["POST"])
def guardar():
    archivo = archivo_excel()
    wb = openpyxl.load_workbook(archivo)
    conteo = wb["Conteo"]
    hist = wb["Historial"]

    fecha = date.today().strftime("%d-%m-%Y")
    hora = datetime.now().strftime("%H:%M:%S")

    def actualizar(cat, cant):
        for row in conteo.iter_rows(min_row=2):
            if row[0].value == cat:
                row[1].value += cant
                row[2].value = fecha
                row[3].value = session["ruta"]
                break
        else:
            conteo.append([cat, cant, fecha, session["ruta"]])

        hist.append([
            fecha, hora, session["ruta"], session["usuario"],
            cat, cant, session["vehiculos_hoy"]
        ])

    for c, n in session["conteos"].items():
        if n > 0:
            actualizar(c, n)

    for c, n in session["nuevas"].items():
        if n > 0:
            actualizar(c, n)

    total = sum(row[1].value for row in conteo.iter_rows(min_row=2)
                if row[0].value != "N° Vehículos")

    for i, row in enumerate(conteo.iter_rows(min_row=2), start=2):
        if row[0].value == "N° Vehículos":
            conteo.delete_rows(i)
            break

    conteo.append(["N° Vehículos", total, fecha, session["ruta"]])
    wb.save(archivo)

    session["conteos"] = {c: 0 for c in categorias_base}
    session["nuevas"] = {c: 0 for c in session["nuevas"]}
    session.modified = True

    return jsonify(ok=True, mensaje="Datos guardados correctamente.")

# =======================================
# GUARDAR EN GOOGLE DRIVE
# =======================================
def subir_a_drive(archivo):
    creds_json = os.environ.get("GOOGLE_CREDENTIALS")
    folder_id = os.environ.get("DRIVE_FOLDER_ID")

    if not creds_json:
        raise Exception("GOOGLE_CREDENTIALS no definido")
    if not folder_id:
        raise Exception("DRIVE_FOLDER_ID no definido")

    creds_dict = json.loads(creds_json)

    creds = Credentials.from_service_account_info(
        creds_dict,
        scopes=["https://www.googleapis.com/auth/drive.file"]
    )

    service = build("drive", "v3", credentials=creds)

    file_metadata = {
        "name": os.path.basename(archivo),
        "parents": [folder_id]
    }

    media = MediaFileUpload(
        archivo,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    service.files().create(
        body=file_metadata,
        media_body=media,
        fields="id"
    ).execute()


def enviar_excel_por_correo(archivo):
    remitente = os.environ.get("MAIL_USER")
    contrasena = os.environ.get("MAIL_PASS")
    destinatario = os.environ.get("MAIL_TO")

    if not remitente or not contrasena or not destinatario:
        raise Exception("Variables de correo no configuradas")

    msg = EmailMessage()
    msg["From"] = remitente
    msg["To"] = destinatario
    msg["Subject"] = "📊 Registro de Conteo de Vehículos"
    msg.set_content("Se adjunta el archivo Excel generado automáticamente.")

    with open(archivo, "rb") as f:
        file_data = f.read()

    msg.add_attachment(
        file_data,
        maintype="application",
        subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=os.path.basename(archivo)
    )

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
        smtp.login(remitente, contrasena)
        smtp.send_message(msg)

# =======================================
# ABRIR / DESCARGAR EXCEL
# =======================================
@app.route("/abrir_excel")
def abrir_excel():
    if "usuario" not in session:
        return redirect("/")

    archivo = archivo_excel()

    if not os.path.exists(archivo):
        return "Archivo no encontrado", 404

    return send_file(
        archivo,
        as_attachment=True,
        download_name=os.path.basename(archivo)
    )

# =======================================
# CERRAR SESIÓN + GUARDAR EN DRIVE
# =======================================
@app.route("/cerrar", methods=["POST"])
def cerrar():
    try:
        # 1️⃣ Guardar los datos antes de enviar
        archivo = archivo_excel()
        wb = openpyxl.load_workbook(archivo)
        conteo = wb["Conteo"]
        hist = wb["Historial"]

        fecha = date.today().strftime("%d-%m-%Y")
        hora = datetime.now().strftime("%H:%M:%S")

        def actualizar(cat, cant):
            for row in conteo.iter_rows(min_row=2):
                if row[0].value == cat:
                    row[1].value += cant
                    row[2].value = fecha
                    row[3].value = session["ruta"]
                    break
            else:
                conteo.append([cat, cant, fecha, session["ruta"]])

            hist.append([
                fecha, hora, session["ruta"], session["usuario"],
                cat, cant, session["vehiculos_hoy"]
            ])

        for c, n in session["conteos"].items():
            if n > 0:
                actualizar(c, n)

        for c, n in session["nuevas"].items():
            if n > 0:
                actualizar(c, n)

        wb.save(archivo)

        # 2️⃣ Enviar correo
        enviar_excel_por_correo(archivo)

        # 3️⃣ Limpiar sesión
        session.clear()

        return jsonify(
            ok=True,
            mensaje="📧 Excel guardado y enviado automáticamente al correo."
        )

    except Exception as e:
        return jsonify(
            ok=False,
            mensaje=f"❌ Error enviando correo: {str(e)}"
        )

# =======================================
# RUN
# =======================================
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))


