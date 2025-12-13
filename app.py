from flask import Flask, render_template, request, redirect, session, url_for, jsonify
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from datetime import datetime, date
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders

app = Flask(__name__)
app.secret_key = "super_secret_key_123"

# =======================================
# ARCHIVO DINÁMICO POR USUARIO + RUTA
# =======================================
def archivo_excel():
    usuario = session.get("usuario", "user")
    ruta = session.get("ruta", "ruta")
    usuario = usuario.replace(" ", "_")
    ruta = ruta.replace(" ", "_")
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
# ESTILOS
# =======================================
def estilizar(ws):
    header_font = Font(bold=True)
    header_fill = PatternFill("solid", fgColor="D9D9D9")
    border = Border(left=Side(style="thin"), right=Side(style="thin"),
                    top=Side(style="thin"), bottom=Side(style="thin"))

    for c in ws[1]:
        c.font = header_font
        c.fill = header_fill
        c.border = border

    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.border = border


# =======================================
# CREAR EXCEL SI NO EXISTE
# =======================================
def inicializar_excel():
    EXCEL_FILE = archivo_excel()

    if not os.path.exists(EXCEL_FILE):
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

        wb.save(EXCEL_FILE)


# =======================================
# LOGIN
# =======================================
@app.route("/", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        usuario = request.form["usuario"].strip()
        ruta = request.form["ruta"].strip()

        if usuario == "" or ruta == "":
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
# CONTADOR PRINCIPAL
# =======================================
@app.route("/contador")
def contador():
    if "usuario" not in session:
        return redirect("/")
    return render_template("contador.html",
                           categorias_base=categorias_base,
                           conteos=session["conteos"],
                           nuevas=session["nuevas"])


# =======================================
# SUMAR / RESTAR
# =======================================
@app.route("/modificar", methods=["POST"])
def modificar():
    nombre = request.form["categoria"]
    accion = request.form["accion"]

    conteos = session["conteos"]
    nuevas = session["nuevas"]

    if nombre in conteos:
        if accion == "sumar":
            conteos[nombre] += 1
            session["vehiculos_hoy"] += 1
        elif accion == "restar" and conteos[nombre] > 0:
            conteos[nombre] -= 1
            session["vehiculos_hoy"] -= 1
    else:
        if accion == "sumar":
            nuevas[nombre] += 1
            session["vehiculos_hoy"] += 1
        elif accion == "restar" and nuevas[nombre] > 0:
            nuevas[nombre] -= 1
            session["vehiculos_hoy"] -= 1

    session.modified = True
    return jsonify({"ok": True})


# =======================================
# CREAR NUEVA CATEGORÍA
# =======================================
@app.route("/nueva_categoria", methods=["POST"])
def nueva_categoria():
    nombre = request.form["nombre"].strip()

    if nombre == "":
        return jsonify({"error": "Debe escribir un nombre"})

    if nombre in session["conteos"] or nombre in session["nuevas"]:
        return jsonify({"error": "La categoría ya existe"})

    session["nuevas"][nombre] = 0
    session.modified = True
    return jsonify({"ok": True})


# =======================================
# GUARDAR EXCEL + ENVIAR CORREO
# =======================================
@app.route("/guardar", methods=["POST"])
def guardar():
    EXCEL_FILE = archivo_excel()
    wb = openpyxl.load_workbook(EXCEL_FILE)
    conteo = wb["Conteo"]
    hist = wb["Historial"]

    usuario = session["usuario"]
    ruta = session["ruta"]
    fecha = date.today().strftime("%d-%m-%Y")
    hora = datetime.now().strftime("%H:%M:%S")

    def actualizar(cat, cant):
        encontrada = False
        for row in conteo.iter_rows(min_row=2):
            if row[0].value == cat:
                row[1].value = (row[1].value or 0) + cant
                row[2].value = fecha
                row[3].value = ruta
                encontrada = True
                break

        # ➕ Si no existe, crear la fila en Conteo
        if not encontrada:
            conteo.append([cat, cant, fecha, ruta])

        # ❗ Historial NO SE TOCA
        hist.append([fecha, hora, ruta, usuario, cat, cant, session["vehiculos_hoy"]])

    for c, n in session["conteos"].items():
        if n > 0:
            actualizar(c, n)

    for c, n in session["nuevas"].items():
        if n > 0:
            actualizar(c, n)

    total = sum([
        row[1].value for row in conteo.iter_rows(min_row=2)
        if row[0].value != "N° Vehículos"
    ])
# Buscar y eliminar fila "N° Vehículos" si existe
    fila_total_idx = None
    for i, row in enumerate(conteo.iter_rows(min_row=2), start=2):
        if row[0].value == "N° Vehículos":
            fila_total_idx = i
            break

    if fila_total_idx:
        conteo.delete_rows(fila_total_idx)

    # Agregar siempre al final
    conteo.append(["N° Vehículos", total, fecha, ruta])

    wb.save(EXCEL_FILE)

    session["conteos"] = {c: 0 for c in categorias_base}
    session["nuevas"] = {c: 0 for c in session["nuevas"]}
    session.modified = True


    return jsonify({"ok": True, "mensaje": "Datos guardados correctamente."})


# =======================================
# CERRAR SESIÓN
# =======================================
@app.route("/cerrar", methods=["POST"])
def cerrar():
    EXCEL_FILE = archivo_excel()

    try:
        remitente = "pinillamartin76@gmail.com"
        contrasena = "uuli gnbs cecy tdod"
        destinatario = "pinillamartin76@gmail.com"

        mensaje = MIMEMultipart()
        mensaje["From"] = remitente
        mensaje["To"] = destinatario
        mensaje["Subject"] = "Registro de vehículos"

        texto = f"""
Usuario: {session["usuario"]}
Ruta: {session["ruta"]}
Archivo enviado correctamente.
"""
        mensaje.attach(MIMEText(texto, "plain"))

        with open(EXCEL_FILE, "rb") as adj:
            parte = MIMEBase("application", "octet-stream")
            parte.set_payload(adj.read())
            encoders.encode_base64(parte)
            parte.add_header(
                "Content-Disposition",
                f"attachment; filename={EXCEL_FILE}"
            )
            mensaje.attach(parte)

        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(remitente, contrasena)
        server.send_message(mensaje)
        server.quit()

        session.clear()
        return jsonify({
            "ok": True,
            "mensaje": "📧 El correo fue enviado exitosamente."
        })

    except Exception as e:
        return jsonify({
            "ok": False,
            "mensaje": f"No se pudo enviar el correo: {e}"
        })


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)


