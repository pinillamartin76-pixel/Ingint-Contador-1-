from flask import Flask, render_template, request, redirect, session, jsonify
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side
from datetime import datetime, date
import os
import base64
import requests

app = Flask(__name__)
app.secret_key = "super_secret_key_123"

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
# ESTILOS
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
# MODIFICAR CONTADORES
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
    EXCEL_FILE = archivo_excel()
    wb = openpyxl.load_workbook(EXCEL_FILE)
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
            fecha, hora, session["ruta"],
            session["usuario"], cat, cant,
            session["vehiculos_hoy"]
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

    wb.save(EXCEL_FILE)

    session["conteos"] = {c: 0 for c in categorias_base}
    session["nuevas"] = {c: 0 for c in session["nuevas"]}
    session.modified = True

    return jsonify(ok=True, mensaje="Datos guardados correctamente.")


# =======================================
# CERRAR SESIÓN + ENVIAR CORREO (RESEND)
# =======================================
@app.route("/cerrar", methods=["POST"])
def cerrar():
    EXCEL_FILE = archivo_excel()

    try:
        with open(EXCEL_FILE, "rb") as f:
            archivo_base64 = base64.b64encode(f.read()).decode()

        payload = {
            "from": "Registro Vehículos <onboarding@resend.dev>",
            "to": [os.environ.get("pinillamartin76@gmail.com")],
            "subject": "Registro de vehículos",
            "html": f"""
                <p><b>Usuario:</b> {session['usuario']}</p>
                <p><b>Ruta:</b> {session['ruta']}</p>
                <p>Archivo adjunto.</p>
            """,
            "attachments": [{
                "filename": EXCEL_FILE,
                "content": archivo_base64
            }]
        }

        headers = {
            "Authorization": f"Bearer {os.environ.get('RESEND_API_KEY')}",
            "Content-Type": "application/json"
        }

        r = requests.post(
            "https://api.resend.com/emails",
            headers=headers,
            json=payload
        )

        r.raise_for_status()

        session.clear()
        return jsonify(ok=True, mensaje="📧 El correo fue enviado exitosamente.")

    except Exception as e:
        return jsonify(ok=False, mensaje=f"No se pudo enviar el correo: {e}")

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)




