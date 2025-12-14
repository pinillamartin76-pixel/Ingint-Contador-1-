from flask import Flask, render_template, request, redirect, session, jsonify, send_file
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side
from datetime import datetime, date
import os
import smtplib
from email.message import EmailMessage

app = Flask(name)
app.secret_key = os.environ.get("SECRET_KEY", "super_secret_key_123")

=======================================

ARCHIVO DINÁMICO POR USUARIO + RUTA

=======================================

def archivo_excel():
usuario = session.get("usuario", "user").replace(" ", "")
ruta = session.get("ruta", "ruta").replace(" ", "")
return f"registro_{usuario}_{ruta}.xlsx"

=======================================

CATEGORÍAS BASE

=======================================

categorias_base = [
"Autos", "Camionetas", "Micro Bus", "Mini Bus", "Bus", "Omnibus",
"Camiones de 1 Eje", "Camiones de 2 Ejes", "Camiones 3 Ejes o más",
"Motos", "Jeeps", "Bicicletas", "Peatones", "V.A", "V.C",
"Tracción Animal", "Rickshaw"
]

=======================================

ESTILOS EXCEL

=======================================

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

=======================================

CREAR EXCEL

=======================================

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

=======================================

LOGIN

=======================================

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

=======================================

CONTADOR

=======================================

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

=======================================

MODIFICAR

=======================================

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

=======================================

NUEVA CATEGORÍA

=======================================

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

=======================================

GUARDAR EXCEL

=======================================

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

=======================================

ABRIR / DESCARGAR EXCEL (NUEVO)

=======================================

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

=======================================

ENVÍO DE CORREO SMTP (GMAIL)

=======================================

def enviar_correo_smtp(archivo):
mail_user = os.environ.get("MAIL_USER")
mail_pass = os.environ.get("MAIL_PASS")
mail_to = os.environ.get("MAIL_TO")

if not mail_user or not mail_pass:  
    raise Exception("MAIL_USER o MAIL_PASS no configurados")  

msg = EmailMessage()  
msg["From"] = mail_user  
msg["To"] = mail_to  
msg["Subject"] = "Registro de vehículos"  

msg.set_content(  
    f"Usuario: {session['usuario']}\n"  
    f"Ruta: {session['ruta']}\n\n"  
    "Se adjunta el archivo Excel."  
)  

with open(archivo, "rb") as f:  
    msg.add_attachment(  
        f.read(),  
        maintype="application",  
        subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",  
        filename=os.path.basename(archivo)  
    )  

with smtplib.SMTP("smtp.gmail.com", 587) as server:  
    server.starttls()  
    server.login(mail_user, mail_pass)  
    server.send_message(msg)

=======================================

CERRAR SESIÓN + CORREO

=======================================

@app.route("/cerrar", methods=["POST"])
def cerrar():
archivo = archivo_excel()

try:  
    enviar_correo_smtp(archivo)  
    session.clear()  
    return jsonify(ok=True, mensaje="📧 Correo enviado exitosamente.")  

except Exception as e:  
    return jsonify(ok=False, mensaje=f"Error enviando correo: {e}")

=======================================

RUN

=======================================

if name == "main":
app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
