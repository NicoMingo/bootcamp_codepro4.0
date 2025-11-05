import random, pandas, smtplib, os, sys, tkinter as tk
from tkinter import filedialog, messagebox, ttk
from email.message import EmailMessage

# -------------------------
# FUNCION para detectar que columna contiene los emails
# -------------------------
def detectar_columna(planilla_de_excel):
    columnas = []
    for c in planilla_de_excel.columns:
        columnas.append(c.lower())

    for columna_correcta in ['email', 'correo', 'mail', 'correos', 'emails', 'mails']:
        if columna_correcta in columnas:
            return planilla_de_excel.columns[columnas.index(columna_correcta)]
    return planilla_de_excel.columns[0]


# -------------------------
# FUNCION para generar contraseñas aleatorias
# -------------------------
def generar_contrasenha():
    letters = ['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm', 'n', 'o', 'p', 'q', 'r', 's', 't', 'u', 'v', 'w', 'x', 'y', 'z', 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']
    numbers = ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9']
    symbols = ['!', '#', '$', '%', '&', '(', ')', '*', '+']

    cantidad_de_letras = int(random.uniform(2, 5))
    cantidad_de_numeros = int(random.uniform(2, 5))
    cantidad_de_simbolos = int(random.uniform(2, 5))

    caracteres_para_lista = []

    for letra in range(0, cantidad_de_letras):
        caracteres_para_lista.append(random.choice(letters))
    for numero in range(0, cantidad_de_numeros):
        caracteres_para_lista.append(random.choice(numbers))
    for simbolo in range(0, cantidad_de_simbolos):
        caracteres_para_lista.append(random.choice(symbols))

    lista_final = []

    for i in range(len(caracteres_para_lista) - 1, -1, -1):
        letra_guardada = random.choice(caracteres_para_lista)
        lista_final.append(letra_guardada)
        if letra_guardada in caracteres_para_lista:
            caracteres_para_lista.remove(letra_guardada)

    return "".join(lista_final)


# -------------------------
# FUNCION PRINCIPAL: Generar contraseñas y enviar correos
# -------------------------
def procesar_archivo(ruta):
    archivo = ruta
    planilla = pandas.read_excel(archivo, engine='openpyxl')

    titulo_de_columna_de_correos = detectar_columna(planilla)

    # --- Eliminar correos duplicados ---
    planilla = planilla.drop_duplicates(subset=[titulo_de_columna_de_correos], keep='first').reset_index(drop=True)
    print("Se eliminaron los correos duplicados, solo se conservaron los primeros registros.")

    # Creamos la nueva columna para las contraseñas
    correos_faltantes = []

    if "Password" not in planilla.columns:
        planilla["Password"] = None

    for i in range(len(planilla)): 
        if pandas.isnull(planilla.loc[i, "Password"]) or str(planilla.loc[i, "Password"]).strip() == "":
            planilla.loc[i, "Password"] = generar_contrasenha()
            correos_faltantes.append(str(planilla.loc[i, titulo_de_columna_de_correos]).strip())
        else:
            print(f"{str(planilla.loc[i, titulo_de_columna_de_correos])} ya tiene una contrasena")
            continue
    
    planilla.to_excel(archivo, index=False, engine='openpyxl')

    # Envio de correos
    EMAIL_REMITENTE = "nicomingouptp@gmail.com"
    CLAVE_APP = "acubadbevcypdpmk"

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
        smtp.login(EMAIL_REMITENTE, CLAVE_APP)
        for _, fila in planilla.iterrows():
            correo = str(fila[titulo_de_columna_de_correos]).strip()
            contrasena = fila["Password"]

            if "@" not in correo:
                continue
            elif correo in correos_faltantes:
                msg = EmailMessage()
                msg["Subject"] = "Tu nueva contraseña"
                msg["From"] = EMAIL_REMITENTE
                msg["To"] = correo
                msg.set_content(f"Hola\n\nTus nueva contraseña es:\n{contrasena}\n\nSaludos.")
                smtp.send_message(msg)
                print(f"Enviado a {correo}")
            else:
                continue


# -------------------------
# INTERFAZ GRAFICA (Launcher)
# -------------------------
def seleccionar_archivo():
    ruta = filedialog.askopenfilename(
        title="Seleccionar archivo Excel",
        filetypes=[("Archivos Excel", "*.xlsx")]
    )
    if ruta:
        archivo_entry.delete(0, tk.END)
        archivo_entry.insert(0, ruta)


def enviar_contraseñas():
    ruta = archivo_entry.get().strip()
    if not ruta or not os.path.isfile(ruta):
        messagebox.showerror("Error", "Selecciona un archivo válido")
        return

    progreso["value"] = 0
    root.update_idletasks()

    try:
        procesar_archivo(ruta)
        progreso["value"] = 100
        messagebox.showinfo("Listo", "Las contraseñas fueron enviadas correctamente")
    except Exception as e:
        messagebox.showerror("Error", f"Ocurrió un error:\n{e}")
        progreso["value"] = 0


# -------------------------
# CREAR VENTANA PRINCIPAL
# -------------------------
root = tk.Tk()
root.title("Launcher - Enviar Contraseñas")
root.geometry("500x200")
root.eval('tk::PlaceWindow . center')

archivo_label = tk.Label(root, text="Archivo Excel:", font=("Arial", 12))
archivo_label.pack(pady=10)

archivo_entry = tk.Entry(root, width=50, font=("Arial", 12))
archivo_entry.pack(pady=5)

seleccionar_btn = tk.Button(root, text="Seleccionar archivo", command=seleccionar_archivo, font=("Arial", 12))
seleccionar_btn.pack(pady=5)

enviar_btn = tk.Button(root, text="Enviar Contraseñas", command=enviar_contraseñas, font=("Arial", 12), bg="#4CAF50", fg="white")
enviar_btn.pack(pady=10)

progreso = ttk.Progressbar(root, length=400, mode='determinate')
progreso.pack(pady=5)

root.mainloop()
