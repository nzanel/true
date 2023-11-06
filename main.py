from kivy.app import App
from kivy.metrics import dp
from kivy.clock import Clock
from kivy.uix.gridlayout import GridLayout
from kivy.uix.anchorlayout import AnchorLayout
    
#importación para apertura y trabajo con excel

from openpyxl import load_workbook
#importación para la fecha
import datetime
#importación para correo
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
# ---------------------------- CONSTANTS ------------------------------- #
CHEQUEO = {"Herramientas Necesarias": "c8", "Elementos de Maniobra": "c9", "Puesta a Tierra": "c10", "EPP": "c11", "Vallas, Conos de Señalización": "c12", "Extintor": "c13","Botiquín Completo": "c14", "Kit de Contención de Derrames": "c15", "Estacionamiento correcto de vehículo": "g8", "Vallado-señalización de zona de trabajo": "g9", "Colocación de EPP acorde a la tarea": "g10", "Trabajo Sin Tensión": "k8", "Trabajo Con Tensión": "k9", "Trabajo Con Tensión a Distancia": "k10", "Trabajo Con Tensión a Contacto": "k11", "Trabajo Con Tensión a Potencial": "k12", "Equipos y coberturas aislantes": "g15", "Condiciones climáticas adecuadas": "g16", "Recierres bloqueados": "g17", "Distancia de seguridad eléctrica": "g18", "Corte efectivo de fuente de tensión": "k15", "Comprobación de ausencia de tensión": "k16", "Bloqueo/señalización de equipos": "k17", "Colocación de puesta a tierra": "k18", "Caída a distinto nivel": "c20", "Exposición a ruidos": "c21", "Exposición a gases/polvos/vapores": "c22", "Caída de elementos desde altura": "c23", "Contacto eléctrico directo": "c24", "Contacto eléctrico indirecto": "c25", "Quemaduras": "g20", "Carga física": "g21", "Carga térmica": "g22", "Incendio": "g23", "Explosión": "g24", "Exposición a agentes químicos": "g25", "Atrapamiento": "k20", "Ataque de animales": "k21", "Proyección de objetos": "k22", "Arco eléctrico": "k23", "Atropellamiento": "k24", "Otros": "k25", "¿Es una maniobra de servicio?": "f27", "Comunicación al personal sobre tareas a realizar": "f28", "Nº de orden de maniobra / licencia / permiso de trabajo / ordenativo": "f29", "Comunicación al personal sobre riesgos existentes": "k27", "Comunicación al personal sobre medidas de control en campo": "k28", "Comunicación al personal de acciones ante emergencias": "k29"}

MY_EMAIL = "transfdigdist@gmail.com"
MY_PASSWORD = "aemo bsuy yfod pbil"
# --------------------------- Hora ------------------------------------ #

hora_actual = datetime.datetime.now()
hora_string = str(hora_actual)
hora_actual_corregida = hora_string.split(":")
segundos = hora_actual_corregida[2].split(".")[0]
hora_para_guardado = f"{hora_actual_corregida[0]} {hora_actual_corregida[1]} {segundos}"
fecha_actual_corregida = hora_string.split(" ")
año=hora_actual.year
mes=hora_actual.month
dia=hora_actual.day
fecha_actual= f"{dia}-{mes}-{año}"
# --------------------------- Apertura plantilla ---------------------- #

libro = load_workbook(".\ATS MODELO FINAL.xlsx")
hoja = libro.active
# --------------------------- Funciones ------------------------------- #




class Interface(AnchorLayout):
    lista_seleccion=[]
    def checking(self, checkbox, labelid):
        if (checkbox.active):
            Interface.lista_seleccion.append(labelid.text)       
        
        if (checkbox.active==False):
            Interface.lista_seleccion.remove(labelid.text)
        
        print(Interface.lista_seleccion)
    def guardar_y_enviar(self):
        for check in Interface.lista_seleccion:
            for key in CHEQUEO:
                if key==check:
                    hoja[CHEQUEO[key]] = "SI"
        
        hoja["b3"]=self.ids.ubicacion.text
        hoja["j3"]=self.ids.fecha.text
        hoja["b4"]=self.ids.descripcion_tarea.text
        hoja["b6"]=self.ids.nombre_encargado.text
        hoja["i6"]=self.ids.sobre.text
        hoja["f29"]=self.ids.licencia.text
        
        ruta_adjunto = f".\ATS COMPLETO {hora_para_guardado}.xlsx"

        libro.save(ruta_adjunto)

        # Iniciamos los parámetros del script
        remitente = "transfdigdist@gmail.com"
        destinatarios = ['nzanel@epec.com.ar', 'zanelnicolas@gmail.com']
        asunto = '[ATS] Correo de prueba'
        cuerpo = 'Aguante Gerencia Distribucion!'
        
        nombre_adjunto = f".\ATS COMPLETO {hora_para_guardado}.xlsx"

        # Creamos el objeto mensaje
        mensaje = MIMEMultipart()

        # Establecemos los atributos del mensaje
        mensaje['From'] = remitente
        mensaje['To'] = ", ".join(destinatarios)
        mensaje['Subject'] = asunto

        # Agregamos el cuerpo del mensaje como objeto MIME de tipo texto
        mensaje.attach(MIMEText(cuerpo, 'plain'))

        # Abrimos el archivo que vamos a adjuntar
        archivo_adjunto = open(ruta_adjunto, 'rb')

        # Creamos un objeto MIME base
        adjunto_MIME = MIMEBase('application', 'octet-stream')
        # Y le cargamos el archivo adjunto
        adjunto_MIME.set_payload((archivo_adjunto).read())
        # Codificamos el objeto en BASE64
        encoders.encode_base64(adjunto_MIME)
        # Agregamos una cabecera al objeto
        adjunto_MIME.add_header('Content-Disposition', "attachment; filename= %s" % nombre_adjunto)
        # Y finalmente lo agregamos al mensaje
        mensaje.attach(adjunto_MIME)

        # Creamos la conexión con el servidor
        sesion_smtp = smtplib.SMTP('smtp.gmail.com', 587)

        # Ciframos la conexión
        sesion_smtp.starttls()

        # Iniciamos sesión en el servidor
        sesion_smtp.login(MY_EMAIL, MY_PASSWORD)

        # Convertimos el objeto mensaje a texto
        texto = mensaje.as_string()

        # Enviamos el mensaje
        sesion_smtp.sendmail(remitente, destinatarios, texto)

        # Cerramos la conexión
        sesion_smtp.quit()

        

class AtsApp(App):
    pass



AtsApp().run()