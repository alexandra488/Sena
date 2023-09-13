from tkinter import ttk
import customtkinter as Ctk
from tkinter import*
from tkinter import messagebox,filedialog
import pymysql
import openpyxl
from tkcalendar import DateEntry
import tkinter as tk
from tkinter.ttk import Combobox
from fpdf.fpdf import FPDF
import subprocess
import webbrowser
from datetime import datetime
import os
Ctk.set_appearance_mode("system")  # Modes: system (default), light, dark
Ctk.set_default_color_theme("blue")  # Themes: blue (default), dark-blue, green
current_date1 = datetime.now().strftime("%d")
current_date2 = datetime.now().strftime("%m")
current_date3 = datetime.now().strftime("%Y")



class PDF(FPDF):

    def header(self):
        # Logo
        self.set_font('Arial', '', 10)
        self.ln(3)
        self.multi_cell(align="J",h=15,w=0,txt="Certificación")
        self.image('./img/logo.jpg', x = 100, y = 10, w = 15, h = 15)
        self.ln(3)

        # Arial bold 25
        

        """ # Title  
        self.multi_cell(w =185, h = 7, txt = 'EL SUSCRITO SUBDIRECTOR (E) DEL CENTRO PARA LA INDUSTRIA DE LA COMUNICACIÓN GRÁFICA DEL SERVICIO NACIONAL DE APRENDIZAJE SENA ', border = 0,
                align = 'C', fill = 0)
        self.multi_cell(align="C",h=13,w=0,txt="HACE CONSTAR")
        # Line break
        self.ln(5) """

    # Page footer
    def footer(self):
        # Position at 1.5 cm from bottom
        self.set_y(-40)

        # Arial italic 8
        self.image('./img/ICONTEC.jpg', x = 183, y = 230, w = 15, h = 43)
        self.set_font('Arial', 'I', 9)
        self.multi_cell(w=0,align="C", fill = 0 ,h=5,txt="Ministerio de Trabajo\n SERVICIO NACIONAL DE APRENDIZAJE")
        # Activa el subrayado
        
        self.multi_cell(w=0,align="C", fill = 0 ,h=5,txt="Centro para la Industria de la Comunicación Gráfica de la Regional Distrito Capital")
        #self.set_underline(False)
        self.multi_cell(w=0,align="C", fill = 0 ,h=5,txt="Dirección Calle 15 No. 31-42, Ciudad Bogotá - PBX (57 1) 5461500")
        #self.set_link(link="https://www.sena.edu.co/")
        self.multi_cell(w=0,align="C", fill = 0 ,h=5,txt="www.sena.edu.co - Línea gratuita nacional: 01 8000 9 10 270"+"GTH-F-131 V03")
        self.set_link(False)
        self.ln(5)
        # Page number
        self.cell(w = 0, h = 10, txt =  'Pagina ' + str(self.page_no()) + '/{nb}', border = 0,
                align = 'C', fill = 0)   

class Contratista: 
        
    def __init__(self, root):
        self.wind =root 
        self.wind.title ("Contratista")
        self.wind.geometry("1090x600")
        #self.wind.attributes('-fullscreen')
        
        self.wind.config(bg="teal")
        
        #self.win.withdraw()
        def Salir():                
                root.quit()
                
                
        
        self.wind.protocol('WM_DELETE_WINDOW', Salir)

        Frame1=Ctk.CTkFrame(self.wind )
        Frame2=LabelFrame(self.wind)
#dimension Bordes
        Frame1.pack(fill="both", expand="yes", padx=20, pady=15 )
        Frame2.pack(fill="both", expand="yes", padx=20, pady=15 )
#====================================================================================================
        ID = StringVar()
        Nombre = StringVar()
        Apellido = StringVar()
        Cedula = StringVar()
        CiudadExpedicion= StringVar()
        Direccion = StringVar()
        Contrato = StringVar()
        ARL = StringVar()
        EPS = StringVar()
        FechaNacimiento=StringVar()
        FechaNacimiento = StringVar()
        genero = StringVar()
        Rh = StringVar()
        Celular=StringVar()
        Telefono=StringVar()
        Correo=StringVar()
        Cargo=StringVar()
        FechaDocumento=StringVar()
        Dependencia=StringVar()
        Jefe=StringVar()
        DescripcionContrato=StringVar()
        ValorContrato=StringVar()
        AutorizacionContrato=StringVar()
        PAA=StringVar()
        Banco=StringVar()
        TipoCuenta=StringVar()
        NumeroCuenta=StringVar()
        CDP=StringVar()
        FechaCDP=StringVar()
        CRP=StringVar()
        server=StringVar()
        UserServer=StringVar()
        passServer=StringVar()
        baseServer=StringVar()
        





##funcion para llamar IP conexion varios equipos REEMPLAZAR localhost



        def cargarServer():
                if os.path.isfile("./server/server.txt"):
                        fileser=open("./server/server.txt","r")
                        contador = 0
                        for linea in fileser:
                                if contador ==0:
                                        server.set(linea.strip())
                                        print(server.get())
                                elif contador ==1:
                                        UserServer.set(linea.strip()) 
                                        print(UserServer.get()) 
                                elif contador ==2:
                                        passServer.set(linea.strip())
                                        print(passServer.get()) 
                                elif contador ==3:
                                        baseServer.set(linea.strip())
                                        print(baseServer.get())
                                elif contador >= 4 :
                                        break     
                                contador += 1
#==========================================================================================
# FUNCIONES DE LOS BOTONEaS

        def Agregar():
                try:
                        

                        if ID.get() == "" or Nombre.get() == "" or Apellido.get() == "" or Cedula.get() == ""or CiudadExpedicion.get() == ""or Direccion.get()=="" or Contrato.get() =="" or ARL.get() =="" or EPS.get()=="" or FechaNacimiento.get()=="" or genero.get()=="" or Rh.get()=="" or Celular.get()=="" or Telefono.get()=="" or Correo.get()=="" or Cargo.get()=="" or FechaDocumento.get()=="" or Dependencia.get()=="" or Jefe.get()=="" or Jefe.get()=="" or DescripcionContrato.get()=="" or ValorContrato.get()=="" or AutorizacionContrato.get()=="" or PAA.get()=="" or Banco.get()=="" or TipoCuenta.get()=="" or NumeroCuenta.get()=="" or CDP.get()=="" or FechaCDP.get()=="" or CRP.get()=="":
                                messagebox.showerror("Por favor", "Ingresa la información correcta")

                        
                        else:
                                
                                base= pymysql.connect(host=server.get(),user=UserServer.get(), password=passServer.get(), database=baseServer.get())
                                cursor =base.cursor()
                                cursor.execute("insert into cliente values(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)",(
                                        
                                ID.get(),
                                Nombre.get(),
                                Apellido.get(),
                                Cedula.get(),
                                CiudadExpedicion.get(),
                                Direccion.get(),
                                Contrato.get(),
                                ARL.get(),
                                EPS.get(),
                                FechaNacimiento.get(),
                                genero.get(),
                                Rh.get(),
                                Celular.get(),
                                Telefono.get(),
                                Correo.get(),
                                Cargo.get(),
                                FechaDocumento.get(),
                                Dependencia.get(),
                                Jefe.get(),
                                DescripcionContrato.get(),
                                ValorContrato.get(),
                                AutorizacionContrato.get(),
                                PAA.get(),
                                Banco.get(),
                                TipoCuenta.get(),
                                NumeroCuenta.get(),
                                FechaCDP.get(),
                                CRP.get(),
                                CDP.get() ))
                                
                                base.commit()
                                base.close()
                                #messagebox.showinfo("Datos completados","se agregaron correctamente")
                except base.Error as er:
                        codigo_error = er.args[0]                
                        messagebox.showwarning("Error de Conexion ","Error "+str({codigo_error}) )
                        #messagebox.showwarning("Error al Agregar Registro","revise los datos")
                        
                        


        def Limpiar():
                self.entID.delete(0, END)
                self.entNombre.delete(0, END)
                self.entApellido.delete(0, END)
                self.entCedula.delete(0,END)
                self.entCiudadExpedicion.delete(0,END)
                self.entDireccion.delete(0, END)
                self.entContrato.delete(0, END)
                self.entARL.delete(0, END)
                self.entEPS.delete(0, END)
                self.entFechaNacimiento.delete(0,END)
                self.entgenero.delete(0,END)
                self.entRh.delete(0,END)
                self.entCelular.delete(0,END)
                self.entTelefono.delete(0,END)
                self.entCorreo.delete(0,END)
                self.entCargo.delete(0,END)
                self.entFechaDocumento.delete(0,END)
                self.entDependencia.delete(0,END)
                self.entJefe.delete(0,END)
                self.entDescripcionContrato.delete(0,END)
                self.entValorContrato.delete(0,END)
                self.entAutorizacionContrato.delete(0,END)
                self.entPAA.delete(0,END)
                self.entBanco.delete(0,END)
                self.entTipoCuenta.delete(0,END)
                self.entNumeroCuenta.delete(0,END)
                self.entCDP.delete(0,END)
                self.entFechaCDP.delete(0,END)
                self.entCRP.delete(0,END)
                


        def Mostrar():
                try:
                        base = pymysql.connect(host=server.get(),user=UserServer.get(), password=passServer.get(), database=baseServer.get())
                        cursor= base.cursor()
                        cursor.execute("select * from cliente")
                        result = cursor.fetchall()
                        
                        if len(result) !=0:
                                self.trv.delete(*self.trv.get_children())
                                for row in result:
                                        self.trv.insert('',END,values= row)
                        base.commit()
                        base.close()
                except pymysql.Connection.Error as er:
                        codigo_error_extendido = er.args[0]                
                        messagebox.showwarning("Error de Conexion ","Error "+str({codigo_error_extendido})+" No se ha podido conectar al Servidor " +server.get() )

        def Prueba():
               base = pymysql.connect(host=server.get(),user=UserServer.get(), password=passServer.get(), database=baseServer.get())
               cursor= base.cursor()
               cursor.execute("select ARL from cliente")
               base.commit()
               base.close()
                #funcion para mostrar los datos en cajas de texto
        
        
        def traineeInfo(ev):
                viewInfo = self.trv.focus()
                learnerData = self.trv.item(viewInfo)
                row = learnerData['values']
                ID.set(row[0])
                Nombre.set(row[1])
                Apellido.set(row[2])                
                Cedula.set(row[3])
                CiudadExpedicion.set(row[4])
                Direccion.set(row[5])
                Contrato.set(row[6])
                ARL.set(row[7])
                EPS.set(row[8])
                FechaNacimiento.set(row[9])
                genero.set(row[10])
                Rh.set(row[11])
                Celular.set(row[12])
                Telefono.set(row[13])
                Correo.set(row[14])
                Cargo.set(row[15])
                FechaDocumento.set(row[16])
                Dependencia.set(row[17])
                Jefe.set(row[18])
                DescripcionContrato.set(row[19])
                ValorContrato.set(row[20])
                AutorizacionContrato.set(row[21])
                PAA.set(row[22])
                Banco.set(row[23])
                TipoCuenta.set(row[24])
                NumeroCuenta.set(row[25])
                CDP.set(row[26])
                FechaCDP.set(row[27])
                CRP.set(row[28])




        def Actualizar():
                
                try:
                        
                        if ID.get() == "" or Nombre.get() == "" or Apellido.get() == "" or Cedula.get() == ""or CiudadExpedicion.get() == ""or Direccion.get()=="" or Contrato.get() =="" or ARL.get() =="" or EPS.get()=="" or FechaNacimiento.get()=="" or genero.get()=="" or Rh.get()=="" or Celular.get()=="" or Telefono.get()=="" or Correo.get()=="" or Cargo.get()=="" or FechaDocumento.get()=="" or Dependencia.get()=="" or Jefe.get()=="" or Jefe.get()=="" or DescripcionContrato.get()=="" or ValorContrato.get()=="" or AutorizacionContrato.get()=="" or PAA.get()=="" or Banco.get()=="" or TipoCuenta.get()=="" or NumeroCuenta.get()=="" or CDP.get()=="" or FechaCDP.get()=="" or CRP.get()=="":
                                messagebox.showerror("Por favor", "Completar Todos Los Campos")
                        else:
                                base = pymysql.connect(host=server.get(),user=UserServer.get(), password=passServer.get(), database=baseServer.get())
                                cursor = base.cursor()
                                cursor.execute("update cliente set nombre=%s,apellido=%s,cedula=%s,ciudadExpedicion=%s,direccion=%s,contrato=%s,ARL=%s,EPS=%s,FechaNacimiento=%s,Genero=%s,Rh=%s,Celular=%s,Telefono=%s,Correo=%s,Cargo=%s,FechaDocumento=%s,Dependencia=%s,Jefe=%s,DescripcionContrato=%s,ValorContrato=%s,AutorizacionContrato=%s,PAA=%s,Banco=%s,TipoCuenta=%s,NumeroCuenta=%s,CDP=%s,FechaCDP=%s,CRP=%s where id=%s",(
                                Nombre.get(),
                                Apellido.get(),
                                Cedula.get(),
                                CiudadExpedicion.get(),
                                Direccion.get(),
                                Contrato.get(),
                                ARL.get(),
                                EPS.get(),
                                FechaNacimiento.get(),
                                genero.get(),
                                Rh.get(),
                                Celular.get(),
                                Telefono.get(),
                                Correo.get(),
                                Cargo.get(),
                                FechaDocumento.get(),
                                Dependencia.get(),
                                Jefe.get(),
                                DescripcionContrato.get(),
                                ValorContrato.get(),
                                AutorizacionContrato.get(),
                                PAA.get(),
                                Banco.get(),
                                TipoCuenta.get(),
                                NumeroCuenta.get(),
                                CDP.get(),
                                FechaCDP.get(),
                                CRP.get(),






                                ID.get()
                                
                                ))
                                
                                base.commit()
                                Mostrar()
                                base.close()
                                messagebox.showinfo("La información ha sido actualizada","Se ha actualizado Correctamente")
                                
                except:
                        next
                        
                
                
        def Eliminar():
                delete=messagebox.askquestion("Eliminar","Realmente Deseas Eliminar el Registro")
                if delete == "yes":
                        base= pymysql.connect(host=server.get(),user=UserServer.get(), password=passServer.get(), database=baseServer.get())
                        cursor =base.cursor()
                        cursor.execute("delete from cliente where id=%s",ID.get())
                        base.commit()
                        Mostrar()
                        base.close()
                        Limpiar()
                        messagebox.showinfo("La informacion ha sido eliminada","El Registro Se ha Eliminado Correctamente")
                
                
        def crearpdf():
                try:
                        if ID.get() == "" or Nombre.get() == "" or Apellido.get() == "" or Cedula.get() == ""or CiudadExpedicion.get() == ""or Direccion.get()=="" or Contrato.get() =="" or ARL.get() =="" or EPS.get()=="" or FechaNacimiento.get()=="" or genero.get()=="" or Rh.get()=="" or Celular.get()=="" or Telefono.get()=="" or Correo.get()=="" or Cargo.get()=="" or FechaDocumento.get()=="" or Dependencia.get()=="" or Jefe.get()=="" or Jefe.get()=="" or DescripcionContrato.get()=="" or ValorContrato.get()=="" or AutorizacionContrato.get()=="" or PAA.get()=="" or Banco.get()=="" or TipoCuenta.get()=="" or NumeroCuenta.get()=="" or CDP.get()=="" or FechaCDP.get()=="" or CRP.get()=="":
                                messagebox.showerror("Por favor", "Ingresa la información correcta")
                        else:
       #import plantillapdf
       #from plantillapdf import PDF
       # Instantiation of inherited class
                                pdf = PDF(orientation="portrait",format="A4")
                                pdf.alias_nb_pages()
                                pdf.add_page()
                                pdf.set_font('Arial', 'B', 9)
                # Title
                                pdf.multi_cell(w =185, h = 7, txt = 'EL SUSCRITO SUBDIRECTOR (E) DEL CENTRO PARA LA INDUSTRIA DE LA COMUNICACIÓN GRÁFICA DEL SERVICIO NACIONAL DE APRENDIZAJE SENA ', border = 0,align = 'C', fill = 0)
                                pdf.multi_cell(align="C",h=13,w=0,txt="HACE CONSTAR")
                        # Line break
                                pdf.ln(5)
                        #from tkinter import StringVar
                
                                pdf.set_font('Arial', '', 9)
                                pdf.multi_cell(w = 0, h = 7, txt = 'Que la señor(a) '+Nombre.get()+" "+Apellido.get()+', identificado con cédula de ciudadanía No. '+Cedula.get()+' de  '+CiudadExpedicion.get()+ ', celebró con EL SERVICIO NACIONAL DE APRENDIZAJE SENA, el siguiente contrato de prestación de servicios personales regulados por la Ley 80 de 1993 (Estatuto General de Contratación de la Administración Pública), modificada por la Ley 1150 de 2007, Decreto 1082 de 2015 y sus demás Decretos o normas reglamentarias, como se describe a continuación:',align ='J', fill = 0)
                                pdf.ln(h=2)
                                pdf.multi_cell(w=0,h=7,txt="1.Número de Contrato:  "+Contrato.get(), align='J')
                                pdf.ln(h=2)
                                pdf.multi_cell(w=0,h=7,txt="Objeto: Prestar servicios de apoyo de carácter temporal a la gestión en los procesos académicos administrativos de la Coordinación Académica virtual y a distancia el Centro para la Industria en la Comunicación Gráfica.", align='J')
                                pdf.ln(h=2)
                                pdf.multi_cell(w=0,h=7,txt="Plazo de ejecución: 1 de febrero de 2023 hasta el 28 de diciembre de 2023", align='J')
                                pdf.ln(h=2)
                                pdf.multi_cell(w=0,h=7,txt="Fecha de Inicio de Ejecución: 1 de febrero de 2023", align='J')
                                pdf.ln(h=2)
                                pdf.multi_cell(w=0,h=7,txt="Fecha de Terminación de Contrato: 28 de diciembre de 2023", align='J')
                                pdf.ln(h=2)
                                pdf.multi_cell(w=0,h=7,txt="Término de Ejecución: El término real ejecutado por la contratista es de Diez (10) meses y Dieciseis (16) días, con una prórroga de Doce (12) días; por lo tanto, el termino total es de Diez (10) meses y Veintiocho (28) Días.", align='J')
                                pdf.ln(h=2)
                                pdf.multi_cell(w=0,h=7,txt="Prorroga: Doce (12) días.", align='J')
                                pdf.ln(h=2)
                                pdf.multi_cell(w=0,h=7,txt="Adición: Hasta el 28 de diciembre de 2023", align='J')
                                pdf.ln(h=2)
                                pdf.multi_cell(w=0,h=7,txt="Valor: El valor del contrato para todos los efectos legales y fiscales se fijó en la suma de VEINTICINCO MILLONES DOSCIENTOS OCHENTA MIL PESOS M/CTE. ( 25.280.000), con una adición por valor de NOVECIENTOS SESENTA MIL PESOS M/CTE. (960.000), por lo tanto, el valor total es de VEINTISEIS MILLONES DOSCIENTOS CUARENTA MIL PESOS M/CTE (26.240.000).", align='J')
                                pdf.add_page()
                                pdf.multi_cell(w=0,h=7,txt="Obligaciones Específicas del Contrato: 1. Apoyar a la Coordinación académica en lo relacionado con el sistema de gestión documental y los demás procesos que hacen parte de la gestión de la formación profesional integral. 2. Gestionar las diferentes actividades administrativas relacionadas con atención al cliente interno y externo y registro de información en los aplicativos correspondientes a la Coordinación Académica de acuerdo con lo establecido en el proceso de Gestión de Formación Profesional Integral. 3.Apoyar la implementación de las acciones de mejora que surjan del proceso de la gestión de la Formación Profesional Integral.4. Tramitar la información necesaria para dar respuesta a las PQRS relacionadas con la formación profesional integral y que sean de la competencia de la Coordinación Académica teniendo en cuenta la promesa de valor institucional.5. Apoyar la depuración de las fichas de formación a cargo de la Coordinación, según requerimientos presentados.6. Apoyar a la Coordinación Académica en la realización de Comités de Evaluación y Seguimiento hasta la elaboración de los actos administrativos. 7. Apoyar en el proceso administrativo y documental de supervisión de contratos de servicios personales y bienes y servicios asignados a la coordinación académica.8. Gestionar el proceso de novedades de aprendices hasta el reporte a la Coordinación de Formación Profesional del Centro. 9. Apoyar y procesar en el aplicativo de gestión académica institucional, en los siguientes aspectos: a) Asociación fichas por programación y reprogramación de competencias técnicas y transversales b) Asociación de los resultados de aprendizaje de las fichas asignadas a los instructores c) Cargue de horas de los instructores de planta y contrato d) Creación de la ruta de formación de las fichas asignadas a los instructores 10. Revisar y actualizar los indicadores de la Coordinación Académica con las respectivas evidencias.11. Verificar el registro de juicios evaluativos de la formación profesional integral en el aplicativo Sofia Plus según lineamientos institucionales y avance del proceso formativo. 12. Vigilar y salvaguardar los bienes e inventarios que son utilizados por la Coordinación académica para su funcionamiento. 13.Ejecutar de manera idónea el objeto del contrato conforme a los lineamientos del Sistema Integrado de Gestión y Autocontrol (SIGA) del SENA, el cual se encuentra documentado en la plataforma compromiso. 14. Mantener la debida reserva sobre los asuntos manejados y conocidos dentro de la ejecución del contrato. 15. Realizar las demás actividades relacionadas con el objeto del contrato que le sean asignadas por el supervisor y/o el subdirector del centro que correspondan a la naturaleza del contrato. 16. Aplicar los procesos y procedimientos establecidos por la entidad, para la gestión documental relacionada con el objeto contractual.", align='J')#pdf.multi_cell(w=0,h=7,txt="Diego Armando Cristancho Cristancho Subdirector (E) Centro para la Industria de la Comunicación Gráfica Servicio Nacional de Aprendizaje SENA", align='J')
                                
                                pdf.multi_cell(w=0,h=7,txt="Se expide a solicitud del interesado, de acuerdo con la información registrada en el sistema ON BASE del SENA y Secop II, a los  "+str(current_date1)+" dias del mes  " +str(current_date2)+" del " +str(current_date3)+".", align='J')

                #pdf.ln(h=2) salto de linea
                                pdf.image('./img/firmapablo.jpg',
                                x = 110, y = 240,
                                w = 10, h = 10)#, link = url
                
                                pdf.image('./img/firmadagny.jpg',
                                x = 110, y = 250,
                                w = 10, h = 10)#, link = url
                                
                                pdf.image('./img/firma.png',
                                x = 100, y = 200,
                                w = 25, h = 25)#, link = url
                
                                pdf.ln(h=10)
                                pdf.set_font('Arial', 'B', 9)
                
                                pdf.multi_cell(w=0,h=7,txt="Diego Armando Cristancho Cristancho ", align='C')
                                pdf.multi_cell(w=0,h=7,txt="Subdirector (E) Centro para la Industria de la Comunicación ", align='C')
                                pdf.multi_cell(w=0,h=7,txt="Gráfica Servicio Nacional de Aprendizaje SENA", align='C')
                                pdf.ln(h=5)
                
                                pdf.multi_cell(w=0,h=7,txt="Proyectó: Juan Pablo Alvarez Sarmiento; Apoyo Administrativo ", align='J')
                                
                                pdf.set_font('Arial', '', 9)
                                
                                pdf.multi_cell(w=0,h=7,txt="Revisó: Dagny Galindo; Profesional en Contratación", align='J')
                                
                                """(((((((((((((((((((((((((CREAR CARPETA Y ABRIR PDF ))))))))))))))))))))))))) """
                                if os.path.exists('./certificados/'):
                                        next
                                else:
                                        subprocess.call("cmd /c mkdir certificados")       
                                pdf.output("./certificados/"+Cedula.get()+'.pdf')
                                #ruta_pdf = os.path.join(os.path.dirname(os.path.abspath(__file__)),"certificados",Cedula.get()+'.pdf')
                                
                                ruta = os.path.abspath("./certificados/"+Cedula.get()+'.pdf')
                                print(ruta)               
                                webbrowser.open_new (ruta)
                                subprocess.call([ruta],shell=True)
                                """(((((((((((((((((((((((((FIN CREAR CARPETA Y ABRIR PDF ))))))))))))))))))))))))) """
                                Frame2=Label(self.wind,bg="teal", text=ruta, font=("calibri", 14 )).pack()
                except:
                        next
       
                #subprocess.call([path],shell=True)
       #subprocess.run([path],shell=True)
                #webbrowser.WindowsDefault(path)
                #os.system(path)
        #crearpdf() 
        def importarExcel():
                
                user=os.environ.get('USERNAME')
                iniciardialogo="C:\\Users\\"+user+"\\Desktop"
                File=filedialog.askopenfilenames(title="Abrir Archivos",initialdir=iniciardialogo,filetypes=[("Excel", "*.xlsm"),("Excel","*.xltx"),("Excel","*.xltm"),("Excel","*.xlsx")])
                #print(File[0])
                path =File[0]
                workbook = openpyxl.load_workbook(path)
                sheet = workbook.active
                contador=0

                list_values = list(sheet.values)
                for row in list_values[1:]:
                        contador+=1
                        ID.set(row[0])
                        Nombre.set(row[1])
                        Apellido.set(row[2])                
                        Cedula.set(row[3])
                        CiudadExpedicion.set(row[4])
                        Direccion.set(row[5])
                        Contrato.set(row[6])
                        ARL.set(row[7])
                        EPS.set(row[8])
                        FechaNacimiento.set(row[9])
                        genero.set(row[10])
                        Rh.set(row[11])
                        Celular.set(row[12])
                        Telefono.set(row[13])
                        Correo.set(row[14])
                        Cargo.set(row[15])
                        FechaDocumento.set(row[16])
                        Dependencia.set(row[17])
                        Jefe.set(row[18])
                        DescripcionContrato.set(row[19])
                        ValorContrato.set(row[20])
                        AutorizacionContrato.set(row[21])
                        PAA.set(row[22])
                        Banco.set(row[23])
                        TipoCuenta.set(row[24])
                        NumeroCuenta.set(row[25])
                        CDP.set(row[26])
                        FechaCDP.set(row[27])
                        CRP.set(row[28])
                        Agregar()
                        Limpiar()
                print (contador)
        def BuscarCodigo(event):
                cargarServer()
                #print(server.get()+"  " +UserServer.get()+"   " +passServer.get()+"   "+baseServer.get())
                try:
                        self.trv.delete(*self.trv.get_children())
                        base= pymysql.connect(host=server.get(),user=UserServer.get(), password=passServer.get(), database=baseServer.get())
                        cursor= base.cursor()
                        cursor.execute('select * from cliente WHERE id LIKE "%'+ID.get()+'%"')
                        #micursor.execute('SELECT * FROM "tblestudiantes" WHERE Nom_Estudiante LIKE "%'+VarNombre.get()+'%"')
                        result = cursor.fetchall()
                        
                        if len(result) !=0:                                
                                for row in result:
                                        self.trv.insert(parent='',index=row[0],iid=row[0],values= row)
                                self.trv.selection_set(row[0])
                                self.trv.see(row[0])
                        base.commit()
                        base.close()
                except pymysql.Connection.Error as er:
                        codigo_error_extendido = er.args[1]                
                        messagebox.showwarning("Error de Conexion ","Error "+str({codigo_error_extendido})+" No se ha podido conectar al Servidor " +server.get() ) 
                               
        def BuscarNombre(event):
                cargarServer()
                #print(server.get()+"  " +UserServer.get()+"   " +passServer.get()+"   "+baseServer.get())
                try:
                        self.trv.delete(*self.trv.get_children())
                        base = pymysql.connect(host=server.get(), user="root", password="", database="base")
                        cursor= base.cursor()
                        cursor.execute('select * from cliente WHERE Nombre LIKE "%'+Nombre.get()+'%"')
                        #micursor.execute('SELECT * FROM "tblestudiantes" WHERE Nom_Estudiante LIKE "%'+VarNombre.get()+'%"')
                        result = cursor.fetchall()
                        
                        if len(result) !=0:                                
                                for row in result:
                                        self.trv.insert(parent='',index=row[0],iid=row[0],values= row)
                                self.trv.selection_set(row[0])
                                self.trv.see(row[0])
                        base.commit()
                        base.close()
                except pymysql.Connection.Error as er:
                        codigo_error_extendido = er.args[1]                
                        messagebox.showwarning("Error de Conexion ","Error "+str({codigo_error_extendido})+" No se ha podido conectar al Servidor " +server.get() ) 
                                       
        def BuscarCedula(event):
                cargarServer()
                #print(server.get()+"  " +UserServer.get()+"   " +passServer.get()+"   "+baseServer.get())
                try:
                        self.trv.delete(*self.trv.get_children())
                        base= pymysql.connect(host=server.get(),user=UserServer.get(), password=passServer.get(), database=baseServer.get())
                        cursor= base.cursor()
                        cursor.execute('select * from cliente WHERE Cedula LIKE "%'+Cedula.get()+'%"')
                        #micursor.execute('SELECT * FROM "tblestudiantes" WHERE Nom_Estudiante LIKE "%'+VarNombre.get()+'%"')
                        result = cursor.fetchall()
                        
                        if len(result) !=0:                                
                                for row in result:
                                        self.trv.insert(parent='',index=row[0],iid=row[0],values= row)
                                self.trv.selection_set(row[0])
                                self.trv.see(row[0])
                        base.commit()
                        base.close()
                except pymysql.Connection.Error as er:
                        codigo_error_extendido = er.args[1]                
                        messagebox.showwarning("Error de Conexion ","Error "+str({codigo_error_extendido})+" No se ha podido conectar al Servidor " +server.get() ) 
                                             
        def BuscarCorreo(event):
                cargarServer()
                #print(server.get()+"  " +UserServer.get()+"   " +passServer.get()+"   "+baseServer.get())
                try:
                        self.trv.delete(*self.trv.get_children())
                        base= pymysql.connect(host=server.get(),user=UserServer.get(), password=passServer.get(), database=baseServer.get())
                        cursor= base.cursor()
                        cursor.execute('select * from cliente WHERE Correo LIKE "%'+Correo.get()+'%"')
                        #micursor.execute('SELECT * FROM "tblestudiantes" WHERE Nom_Estudiante LIKE "%'+VarNombre.get()+'%"')
                        result = cursor.fetchall()
                        
                        if len(result) !=0:                                
                                for row in result:
                                        self.trv.insert(parent='',index=row[0],iid=row[0],values= row)
                                self.trv.selection_set(row[0])
                                self.trv.see(row[0])
                        base.commit()
                        base.close()
                except pymysql.Connection.Error as er:
                        codigo_error_extendido = er.args[1]                
                        messagebox.showwarning("Error de Conexion ","Error "+str({codigo_error_extendido})+" No se ha podido conectar al Servidor " +server.get() ) 
                                                     
        def BuscarContrato(event):
                cargarServer()
                #print(server.get()+"  " +UserServer.get()+"   " +passServer.get()+"   "+baseServer.get())
                try:
                        self.trv.delete(*self.trv.get_children())
                        base= pymysql.connect(host=server.get(),user=UserServer.get(), password=passServer.get(), database=baseServer.get())
                        cursor= base.cursor()
                        cursor.execute('select * from cliente WHERE Contrato LIKE "%'+Contrato.get()+'%"')
                        #micursor.execute('SELECT * FROM "tblestudiantes" WHERE Nom_Estudiante LIKE "%'+VarNombre.get()+'%"')
                        result = cursor.fetchall()
                        
                        if len(result) !=0:                                
                                for row in result:
                                        self.trv.insert(parent='',index=row[0],iid=row[0],values= row)
                                self.trv.selection_set(row[0])
                                self.trv.see(row[0])
                        base.commit()
                        base.close()
                except pymysql.Connection.Error as er:
                        codigo_error_extendido = er.args[1]                
                        messagebox.showwarning("Error de Conexion ","Error "+str({codigo_error_extendido})+" No se ha podido conectar al Servidor " +server.get() ) 
                                                                                                  
                                
                


#=====================================================================================================

# BARRA MENU

        barraMenu=tk.Menu(root,foreground="orange",activebackground="blue")
        root.config(menu=barraMenu, width=300, height=300,cursor="heart",highlightbackground="yellow")
        menuconect=tk.Menu(barraMenu, tearoff=0,activebackground ="orange",activeforeground="red",)
        menuconect.add_command(label="Conectar",activebackground ="green")
        menuconect.add_command(label="Salir",activebackground ="green")
        barraMenu.add_cascade(label="Inicio",menu=menuconect,activebackground ="orange")
        bbddMenu=tk.Menu(barraMenu, tearoff=0,activebackground ="green",activeforeground="red",)
        bbddMenu.add_command(label="Agregar",activebackground ="green",command=Agregar) 
        bbddMenu.add_command(label="Actualizar",activebackground ="green",command=Actualizar)
        bbddMenu.add_command(label="Nuevo",activebackground ="green",command=Limpiar)
        bbddMenu.add_command(label="Crear PDF",activebackground ="green",command=crearpdf)
        bbddMenu.add_command(label="Importar Excel",activebackground ="green",command=importarExcel)
        bbddMenu.add_command(label="Eliminar",activebackground ="green",command=Eliminar)
        barraMenu.add_cascade(label="Registro",menu=bbddMenu,activebackground ="Orange")

        menuventanas=tk.Menu(barraMenu, tearoff=0,activebackground ="green",activeforeground="red",)
        menuventanas.add_command(label="Primero",activebackground ="green")
        menuventanas.add_command(label="Siguiente",activebackground ="green")
        menuventanas.add_command(label="Anterior",activebackground ="green")
        menuventanas.add_command(label="Ultimo",activebackground ="green")
        barraMenu.add_cascade(label="Movimientos",menu=menuventanas,activebackground ="orange")

        menubuscar=tk.Menu(barraMenu, tearoff=0,activebackground ="green",activeforeground="red",)
        menubuscar.add_command(label="Todos",activebackground ="green")
        menubuscar.add_separator()
        menubuscar.add_command(label="Por Email",activebackground ="green",command=lambda :BuscarCorreo(2))
        menubuscar.add_separator()
        menubuscar.add_command(label="Por Codigo",activebackground ="green",command=lambda :BuscarCodigo(2))
        menubuscar.add_separator()
        menubuscar.add_command(label="Por Contrato",activebackground ="green",command=lambda :BuscarContrato(2))
        menubuscar.add_separator()


        menubuscar.add_command(label="Por Nombre",activebackground ="green",command=lambda :BuscarNombre(2))
        menubuscar.add_separator()
        menubuscar.add_command(label="Por Cedula",activebackground ="green",command=lambda :BuscarCedula(2))
        menubuscar.add_separator()
        menubuscar.add_command(label="Vaciar Tabla",activebackground ="green")
        barraMenu.add_cascade(label="Busqueda",menu=menubuscar,activebackground ="orange")
        menuinfo=tk.Menu(barraMenu, tearoff=0,activebackground ="green",activeforeground="red",)
        menuinfo.add_command(label="Acerca de...",activebackground ="green")
        barraMenu.add_cascade(label="Ayuda...",menu=menuinfo,activebackground ="orange")
        #FIN BARRA MENU
#Cajas de texto
        lbl1 = Ctk.CTkLabel(Frame1, text= "ID", width=20)
        lbl1.grid(row=0, column=0,padx=5, pady=3,sticky="e")
        self.entID=Ctk.CTkEntry(Frame1,border_color="teal",textvariable=ID)
        self.entID.grid(row=0,column=1, padx=5, pady=3)
        self.entID.bind("<Return>", lambda event: BuscarCodigo(1))

        lbl2 = Ctk.CTkLabel(Frame1, text= "Nombre", width=20)
        lbl2.grid(row=1, column=0,padx=5, pady=3,sticky="e")
        self.entNombre = Ctk.CTkEntry(Frame1,border_color="teal", textvariable=Nombre)
        self.entNombre.grid(row=1,column=1, padx=5, pady=3)
        self.entNombre.bind("<Return>", lambda event: BuscarNombre(1))

        lbl3 = Ctk.CTkLabel(Frame1, text= "Apellido", width=20)
        lbl3.grid(row=2, column=0,padx=5, pady=3,sticky="e")
        self.entApellido = Ctk.CTkEntry(Frame1,border_color="teal", textvariable=Apellido)
        self.entApellido.grid(row=2,column=1, padx=5, pady=3)
        
        lbl4 = Ctk.CTkLabel(Frame1, text= "Cedula", width=20)
        lbl4.grid(row=3, column=0,padx=5, pady=3,sticky="e")
        self.entCedula = Ctk.CTkEntry(Frame1,border_color="teal", textvariable=Cedula)
        self.entCedula.grid(row=3,column=1, padx=5, pady=3)
        self.entCedula.bind("<Return>", lambda event: BuscarCedula(1))
        
        lbl5 = Ctk.CTkLabel(Frame1, text= "Ciudad Expedicion", width=20)
        lbl5.grid(row=4, column=0,padx=5, pady=3,sticky="e")
        self.entCiudadExpedicion= Ctk.CTkEntry(Frame1,border_color="teal", textvariable=CiudadExpedicion)
        self.entCiudadExpedicion.grid(row=4,column=1, padx=5, pady=3)

        lbl6 = Ctk.CTkLabel(Frame1, text= "Dirección", width=20)
        lbl6.grid(row=5, column=0,padx=5, pady=3,sticky="e")
        self.entDireccion = Ctk.CTkEntry(Frame1,border_color="teal", textvariable=Direccion)
        self.entDireccion.grid(row=5,column=1, padx=5, pady=3)
        
        lbl7 = Ctk.CTkLabel(Frame1, text= "Contrato", width=20)
        lbl7.grid(row=6, column=0,padx=5, pady=3,sticky="e")
        self.entContrato = Ctk.CTkEntry(Frame1,border_color="teal", textvariable=Contrato)
        self.entContrato.grid(row=6,column=1, padx=5, pady=3)
        self.entContrato.bind("<Return>", lambda event: BuscarContrato(1))
        
        lbl8= Ctk.CTkLabel(Frame1, text= "ARL", width=20)
        lbl8.grid(row=7, column=0,padx=5, pady=3,sticky="e")
        self.entARL = Ctk.CTkEntry(Frame1,border_color="teal",textvariable=ARL)
        self.entARL.grid(row=7,column=1, padx=5, pady=3)
        
        lbl9 = Ctk.CTkLabel(Frame1, text= "EPS", width=20)
        lbl9.grid(row=8, column=0,padx=5, pady=3,sticky="e")
        self.entEPS = Ctk.CTkEntry(Frame1,border_color="teal", textvariable=EPS)
        self.entEPS.grid(row=8,column=1, padx=5, pady=3)

        self.entFechaNacimiento = Ctk.CTkEntry(Frame1,border_color="teal", textvariable=FechaNacimiento)
        self.entFechaNacimiento.grid(row=9,column=1, padx=4, pady=2)
        lbl10 = Ctk.CTkLabel(Frame1, text= "Fecha de Nacimiento", width=20)
        lbl10.grid(row=9, column=0,padx=5, pady=3,sticky="e")
        lbl10 = DateEntry(Frame1, date_pattern="dd/MM/yyyy", textvariable=FechaNacimiento, width=17)
        lbl10.grid(row=9,column=1, padx=5, pady=3)
        
        # lbl9 = Label(Frame1, text= "Género", width=20)
        # lbl9.grid(row=0, column=2,padx=5, pady=3)

        #@self.entgenero = Ctk.CTkEntry(Frame1, textvariable=genero)
        #self.entgenero.grid(row=0,column=3, padx=4, pady=2)
        self.entgenero = Entry(Frame1, textvariable=genero)
        self.entgenero.grid(row=0,column=3, padx=4, pady=2)
        lbl11= Ctk.CTkLabel(Frame1, text= "Género", width=20)
        lbl11.grid(row=0, column=2,padx=5, pady=3,sticky="e")
        
        comGenero=Ctk.CTkComboBox(Frame1,border_color="teal",dropdown_hover_color="teal",variable=genero,justify="right",values=("Masculino","Femenino","Otros"))  
        comGenero.place( x=395, y = 1)
       
        

        #self.entRh = Ctk.CTkEntry(Frame1, textvariable=Rh)
        #self.entRh.grid(row=1,column=3, padx=5, pady=3)
        self.entRh = Entry(Frame1, textvariable=Rh)
        self.entRh.grid(row=1,column=3, padx=5, pady=3)
        lbl12 = Ctk.CTkLabel(Frame1, text= "RH", width=20)
        lbl12.grid(row=1, column=2,padx=5, pady=3,sticky="e")
        
        comRH =Ctk.CTkComboBox(Frame1,border_color="teal",dropdown_hover_color="teal",variable=Rh,justify="right",values=("O+","O-","A+","A-","AB+","AB-"))  
        comRH.place( x=395, y = 32)
        #lbl12.config(width=17, height=5)
        #lbl12["values"]=("O+","O-","A+","A-","AB+","AB-")

        lbl13 = Ctk.CTkLabel(Frame1, text= "Celular", width=20)
        lbl13.grid(row=2, column=2,padx=5, pady=3,sticky="e")
        self.entCelular = Ctk.CTkEntry(Frame1,border_color="teal", textvariable=Celular)
        self.entCelular.grid(row=2,column=3, padx=5, pady=3)

        lbl14 = Ctk.CTkLabel(Frame1, text= "Telefono", width=20)
        lbl14.grid(row=3, column=2,padx=5, pady=3,sticky="e")
        self.entTelefono = Ctk.CTkEntry(Frame1,border_color="teal", textvariable=Telefono)
        self.entTelefono.grid(row=3,column=3, padx=5, pady=3)

        lbl15 = Ctk.CTkLabel(Frame1, text= "Correo", width=20)
        lbl15.grid(row=4, column=2,padx=5, pady=3,sticky="e")
        self.entCorreo = Ctk.CTkEntry(Frame1,border_color="teal", textvariable=Correo)
        self.entCorreo.grid(row=4,column=3, padx=5, pady=3)
        self.entCorreo.bind("<Return>", lambda event: BuscarCorreo(1))

        lbl16 = Ctk.CTkLabel(Frame1, text= "Cargo", width=20)
        lbl16.grid(row=5, column=2,padx=5, pady=3,sticky="e")
        self.entCargo = Ctk.CTkEntry(Frame1,border_color="teal", textvariable=Cargo)
        self.entCargo.grid(row=5,column=3, padx=5, pady=3)

        # lbl15 = Label(Frame1, text= "Fecha de ID (C.C)", width=20)
        # lbl15.grid(row=6, column=2,padx=5, pady=3)
        
        self.entFechaDocumento = Ctk.CTkEntry(Frame1,border_color="teal", textvariable=FechaDocumento)
        self.entFechaDocumento.grid(row=6,column=3, padx=5, pady=3)
        lbl17 = Ctk.CTkLabel(Frame1, text= "Fecha de ID (C.C)", width=20)
        lbl17.grid(row=6, column=2,padx=5, pady=3,sticky="e")
        lbl17= DateEntry(Frame1, date_pattern="dd/MM/yyyy", textvariable=FechaDocumento, width=17)
        lbl17.grid(row=6,column=3, padx=5, pady=3)
                     
        lbl18 = Ctk.CTkLabel(Frame1, text= "Dependencia", width=20)
        lbl18.grid(row=7, column=2,padx=5, pady=3,sticky="e")
        self.entDependencia = Ctk.CTkEntry(Frame1,border_color="teal", textvariable=Dependencia)
        self.entDependencia.grid(row=7,column=3, padx=5, pady=3)

        lbl19 = Ctk.CTkLabel(Frame1, text= "Persona a Cargo", width=20)
        lbl19.grid(row=0, column=4,padx=5, pady=3,sticky="e")
        self.entJefe = Ctk.CTkEntry(Frame1,border_color="teal", textvariable=Jefe)
        self.entJefe.grid(row=0,column=5, padx=5, pady=3)

        lbl20 = Ctk.CTkLabel(Frame1, text= "Descripción del contrato", width=20)
        lbl20.grid(row=1, column=4,padx=5, pady=3,sticky="e")
        self.entDescripcionContrato = Ctk.CTkEntry(Frame1,border_color="teal", textvariable=DescripcionContrato)
        self.entDescripcionContrato.grid(row=1,column=5, padx=5, pady=3)

        lbl21 = Ctk.CTkLabel(Frame1, text= "Valor del contrato", width=20)
        lbl21.grid(row=2, column=4,padx=5, pady=3,sticky="e")
        self.entValorContrato = Ctk.CTkEntry(Frame1,border_color="teal", textvariable=ValorContrato)
        self.entValorContrato.grid(row=2,column=5, padx=5, pady=3)

        lbl22 = Ctk.CTkLabel(Frame1, text= "Autorización del contrato", width=20)
        lbl22.grid(row=3, column=4,padx=5, pady=3,sticky="e")
        self.entAutorizacionContrato = Ctk.CTkEntry(Frame1,border_color="teal", textvariable=AutorizacionContrato)
        self.entAutorizacionContrato.grid(row=3,column=5, padx=5, pady=3)

        lbl23 = Ctk.CTkLabel(Frame1, text= "PAA", width=20)
        lbl23.grid(row=4, column=4,padx=5, pady=3,sticky="e")
        self.entPAA = Ctk.CTkEntry(Frame1,border_color="teal", textvariable=PAA)
        self.entPAA.grid(row=4,column=5, padx=5, pady=3)

        #lbl22 = Ctk.CTkLabel(Frame1, text= "Banco", width=20)
        # lbl22.grid(row=5, column=4,padx=5, pady=3)
       
        
        
        lbl24 = Ctk.CTkLabel(Frame1, text= "Banco", width=20)
        lbl24.grid(row=5, column=4,padx=5, pady=3,sticky="e")
        comBanco =Ctk.CTkComboBox(Frame1,border_color="teal",variable=Banco,dropdown_hover_color="teal",justify="right",values=("Bancolombia","AV Villas","","Davivienda","Servibanca","Banco de Bogotá","banco FaCtk.CTkLabella","Nequi","Daviplata","LuloBank","","Banco Agrario","Otro"))  
        comBanco.place( x=695, y = 179)
        self.entBanco = Ctk.CTkEntry(Frame1, textvariable=Banco)
        #self.entBanco.grid(row=5,column=5, padx=5, pady=3)
        #comBanco.configure(width=17, height=5)
        

        lbl25 = Ctk.CTkLabel(Frame1, text= "Tipo de cuenta", width=20)
        lbl25.grid(row=6, column=4,padx=5, pady=3,sticky="e")
        self.entTipoCuenta = Ctk.CTkEntry(Frame1,border_color="teal", textvariable=TipoCuenta)
        self.entTipoCuenta.grid(row=6,column=5, padx=5, pady=3)

        lbl26 = Ctk.CTkLabel(Frame1, text= "Numero de cuenta", width=20)
        lbl26.grid(row=7, column=4,padx=5, pady=3,sticky="e")
        self.entNumeroCuenta = Ctk.CTkEntry(Frame1,border_color="teal", textvariable=NumeroCuenta)
        self.entNumeroCuenta.grid(row=7,column=5, padx=5, pady=3)

        lbl27 = Ctk.CTkLabel(Frame1, text= "CDP", width=20)
        lbl27.grid(row=0, column=6,padx=5, pady=3,sticky="e")
        self.entCDP = Ctk.CTkEntry(Frame1,border_color="teal", textvariable=CDP)
        self.entCDP.grid(row=0,column=7, padx=5, pady=3)

        # lbl23 = Ctk.CTkLabel(Frame1, text= "FechaCDP", width=20)
        # lbl23.grid(row=1, column=6,padx=5, pady=3)
        
        self.entFechaCDP = Ctk.CTkEntry(Frame1,border_color="teal", textvariable=FechaCDP)
        self.entFechaCDP.grid(row=1,column=7, padx=5, pady=3)
        lbl28 = Ctk.CTkLabel(Frame1, text= "Fecha del CDP", width=20)
        lbl28.grid(row=1, column=6,padx=5, pady=3,sticky="e")
        lbl28= DateEntry(Frame1, date_pattern="dd/MM/yyyy", textvariable=FechaCDP, width=17)
        lbl28.grid(row=1,column=7, padx=5, pady=3)

        lbl29 = Ctk.CTkLabel(Frame1, text= "CRP", width=20)
        lbl29.grid(row=2, column=6,padx=5, pady=3,sticky="e")
        self.entCRP = Ctk.CTkEntry(Frame1,border_color="teal", textvariable=CRP)
        self.entCRP.grid(row=2,column=7, padx=5, pady=3)





#botones
        btn1 =Ctk.CTkButton(Frame1,fg_color="teal" ,text="Agregar", width=12, height=2, command= Agregar)
        btn1.grid(row=11, column=0, padx=10, pady=10)
        btn1.place(rely=0.86,relx=0.27,relheight=0.11,relwidth=0.10)

        btn2 =Ctk.CTkButton(Frame1,fg_color="teal" , text="Eliminar", width=12, height=2, command=Eliminar)
        btn2.grid(row=11, column=1, padx=10, pady=10)
        btn2.place(rely=0.86,relx=0.75,relheight=0.11,relwidth=0.10)
        
        btn3 =Ctk.CTkButton(Frame1,fg_color="teal" , text="Actualizar", width=12, height=2,command= Actualizar)
        btn3.grid(row=11, column=2, padx=10, pady=10)
        btn3.place(rely=0.86,relx=0.39,relheight=0.11,relwidth=0.10)

        btn4 =Ctk.CTkButton(Frame1,fg_color="teal" , text="Monitor", width=12, height=2, command= Mostrar)
        btn4.grid(row=11, column=3, padx=10, pady=10)
        btn4.place(rely=0.86,relx=0.51,relheight=0.11,relwidth=0.10)
        
        btn5 =Ctk.CTkButton(Frame1,fg_color="teal" , text="Limpiar", width=12, height=2, command=Limpiar)
        btn5.grid(row=11, column=4, padx=10, pady=10)
        btn5.place(rely=0.86,relx=0.63,relheight=0.11,relwidth=0.10)
        
        btn6 =Ctk.CTkButton(Frame1,fg_color="teal", text="expide tu certificado", width=20, height=2, command=crearpdf)
        btn6.grid(row=11, column=5, padx=1, pady=10)
        btn6.place(rely=0.86,relx=0.75,relheight=0.11,relwidth=0.10)
        
        btn7 = Ctk.CTkButton(Frame1,fg_color="teal", text="Importar Excel", width=20, height=2, command=importarExcel)
        btn7.grid(row=11, column=6, padx=1, pady=10)
        btn7.place(rely=0.86,relx=0.87,relheight=0.11,relwidth=0.10)
        
        treeScroll = ttk.Scrollbar(Frame2,orient="vertical")
        treeScroll.pack(side="right",fill="y")
         
        treeScrollx = ttk.Scrollbar(Frame2,orient="horizontal")
        treeScrollx.pack(side= "bottom",fill="x")       
       
        
#ubicacion de columnas de informacion
        self.trv = ttk.Treeview(Frame2, columns=(1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29), show="headings", height="18 ",yscrollcommand=treeScroll.set,xscrollcommand=treeScrollx.set)
        self.trv.pack()

        self.trv.heading(1, text="ID")
        self.trv.heading(2, text="Nombre")
        self.trv.heading(3, text="Apellido")
        self.trv.heading(4, text="Cedula")
        self.trv.heading(5, text="CiudadExpedicion")
        self.trv.heading(6, text="Direccion")
        self.trv.heading(7, text="Contrato")
        self.trv.heading(8, text="ARL")
        self.trv.heading(9, text="EPS")
        self.trv.heading(10, text="Fecha de nacimiento")
        self.trv.heading(11, text="Género")
        self.trv.heading(12, text="RH")
        self.trv.heading(13, text="Celular")
        self.trv.heading(14, text="Telefono")
        self.trv.heading(15, text="Correo")
        self.trv.heading(16, text="Cargo")
        self.trv.heading(17, text="Fecha de Documento")
        self.trv.heading(18, text="Dependencia")
        self.trv.heading(19, text="Jefe")
        self.trv.heading(20, text="DescripcionContrato")
        self.trv.heading(21, text="ValorContrato")
        self.trv.heading(22, text="AutorizacionContrato")
        self.trv.heading(23, text="PAA")
        self.trv.heading(24, text="Banco")
        self.trv.heading(25, text="TipoCuenta")
        self.trv.heading(26, text="NumeroCuenta")
        self.trv.heading(27, text="CDP")
        self.trv.heading(28, text="FechaCDP")
        self.trv.heading(29, text="CRP")

        self.trv.bind("<ButtonRelease-1>", traineeInfo)
        cursores=["arrow","circle","clock","cross","dotbox","exchange","fleur","heart","heart","man","mouse","pirate","plus","shuttle","sizing","spider","spraycan","star","target","tcross","trek","watch"]

        treeScroll.config(command= self.trv.yview)
        treeScrollx.config(command=self.trv.xview,cursor="arrow")
        cargarServer()
        Mostrar()

if __name__ == '__main__':

    root = Ctk.CTk()
    root.wm_attributes("-topmost", 0)
    # Cargar el archivo de imagen desde el disco
    #icono = tk.PhotoImage(file="./img/libros.png",format="png")

    # Establecerlo como ícono de la ventana
    #root.iconphoto(True, icono)
    root.minsize(width=1200,height=600)
    Contratista =Contratista(root)
    root.attributes('-fullscreen')
    root.mainloop()