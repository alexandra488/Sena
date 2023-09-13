import tkinter
import customtkinter as ctk
from tkinter import*
from tkinter import ttk
import pymysql
from tkinter import messagebox
from Programa import Contratista
from Programa import *
ctk.set_appearance_mode("System")  # Modes: system (default), light, dark
ctk.set_default_color_theme("blue")  # Themes: blue (default), dark-blue, green
#si desea colocar icono descomentar numeral
#pantalla.iconbimap("nombre icono")
global pantalla1
pantalla1=ctk.CTk()

server=StringVar()
UserServer=StringVar()
passServer=StringVar()
baseServer=StringVar()         
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
cargarServer()                        
#pantalla.protocol('WM_DELETE_WINDOW', SalirApp) 



   
    
    
    
def abrir_ventana_principal():
        
    global Contratista
          
   
    root = Toplevel(pantalla1)
    Contratista =Contratista(root)
    pantalla1.withdraw()
    
    
    
    #pantalla2.destroy()
    #pantalla1.destroy()
      
     
def inicio_sesion():
    
    
    pantalla1.geometry("290x400")
    pantalla1.title("Inicio De Sesion")
    pantalla1.resizable(False,0)
    pantalla1.wm_attributes("-topmost", 0)
    def Salir():  
        
        pantalla1.quit()
    pantalla1.protocol('WM_DELETE_WINDOW', Salir)
    Label(pantalla1, text="Por favor ingrese su \n usuario y Contraseña",bg="Green",fg="white",width="300",height="3",font=("calibri",15)).pack()
    
    Label(pantalla1, text="").pack()
    
    global nombreusuario_verify
    global contraseñausuario_verify
    
    nombreusuario_verify=StringVar()
    contraseñausuario_verify=StringVar()
    
    global nombre_usuario_entry
    global contraseña_usuario_entry
    
    Label(pantalla1, text="usuario").pack()
    nombre_usuario_entry =ctk.CTkEntry(pantalla1, textvariable=nombreusuario_verify)
    nombre_usuario_entry.pack() 
    Label(pantalla1).pack()    
    Label(pantalla1, text="contraseña").pack()
    contraseña_usuario_entry = ctk.CTkEntry(pantalla1, show="*", textvariable=contraseñausuario_verify)
    contraseña_usuario_entry.pack() 
    Label(pantalla1).pack()    
    ctk.CTkButton(pantalla1,width=140,height=45,fg_color="teal" ,text="Iniciar Sesion",  command=validacion_datos).pack(pady=15)
    pantalla1.mainloop()
    
    

        
        
def validacion_datos():
    if nombreusuario_verify.get()=="" or contraseñausuario_verify.get()=="":
        messagebox.showwarning("Error ","El usuario ni la contraseña pueden estar Vacios ")
    else:
        
        bd = pymysql.connect(host=server.get(),user=UserServer.get(), password=passServer.get(), database=baseServer.get())
        
        fcursor=bd.cursor()

        fcursor.execute("SELECT contraseña FROM login WhERE usuario='"+nombreusuario_verify.get()+"' and contraseña='"+contraseñausuario_verify.get()+"'")
        
        if fcursor.fetchall():
            print(fcursor)
            #messagebox.showinfo(title="Inicio de Sesion correcto", message="Usuario y contraseña correcta")
            bd.close()
            
            abrir_ventana_principal()
        
        
        
    
        else:
            messagebox.showinfo(title="Inicio de Sesion incorrecto",message="usurario y contraseña incorrecta")
            
            bd.close()
   
            
inicio_sesion()


    

