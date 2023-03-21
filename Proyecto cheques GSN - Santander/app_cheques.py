import tkinter as tk
from tkinter import messagebox
from tkinter.ttk import *
from datetime import datetime
import sqlite3
import matplotlib.pyplot as plt
from matplotlib.lines import Line2D
import matplotlib
from PIL import ImageTk, Image
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from matplotlib.figure import Figure
import numpy as np
import openpyxl



class Aplicacion():
    
    db_name = "cheques.db"
    
    def run_query(self,query,parameters=()):
        with sqlite3.connect(self.db_name) as conn:
            cursor = conn.cursor()
            result = cursor.execute(query,parameters)
            conn.commit()
        return result
    
    def obtener_periodos(self):
        #Elimino los elementos para actualizar
        for elements in self.listaSeleccion.get_children():
            self.listaSeleccion.delete(elements)
        
        #Me genero una variable acumuladora
        saldo_acumulado = 0
   
        query = "SELECT saldo, fecha FROM saldo WHERE id = (SELECT MAX(id) FROM saldo)"
        saldo_inicio = self.run_query(query)
        for i in saldo_inicio:
            saldo_acumulado = float(i[0])
            
        
        query = "SELECT SUM(importe) FROM cheques WHERE (fecha_v <= date('now','localtime')) AND pendiente = 1 "
        importe = self.run_query(query)
        for imp in importe:
            if str(imp[0])==str("NONE") or str(imp[0])==str("None"):
                self.listaSeleccion.insert("","0",text="Cheques pendientes de débito/crédito",values=[0,round(saldo_acumulado,2)])
            else:
                saldo_acumulado = saldo_acumulado + float(imp[0])
                self.listaSeleccion.insert("","0",text="Cheques pendientes de débito/crédito",values=[round(float(imp[0]),2),round(saldo_acumulado,2)])
            
        query = "SELECT SUM(importe),strftime('%m-%Y',fecha_v) FROM cheques WHERE (fecha_v > date('now','localtime')) AND pendiente = 1 GROUP BY strftime('%m-%Y',fecha_v) ORDER BY fecha_v ASC"
        #query = "SELECT DISTINCT strftime('%m-%Y',fecha_v) FROM cheques WHERE fecha_v > date('now','localtime') ORDER BY fecha_v ASC"
        periodos = self.run_query(query)
        

        
        for a in periodos:
            saldo_acumulado = saldo_acumulado + float(a[0])  
            self.listaSeleccion.insert("", "end", text=a[1],values=[round(float(a[0]),2),round(saldo_acumulado,2)])  

    #Nueva funcionalidad listado de cheques en Excel
    def exportar_periodo_excel(self):
        query = "SELECT cheques.id,cheques.numero, strftime('%d/%m/%Y',cheques.fecha_e), strftime('%d/%m/%Y',cheques.fecha_v), cheques.importe, entidades.entidad, cheques.echeq, cheques.estado FROM cheques INNER JOIN entidades ON entidades.id_entidad = cheques.id_entidad WHERE (cheques.pendiente = 1) ORDER BY fecha_v" #GROUP BY strftime('%m-%Y',fecha_v)
        
        arreglo_datos = []

        cheques = self.run_query(query)
        for c in cheques:
            if int(c[6]) == 1:
                if int(c[7]) == 1:
                    arreglo_datos.append([c[0],c[2],c[1],c[5],c[3],"ECHEQ","EN CURSO",c[4]])
                else:
                    arreglo_datos.append([c[0],c[2],c[1],c[5],c[3],"ECHEQ","",c[4]])
            else:
                if int(c[7]) == 1:
                    arreglo_datos.append([c[0],c[2],c[1],c[5],c[3],"","EN CURSO",c[4]])
                else:
                    arreglo_datos.append([c[0],c[2],c[1],c[5],c[3],"","",c[4]])
                    
        
        #convertir a excel
        try:
            wb = openpyxl.Workbook() #Genero un libro de trabajo
            hoja = wb.active
            hoja.append(["ID", "F. Emisión", "Nro Cheque", "Entidad", "F. Vto", "Tipo", "Estado", "Importe","Saldo Acumulado"])

            #ultimo saldo
            query = "SELECT saldo, fecha FROM saldo WHERE id = (SELECT MAX(id) FROM saldo)"
            ult_saldo = self.run_query(query)
            for s in ult_saldo:
                fecha = datetime.strptime(s[1], "%Y-%m-%d %H:%M:%S")
                fecha = fecha.strftime("%d/%m/%Y")
                hoja.append(["",str(fecha),"","Saldo Bancario",str(fecha),"","","",float(s[0])])

            contador = 1
            for fila in arreglo_datos:
                hoja.append(fila)
                acum = float(hoja["I{}".format(contador+1)].value)+float(fila[7])
                hoja.cell(row=contador+2, column=9, value=acum)
                contador = contador + 1
            wb.save('cheques.xlsx')
            messagebox.showinfo(message="Archivo .xlsx creado con exito")
            
        except:
            messagebox.showerror(title="Error",message="Se produjo un error en el proceso. Verifica que el archivo no se encuentre abierto.")


    ####################################################

    def mostrar_detalle(self):
        for elements in self.listaDetalle.get_children():
            self.listaDetalle.delete(elements)
        try:
            item = self.listaSeleccion.item(self.listaSeleccion.selection()[0])
        except:
            for elements in self.listaDetalle.get_children():
                self.listaDetalle.delete(elements)
            return
            #item = self.listaSeleccion.item(self.listaSeleccion.get_children()[0])
        if item['text']=="Cheques pendientes de débito/crédito":
            query = "SELECT cheques.id,cheques.numero, strftime('%d/%m/%Y',cheques.fecha_e), strftime('%d/%m/%Y',cheques.fecha_v), cheques.importe, entidades.entidad, cheques.echeq, cheques.estado FROM cheques INNER JOIN entidades ON entidades.id_entidad = cheques.id_entidad WHERE (cheques.pendiente = 1) AND (fecha_v<=date('now','localtime')) AND cheques.importe>0 ORDER BY fecha_v".format(item['text']) 
            query2 = "SELECT cheques.id,cheques.numero, strftime('%d/%m/%Y',cheques.fecha_e), strftime('%d/%m/%Y',cheques.fecha_v), cheques.importe, entidades.entidad, cheques.echeq, cheques.estado FROM cheques INNER JOIN entidades ON entidades.id_entidad = cheques.id_entidad WHERE (cheques.pendiente = 1) AND (fecha_v<=date('now','localtime')) AND cheques.importe<0 ORDER BY fecha_v".format(item['text'])
        else:
            query = "SELECT cheques.id,cheques.numero, strftime('%d/%m/%Y',cheques.fecha_e), strftime('%d/%m/%Y',cheques.fecha_v), cheques.importe, entidades.entidad, cheques.echeq, cheques.estado FROM cheques INNER JOIN entidades ON entidades.id_entidad = cheques.id_entidad WHERE (cheques.pendiente = 1) AND (strftime('%m-%Y',fecha_v)='{}') AND cheques.importe>0 AND (fecha_v>date('now','localtime')) ORDER BY fecha_v".format(item['text']) 
            query2 = "SELECT cheques.id,cheques.numero, strftime('%d/%m/%Y',cheques.fecha_e), strftime('%d/%m/%Y',cheques.fecha_v), cheques.importe, entidades.entidad, cheques.echeq, cheques.estado FROM cheques INNER JOIN entidades ON entidades.id_entidad = cheques.id_entidad WHERE (cheques.pendiente = 1) AND (strftime('%m-%Y',fecha_v)='{}') AND cheques.importe<0 AND (fecha_v>date('now','localtime')) ORDER BY fecha_v".format(item['text']) 
        cheques = self.run_query(query)
        for c in cheques:
            if int(c[6]) == 1:
                if int(c[7]) == 1:
                    self.listaDetalle.insert("","end",text=c[0],values=[c[2],c[1],c[5],c[3],"ECHEQ","EN CURSO",c[4]],tags=("Pos","ECH"))
                else:
                    self.listaDetalle.insert("","end",text=c[0],values=[c[2],c[1],c[5],c[3],"ECHEQ","",c[4]],tags=("Pos","ECH"))
            else:
                if int(c[7]) == 1:
                    self.listaDetalle.insert("","end",text=c[0],values=[c[2],c[1],c[5],c[3],"","EN CURSO",c[4]],tags=("Pos",))
                else:
                    self.listaDetalle.insert("","end",text=c[0],values=[c[2],c[1],c[5],c[3],"","",c[4]],tags=("Pos",))
        cheques = self.run_query(query2)
        for c in cheques:
            if int(c[6]) == 1:
                if int(c[7]) == 1:
                    self.listaDetalle.insert("","end",text=c[0],values=[c[2],c[1],c[5],c[3],"ECHEQ","EN CURSO",c[4]],tags=("Neg","ECH"))
                else:
                    self.listaDetalle.insert("","end",text=c[0],values=[c[2],c[1],c[5],c[3],"ECHEQ","",c[4]],tags=("Neg","ECH"))
            else:
                if int(c[7]) == 1:
                    self.listaDetalle.insert("","end",text=c[0],values=[c[2],c[1],c[5],c[3],"","EN CURSO",c[4]],tags=("Neg",))
                else:
                    self.listaDetalle.insert("","end",text=c[0],values=[c[2],c[1],c[5],c[3],"","",c[4]],tags=("Neg",))
                
    def mostrar_ultimo_saldo(self):
        query = "SELECT saldo, fecha FROM saldo WHERE id = (SELECT MAX(id) FROM saldo)"
        datos = self.run_query(query)
        for i in datos:
            self.e_saldo['state']="normal"
            self.e_fecha_saldo['state'] = "normal"
            self.e_saldo.delete(0,"end")
            self.e_fecha_saldo.delete(0,"end")
            
            fecha = datetime.strptime(i[1], "%Y-%m-%d %H:%M:%S")
            fecha = fecha.strftime("%H:%M %d/%m/%Y")
            self.e_saldo.insert(0,str(i[0]))
            self.e_fecha_saldo.insert(0,str(fecha))
            
            self.e_saldo['state']="readonly"
            self.e_fecha_saldo['state']="readonly"
            
            
    def activar_cheque(self):
        if self.listaDetalle.selection():
            item = self.listaDetalle.item(self.listaDetalle.selection()[0])
            if datetime.strptime(str(item['values'][3]), '%d/%m/%Y')>datetime.now():
                messagebox.showerror(title="Error",message="No se puede depositar el cheque n° "+str(item['values'][1])+ ".\n La fecha de vencimiento es el "+str(item['values'][3]))
                return
            if str(item['values'][5]) != "EN CURSO":
                messagebox.showerror(title="Error",message="El cheque n° "+str(item['values'][1])+" no se encuentra en curso.\nPara que un cheque sea depositado, primero debe marcarse como 'EN CURSO'")
                return
            self.inhabilitar_botones()
            op = messagebox.askyesno(title="Confirmación",message="¿Desea depositar el cheque nro "+str(item['values'][1])+" - importe = $"+str(item['values'][6]))
            if op == True:
                query = "UPDATE cheques SET pendiente = 0 WHERE id = {}".format(int(item['text']))
                self.run_query(query)
                self.mostrar_detalle()
                self.obtener_periodos()
                messagebox.showinfo(message="El cheque fue depositado y ya no figurará en la vista 'Detalle'.\nRecordá actualizar el saldo a la brevedad para reflejar la actualización del mismo.")
            self.habilitar_botones()
        else:
            messagebox.showerror(title="Error",message="Primero debes seleccionar el cheque a depositar")
    
    
            
    
    def eliminar_cheque(self):

        if self.listaDetalle.selection():
            op=messagebox.askyesno("Eliminar cheque",message="¿Deseas eliminar el registro?")
            if op== True:
                id_ = self.listaDetalle.item(self.listaDetalle.selection()[0])
                query = "DELETE FROM cheques WHERE id={}".format(id_['text'])
                self.run_query(query)

                self.mostrar_detalle()
                self.obtener_periodos()
        else:
            messagebox.showerror("Error",message="Primero seleccioná el cheque que deseas eliminar de la lista detalle")
            
    def graficar_saldo(self):
        self.inhabilitar_botones()
        etiquetas = []
        saldo = []
        
        for i in self.listaSeleccion.get_children():
            item = self.listaSeleccion.item(i)
            etiquetas.append(str(item['text']))
            saldo.append(float(item['values'][0]))
        
        hoy = datetime.now()
        
        etiquetas[0] = "Hoy\n{}-{}-{}".format(hoy.day,hoy.month,hoy.year)
        
        self.grafico(absisas=etiquetas, ordenadas=saldo,t="m")
        
        
    def graficar(self,absisas,ordenadas,t="a"):
        def add_value_label(x_list,y_list,y_label_list):
            for i in range(1, len(x_list)+1):
                plt.text(i-1,y_list[i-1]/2,y_label_list[i-1], ha="center",weight='bold')
        fix, ax = plt.subplots()
        if t=="m":
            ax.set_title("Gráfico de Saldos Mensuales",weight="bold")
            fix.canvas.manager.set_window_title("Gráfico de resúmen [Saldo Mensual]")
        else:
            ax.set_title("Gráfico de Saldo Acumulado",weight="bold")
            fix.canvas.manager.set_window_title("Gráfico de resúmen [Saldo Acumulado]")
        ax.set_xlabel("Mes",weight="bold",color="b",labelpad=2)
        ax.set_ylabel("Saldo ($)",weight="bold",color="b",labelpad=2)
        plt.bar(absisas, ordenadas, color=['tomato' if s < 0 else 'springgreen' for s in ordenadas])

        plt.grid()
        ax.get_yaxis().set_major_formatter(matplotlib.ticker.FuncFormatter(lambda x, p: format(int(x), ',')))
        
        legend_handles = [Line2D([0], [0], linewidth=0, marker='o', markerfacecolor=color, markersize=12, markeredgecolor='none')
                  for color in ['springgreen', 'tomato']]
        ax.legend(legend_handles, ['Saldo positivo', 'Saldo negativo'])
        
        valores_formato = []
        for i in ordenadas:
            valores_formato.append('{:,.2f}'.format(i).replace(",", "@").replace(".", ",").replace("@", "."))
        add_value_label(absisas,ordenadas,valores_formato)
        
        def on_closing(evt):
            self.habilitar_botones()
        
        
        fix.canvas.mpl_connect('close_event', on_closing)
        plt.show()
    
    def grafico(self,absisas,ordenadas,t="a"):
        def add_value_label(x_list,y_list,y_label_list):
            for i in range(1, len(x_list)+1):
                plt.text(i-1,y_list[i-1]/2,y_label_list[i-1], ha="center",weight='bold',fontsize=9)
        
        ventana_grafica = tk.Toplevel()
        ventana_grafica.focus_force()
        ventana_grafica.iconbitmap("icono_genericos.ico")
        
        labelpos = np.arange(len(absisas))
        
        fig = plt.figure(figsize=(6,5),dpi=100)
        plt.bar(labelpos, ordenadas, align= "center", alpha=1.0 ,color=['tomato' if s < 0 else 'springgreen' for s in ordenadas])
        plt.xticks(labelpos,absisas)
        plt.ylabel("Saldo",fontweight='bold')
        plt.xlabel("Período",fontweight='bold')
        plt.tight_layout(pad=2.2,w_pad=0.5,h_pad=0.1)
        plt.grid()
        if t=="a":
            plt.title("Saldo Acumulado por período",fontweight='bold')
        else:
            plt.title("Saldo Mensual",fontweight='bold')
        plt.xticks(rotation=30,horizontalalignment="center")
        
        valores_formato = []
        for i in ordenadas:
            valores_formato.append('{:,.2f}'.format(i).replace(",", "@").replace(".", ",").replace("@", "."))
        add_value_label(absisas,ordenadas,valores_formato)
                
        
        canvas = FigureCanvasTkAgg(fig, master= ventana_grafica)
        canvas.draw()
        canvas.get_tk_widget().pack(side=tk.TOP,fill=tk.BOTH,expand=1)
        
        toolbar = NavigationToolbar2Tk(canvas, ventana_grafica)       
        toolbar.update()
        canvas.get_tk_widget().pack(side=tk.TOP,fill=tk.BOTH,expand=1)
        
        def on_closing():
            self.habilitar_botones()
            ventana_grafica.destroy()
            
        ventana_grafica.protocol("WM_DELETE_WINDOW", on_closing)
        
        self.center(ventana_grafica)
        
        
        # fix, ax = plt.subplots()
        # if t=="m":
        #     ax.set_title("Gráfico de Saldos Mensuales",weight="bold")
        #     fix.canvas.manager.set_window_title("Gráfico de resúmen [Saldo Mensual]")
        # else:
        #     ax.set_title("Gráfico de Saldo Acumulado",weight="bold")
        #     fix.canvas.manager.set_window_title("Gráfico de resúmen [Saldo Acumulado]")
        # ax.set_xlabel("Mes",weight="bold",color="b",labelpad=2)
        # ax.set_ylabel("Saldo ($)",weight="bold",color="b",labelpad=2)
        

        # plt.grid()
        # ax.get_yaxis().set_major_formatter(matplotlib.ticker.FuncFormatter(lambda x, p: format(int(x), ',')))
        
        # legend_handles = [Line2D([0], [0], linewidth=0, marker='o', markerfacecolor=color, markersize=12, markeredgecolor='none')
        #           for color in ['springgreen', 'tomato']]
        # ax.legend(legend_handles, ['Saldo positivo', 'Saldo negativo'])
        

        

        
        
        
        
        
    def graficar_saldo_acum(self):
        self.inhabilitar_botones()
        etiquetas = []
        saldo_a= []
        
        for i in self.listaSeleccion.get_children():
            item = self.listaSeleccion.item(i)
            etiquetas.append(str(item['text']))
            saldo_a.append(float(item['values'][1]))
        
        hoy = datetime.now()
        
        etiquetas[0] = "Hoy\n{}-{}-{}".format(hoy.day,hoy.month,hoy.year)
        
        self.grafico(absisas=etiquetas, ordenadas=saldo_a)
    
            
    
    
    def curso_cheque(self):
        if self.listaDetalle.selection():
            item = self.listaDetalle.item(self.listaDetalle.selection()[0])
            if datetime.strptime(str(item['values'][3]), '%d/%m/%Y')>datetime.now():
                messagebox.showerror(title="Error",message="No se puede marcar/desmarcar 'en curso' el cheque n° "+str(item['values'][1])+ ".\nLa fecha de vencimiento es el "+str(item['values'][3])+".")
                return
            self.inhabilitar_botones()
            op = messagebox.askyesno(title="Confirmación",message="¿Desea marcar/desmarcar 'en curso' al cheque nro "+str(item['values'][1])+" - importe = $"+str(item['values'][6]))
            if op == True:
                query = "SELECT estado FROM cheques WHERE id ={}".format(int(item['text']))
                estado = self.run_query(query)
                for i in estado:
                    if int(i[0])==0:
                        query = "UPDATE cheques SET estado = 1 WHERE id = {}".format(int(item['text']))
                        self.run_query(query)
                    else:
                        query = "UPDATE cheques SET estado = 0 WHERE id = {}".format(int(item['text']))
                        self.run_query(query)
                self.mostrar_detalle()
                self.obtener_periodos()
                messagebox.showinfo(message="El estado del cheque se modificó.\nSi marcaste al cheque como 'En curso' recorda que deberás depositar el cheque una vez que lo identifiques en el resumen bancario.")
            self.habilitar_botones()
        else:         
            messagebox.showerror(title="Error",message="Primero debes seleccionar el cheque al que deseas cambiar su estado.")
    
    def actualizar_ventana(self):
        self.obtener_periodos()
        self.mostrar_detalle()
        
    def emitir_listado(self):
        
        def generar_lista_cheques(base,long):
            for i in self.listaNumerada.get_children():
                self.listaNumerada.delete(i)
            try: 
                base = int(base)
            except:
                messagebox.showerror(parent=ventana_lista,title="ERROR",message="Ingresa un número de cheque valido, sin comas ni puntos.")
                return
            try:
                long = int(long)
            except:
                messagebox.showerror(parent=ventana_lista,title="ERROR",message="La longitud de serie debe ser un número entero positivo.")
                return
            if long > 1000:
                long = 1000
            lista = []
            lista.append(base)
            for i in range(0,long):
                base = base + 1
                lista.append(base)
            #print(lista)
            
            for i in lista:
                query = "SELECT cheques.id,cheques.numero,strftime('%d/%m/%Y',cheques.fecha_e),strftime('%d/%m/%Y',cheques.fecha_v), entidades.entidad, cheques.importe, cheques.pendiente, cheques.echeq, cheques.estado FROM cheques INNER JOIN entidades ON entidades.id_entidad = cheques.id_entidad WHERE cheques.numero LIKE '%"+str(i)+"%'"
                cheque = self.run_query(query)
                
                bandera = 0
                
                for ch in cheque:
                    #print("para i= "+str(i)+"los cheques encontrados son"+ch[1])
                    if i == int(ch[1]):
                        if ch[6]==0:
                            if ch[7] == 1:
                                self.listaNumerada.insert("","end",text=ch[0],values=[ch[2],ch[1],ch[4],ch[3],"ECHEQ","DEPOSITADO",ch[5]],tags=("DEPOSITADO","ECHEQ"))
                            else:
                                self.listaNumerada.insert("","end",text=ch[0],values=[ch[2],ch[1],ch[4],ch[3],"","DEPOSITADO",ch[5]],tags=("DEPOSITADO",))
                        else:
                            
                            if ch[7] == 1:
                                if ch[8] == 1:
                                    if float(ch[5])>=0:
                                        self.listaNumerada.insert("","end",text=ch[0],values=[ch[2],ch[1],ch[4],ch[3],"ECHEQ","EN CURSO",ch[5]],tags=("PENDIENTE POS","ECHEQ","EN CURSO"))
                                    else:
                                        self.listaNumerada.insert("","end",text=ch[0],values=[ch[2],ch[1],ch[4],ch[3],"ECHEQ","EN CURSO",ch[5]],tags=("PENDIENTE NEG","ECHEQ","EN CURSO"))
                                else:
                                    if float(ch[5])>=0:
                                        self.listaNumerada.insert("","end",text=ch[0],values=[ch[2],ch[1],ch[4],ch[3],"ECHEQ","",ch[5]],tags=("PENDIENTE POS","ECHEQ",))
                                    else:
                                        self.listaNumerada.insert("","end",text=ch[0],values=[ch[2],ch[1],ch[4],ch[3],"ECHEQ","",ch[5]],tags=("PENDIENTE NEG","ECHEQ",))
                            else:
                                if ch[8] == 1:
                                    if float(ch[5])>=0:
                                        self.listaNumerada.insert("","end",text=ch[0],values=[ch[2],ch[1],ch[4],ch[3],"","EN CURSO",ch[5]],tags=("PENDIENTE POS","EN CURSO"))
                                    else:
                                        self.listaNumerada.insert("","end",text=ch[0],values=[ch[2],ch[1],ch[4],ch[3],"","EN CURSO",ch[5]],tags=("PENDIENTE NEG","EN CURSO"))
                                else:
                                    if float(ch[5])>=0:
                                        self.listaNumerada.insert("","end",text=ch[0],values=[ch[2],ch[1],ch[4],ch[3],"","",ch[5]],tags=("PENDIENTE POS",))
                                    else:
                                        self.listaNumerada.insert("","end",text=ch[0],values=[ch[2],ch[1],ch[4],ch[3],"","",ch[5]],tags=("PENDIENTE NEG",))
                                
                        bandera = 1
                
                if bandera == 0:
                        self.listaNumerada.insert("","end",text="",values=["",i,"","","","",""],tags=("NE",))
                #self.listaNumerada.insert("",END,text=)
        
        ventana_lista = tk.Toplevel()
        ventana_lista.title("Listado detallado de cheques")
        ventana_lista.focus_force()
        ventana_lista.resizable(0,0)
        ventana_lista.iconbitmap("icono_genericos.ico")
        
        self.inhabilitar_botones()
        
        frame = tk.LabelFrame(ventana_lista,text="Detalle",foreground="green",font=("Arial",12,"bold"))
        frame.grid(row=0,column=0,padx=5,pady=5)
        
        tk.Label(frame,text="N° de origen de serie",font=('Arial', 10,'bold')).grid(row=0,column=0,padx=5,pady=5)
        
        e_origen = tk.Entry(frame,fg="blue",font=("Arial",10,"bold"))
        e_origen.grid(row=0,column=1,padx=5,pady=5)
        e_origen.focus_force()
        
        tk.Label(frame,text="Longitud de la serie",font=('Arial', 10,'bold')).grid(row=0,column=2,padx=5,pady=5)
        
        e_long = tk.Entry(frame,fg="blue",font=("Arial",10,"bold"))
        e_long.grid(row=0,column=3,padx=5,pady=5)
        
        boton = tk.Button(frame,command=lambda: generar_lista_cheques(e_origen.get(),e_long.get()),text="Ver Listado",width=12,bg = "greenyellow",font=("Arial",8,"bold"),activebackground="green")
        boton.grid(row=0,column=4,padx=5,pady=5)
        
        boton.bind("<Return>", lambda e: generar_lista_cheques(e_origen.get(),e_long.get()))
        
        def fixed_map(option):

            # Fix for setting text colour for Tkinter 8.6.9
            # From: https://core.tcl.tk/tk/info/509cafafae
            #
            # Returns the style map for 'option' with any styles starting with
            # ('!disabled', '!selected', ...) filtered out.
            # style.map() returns an empty list for missing options, so this
                                    
            # should be future-safe.
            return [elm for elm in style.map('Treeview', query_opt=option) if
                elm[:2] != ('!disabled', '!selected')]
        
        style = Style()
        style.theme_use("clam")
        style.configure("Treeview",background="silver",foreground="black",fieldbackground="silver", highlightthickness=0, bd=0, font=('Arial', 9,"bold")) # Modify the font of the body
        
        style.configure("Treeview.Heading", font=('Arial', 10,'bold')) # Modify the font of the headings
        style.map('Treeview', foreground=fixed_map('foreground'), background=fixed_map('background'))
        style.map("Treeview",background=[("selected","midnight blue")])
        
        frame_listado = tk.Frame(ventana_lista)
        frame_listado.grid(row=1,column=0,padx=5,pady=5)
        
        scrol = Scrollbar(frame_listado)
        scrol.grid(row=2,column=3,sticky="NSEW")
        
        self.listaNumerada = Treeview(frame_listado,yscrollcommand=scrol.set,height=10,style="Treeview",columns=["#1","#2","#3","#4","#5","#6","#7"])
        self.listaNumerada.grid(row=2,column=0,columnspan=3,sticky="NSEW")
        
        scrol.config(command=self.listaNumerada.yview)
        
        self.listaNumerada.tag_configure("DEPOSITADO", background='gray')
        self.listaNumerada.tag_configure("NE", background='red')
        self.listaNumerada.tag_configure("ECHEQ",foreground="blue")
        self.listaNumerada.tag_configure("PENDIENTE POS",background="lawn green")
        self.listaNumerada.tag_configure("PENDIENTE NEG", background= "orange")

        
        self.listaNumerada.column("#0", width=0, minwidth=0)
        self.listaNumerada.column("#1", width=75,anchor=tk.CENTER,minwidth=70)
        self.listaNumerada.column("#2", width=80,anchor=tk.CENTER,minwidth=70)
        self.listaNumerada.column("#3", width=200,anchor=tk.CENTER,minwidth=70)
        self.listaNumerada.column("#4", width=75,anchor=tk.CENTER,minwidth=70)
        self.listaNumerada.column("#5", width=60,anchor=tk.CENTER,minwidth=60)
        self.listaNumerada.column("#6", width=100,anchor=tk.CENTER, minwidth=70)
        self.listaNumerada.column("#7", width=100,anchor=tk.CENTER,minwidth=70)
        
        
        self.listaNumerada.heading("#1", text="F. Emisión")
        self.listaNumerada.heading("#2",text="Número")
        self.listaNumerada.heading("#3",text="Entidad")
        self.listaNumerada.heading("#4",text="F. Vto")
        self.listaNumerada.heading("#5",text="Tipo")
        self.listaNumerada.heading("#6",text="Estado")
        self.listaNumerada.heading("#7",text="Importe")
        
        self.listaNumerada["displaycolumns"]=["#1","#2","#3","#4","#5","#6","#7"]
        
        
        
        def on_closing():
            self.habilitar_botones()
            ventana_lista.destroy()
            
        ventana_lista.protocol("WM_DELETE_WINDOW", on_closing)
        
        self.center(ventana_lista)
        
    def __init__(self,ventana):
        self.ventana = ventana
        self.ventana.title("Cascada de Cheques - Genéricos San Nicolás SRL - © Ramiro Perez")
        self.ventana.iconbitmap("icono_genericos.ico")
        self.ventana.focus_force()
        self.ventana.resizable(0,0)
        
        
        #===================================================================
        # Botones
        
        framelogo = tk.Frame(self.ventana)
        framelogo.grid(row=0,column=0,padx=5,pady=1,sticky="NS")

        self.img = Image.open(fp=r'gen.png')
        o_size = self.img.size
        f_size = (80,80)
        
        factor = min(float(f_size[1])/o_size[1], float(f_size[0])/o_size[0])
        width = int(o_size[0] * factor)
        height = int(o_size[1] * factor)
        self.rImg= self.img.resize((width, height), Image.ANTIALIAS)
        self.rImg = ImageTk.PhotoImage(self.rImg)
        
        canvas = tk.Canvas(framelogo, width=f_size[0], height= f_size[1])
        canvas.create_image(f_size[0]/2, f_size[1]/2, anchor=tk.CENTER, image=self.rImg, tags="img")
        canvas.grid(row=0,column=0)
        
        
        framestats = tk.LabelFrame(self.ventana,text="Visual",font=("Arial",12,"bold"))
        framestats.grid(row=1,column=0,padx=5,pady=1,sticky="NSEW")
        
        frameGraf = tk.LabelFrame(framestats,text="Gráficos",font=("Arial",10,"bold"))
        frameGraf.grid(row=0,column=0,padx=5,pady=1,sticky="NS")
        
        self.botonVerGraficaSaldo = tk.Button(frameGraf,command=self.graficar_saldo,text="Graficar Saldo",width=15,bg = "lightblue",font=("Arial",8,"bold"),activebackground="blue")
        self.botonVerGraficaSaldo.grid(row=0,column=1,padx=5,pady=3,sticky="NS")
        
        self.botonVerGraficaSaldoAc = tk.Button(frameGraf,command=self.graficar_saldo_acum,text="Graficar Saldo\nAcumulado",width=15,bg = "lightblue",font=("Arial",8,"bold"),activebackground="blue")
        self.botonVerGraficaSaldoAc.grid(row=1,column=1,padx=5,pady=3,sticky="NS")
        
        self.botonExportarAExcel = tk.Button(frameGraf,command=self.exportar_periodo_excel,text="Exportar a Excel",width=15,bg = "lightblue",font=("Arial",8,"bold"),activebackground="blue")
        self.botonExportarAExcel.grid(row=2,column=1,padx=5,pady=3,sticky="NS")
        
        frameBotones = tk.LabelFrame(self.ventana,text="Opciones",font=("Arial",12,"bold"))
        frameBotones.grid(row=2,column=0,padx=5,pady=1,sticky="NS")
        
        frameBotonesCheques = tk.LabelFrame(frameBotones,text="Cheques",font=("Arial",10,"bold"))
        frameBotonesCheques.grid(row=0,column=0,padx=5,pady=1,sticky="NS")
        
        self.botonAgregarCheque = tk.Button(frameBotonesCheques,command=self.ventana_nuevo_cheque,text="Cargar Cheque",width=15,bg = "greenyellow",font=("Arial",8,"bold"),activebackground="green")
        self.botonAgregarCheque.grid(row=0,column=1,padx=5,pady=3,sticky="NS")
        
        self.botonEliminarCheque = tk.Button(frameBotonesCheques,command=self.eliminar_cheque,text="Eliminar Cheque",width=15, bg= "tomato",font=("Arial",8,"bold"),activebackground="red")
        self.botonEliminarCheque.grid(row=1,column=1,padx=5,pady=3,sticky="NS")
        
        self.botonModificarCheque = tk.Button(frameBotonesCheques,command=self.modificar_cheque,text="Modificar Cheque",width=15, bg = "gold",font=("Arial",8,"bold"),activebackground="yellow")
        self.botonModificarCheque.grid(row=2,column=1,padx=5,pady=3,sticky="NS")
        
        self.botonCursoCheque = tk.Button(frameBotonesCheques,command=self.curso_cheque,text="Marcar / Desmarcar\n'En Curso'",width=15, bg = "hotpink",font=("Arial",8,"bold"),activebackground="magenta")
        self.botonCursoCheque.grid(row=3,column=1,padx=5,pady=3,sticky="NS")
        
        self.botonDepositarCheque = tk.Button(frameBotonesCheques,command=self.activar_cheque,text="Depositar Cheque",width=15,bg = "medium purple", font=("Arial",8,"bold"),activebackground="purple")
        self.botonDepositarCheque.grid(row=4,column=1,padx=5,pady=3,sticky="NS")
        
        self.botonActuaizar = tk.Button(frameBotonesCheques,command=self.actualizar_ventana,text="Actualizar Vistas",width=15,bg = "greenyellow", font=("Arial",8,"bold"),activebackground="green")
        self.botonActuaizar.grid(row=5,column=1,padx=5,pady=3,sticky="NS")
        
        self.botonListarCheques = tk.Button(frameBotonesCheques,command=self.emitir_listado,text="Emitir listado",width=15,bg = "greenyellow", font=("Arial",8,"bold"),activebackground="green")
        self.botonListarCheques.grid(row=6,column=1,padx=5,pady=3,sticky="NS")
        
        frameBotonesEntes = tk.LabelFrame(frameBotones,text="Entidades",font=("Arial",10,"bold"))
        frameBotonesEntes.grid(row=1,column=0,padx=5,pady=1,sticky="NS")
        
        self.botonAgregarEnte = tk.Button(frameBotonesEntes,command=self.ventana_nueva_entidad,text="Entidades",width=15,bg = "lightblue",font=("Arial",8,"bold"),activebackground="blue")
        self.botonAgregarEnte.grid(row=0,column=1,padx=5,pady=3,sticky="NS")
        
        #------- Saldos --------#
        ####################################################################
        #Saldos
        frame_saldo = tk.LabelFrame(self.ventana,text="Saldo Bco. Santander",font=("Arial",12,"bold"), fg="red")
        frame_saldo.grid(row=0,column=1,padx=5,pady=1,sticky="NSEW")
        
        frame_entrys = tk.Frame(frame_saldo)
        frame_entrys.pack(side=tk.RIGHT)
        
        tk.Label(frame_entrys,text="Última actualización",font=('Arial', 10,'bold')).grid(row=0,column=0,padx=5,pady=5,sticky="E")
        
        self.e_fecha_saldo = tk.Entry(frame_entrys,font=('Arial', 11,'bold'),fg="blue",state="readonly")
        self.e_fecha_saldo.grid(row=0,column=1,padx=5,pady=3,sticky="E")
        
        tk.Label(frame_entrys,text="Saldo",font=('Arial', 10,'bold')).grid(row=1,column=0,padx=5,pady=5,sticky="E")
        
        self.e_saldo = tk.Entry(frame_entrys,font=('Arial', 11,'bold'),fg="blue",state="readonly")
        self.e_saldo.grid(row=1,column=1,padx=5,pady=[0,5],sticky="E")
        
        frame_botones = tk.Frame(frame_saldo)
        frame_botones.pack(side=tk.RIGHT)
        
        def actualizar_saldo():
            self.e_saldo['state'] = "normal"
            self.e_saldo['background'] = "spring green"
            self.inhabilitar_botones()
            self.boton_guardar['state']="normal"
        
        def guardar_saldo():
            op = messagebox.askyesno(title="Confirmación",message="¿Deseas actualizar el saldo bancario?")
            if op == True:
                try: 
                    saldo = round(float(self.e_saldo.get()),2)
                except:
                    messagebox.showerror(title="Error",message="Debés ingresar el saldo en formato numérico.\n Utiliza punto (.) en lugar de coma (,) para expresar la parte decimal.")
                
                fe_saldo = datetime.now()
                fe_saldo= fe_saldo.strftime("%Y-%m-%d %H:%M:%S")   #"%H:%M %d/%m/%Y"
                query = "INSERT INTO saldo VALUES (NULL,{},'{}')".format(saldo,fe_saldo)
                self.run_query(query)
                # self.e_fecha_saldo.delete(0,"end")
                # self.e_fecha_saldo.insert(0,str(saldo))
                # self.e_fecha_saldo['state']="normal"
                # self.e_fecha_saldo.delete(0,"end")
                # fe_saldo = datetime.strptime(fe_saldo,"%Y-%m-%d %H:%M:%S" )
                # fe_saldo = fe_saldo.strftime("%H:%M %d/%m/%Y")
                # self.e_fecha_saldo.insert(0,str(fe_saldo))
                self.habilitar_botones()
                self.boton_guardar['state']="disabled"
                self.e_saldo['background']="white"
                self.mostrar_ultimo_saldo()
                self.obtener_periodos()
                messagebox.showinfo(title="Carga Exitosa",message="El saldo bancario se ha actualizado correctamente.")
                
        
        self.boton_actualizar = tk.Button(frame_botones,command=actualizar_saldo,text="Actualizar",width=8,bg = "lightblue",font=("Arial", 8,"bold"),activebackground="blue")
        self.boton_actualizar.grid(row=0,column=0,padx=5,pady=[2,2])
        
        self.boton_guardar = tk.Button(frame_botones,state="disabled",command=guardar_saldo,text="Guardar",width=8,bg = "lightblue", font=("Arial", 8,"bold"),activebackground="blue")
        self.boton_guardar.grid(row=1,column=0,padx=5,pady=[0,2])
        
        self.mostrar_ultimo_saldo()
        
        #====================================================================
        # Lista de seleccion
        
        def fixed_map(option):

            # Fix for setting text colour for Tkinter 8.6.9
            # From: https://core.tcl.tk/tk/info/509cafafae
            #
            # Returns the style map for 'option' with any styles starting with
            # ('!disabled', '!selected', ...) filtered out.
            # style.map() returns an empty list for missing options, so this
                                    
            # should be future-safe.
            return [elm for elm in style.map('Treeview', query_opt=option) if
                elm[:2] != ('!disabled', '!selected')]
        
        style = Style()
        style.theme_use("clam")
        style.configure("Treeview",background="silver",foreground="black",fieldbackground="silver", highlightthickness=0, bd=0, font=('Arial', 9,"bold")) # Modify the font of the body
        
        style.configure("Treeview.Heading", font=('Arial', 10,'bold')) # Modify the font of the headings
        style.map('Treeview', foreground=fixed_map('foreground'), background=fixed_map('background'))
        style.map("Treeview",background=[("selected","midnight blue")])
        
        
        

        
        frameSeleccion = tk.LabelFrame(self.ventana,text="Selección",font=("Arial",12,"bold"))
        frameSeleccion.grid(row=1,column=1,padx=5,pady=1,sticky="NS")
        
        scrol = Scrollbar(frameSeleccion)
        scrol.grid(row=0,column=2,sticky="NSEW")
        
        self.listaSeleccion = Treeview(frameSeleccion,yscrollcommand=scrol.set,height=5,style="Treeview",columns=["#1","#2"],selectmode="browse")
        self.listaSeleccion.grid(row=0,column=0,sticky="NS")
        
        scrol.config(command=self.listaSeleccion.yview)
        
        self.listaSeleccion.column("#0",width=300,anchor=tk.CENTER,minwidth=100)
        self.listaSeleccion.column("#1",width=200,anchor=tk.CENTER,minwidth=90)
        self.listaSeleccion.column("#2",width=200,anchor=tk.CENTER,minwidth=90)

        
        self.listaSeleccion.heading("#0", text="Período")
        self.listaSeleccion.heading("#1", text="Saldo del mes")
        self.listaSeleccion.heading("#2", text="Saldo acumulado")
        

        
        self.obtener_periodos()
        
        self.listaSeleccion.bind("<<TreeviewSelect>>",lambda e: self.mostrar_detalle())
        
        
        #====================================================================
        # Lista de detalle
        
        frameDetalle = tk.LabelFrame(self.ventana,text="Detalle", font=("Arial",12,"bold"))
        frameDetalle.grid(row=2,column=1,padx=5,pady=1,sticky="NS")
        
        scrol = Scrollbar(frameDetalle)
        scrol.grid(row=0,column=2,sticky="NSEW")
        
        self.listaDetalle = Treeview(frameDetalle,height=13,yscrollcommand=scrol.set,selectmode="browse",columns=("#1","#2","#3","#4","#5","#6","#7"),style="mystyle.Treeview")
        self.listaDetalle.grid(row=0,column=0,sticky="NS")
        
        scrol.config(command=self.listaSeleccion.yview)
        
        self.listaDetalle.tag_configure("Pos", background='lawn green')
        self.listaDetalle.tag_configure("Neg", background='orange')
        self.listaDetalle.tag_configure("ECH", foreground="blue")
        
        
        self.listaDetalle.column("#0", width=0, minwidth=0)
        self.listaDetalle.column("#1", width=80,anchor=tk.CENTER,minwidth=70)
        self.listaDetalle.column("#2", width=80,anchor=tk.CENTER,minwidth=70)
        self.listaDetalle.column("#3", width=200,anchor=tk.CENTER,minwidth=70)
        self.listaDetalle.column("#4", width=80,anchor=tk.CENTER,minwidth=70)
        self.listaDetalle.column("#5", width=60,anchor=tk.CENTER,minwidth=70)
        self.listaDetalle.column("#6", width=100,anchor=tk.CENTER, minwidth=70)
        self.listaDetalle.column("#7", width=100,anchor=tk.CENTER,minwidth=70)
        
        
        self.listaDetalle.heading("#1", text="F. Emisión")
        self.listaDetalle.heading("#2",text="Número")
        self.listaDetalle.heading("#3",text="Entidad")
        self.listaDetalle.heading("#4",text="F. Vto")
        self.listaDetalle.heading("#5",text="Tipo")
        self.listaDetalle.heading("#6",text="Estado")
        self.listaDetalle.heading("#7",text="Importe")
        
        
        self.center(self.ventana)
        
    def habilitar_botones(self):
        self.botonAgregarCheque["state"]="normal"
        self.botonEliminarCheque["state"]="normal"
        self.botonModificarCheque["state"]="normal"
        self.botonAgregarEnte["state"]="normal"
        self.boton_actualizar["state"]="normal"
        self.botonDepositarCheque["state"]="normal"
        self.botonCursoCheque['state']="normal"
        self.botonVerGraficaSaldo['state']="normal"
        self.botonVerGraficaSaldoAc['state']="normal"
        self.botonActuaizar['state']="normal"
        self.botonListarCheques['state']="normal"


        
    def inhabilitar_botones(self):
        self.botonAgregarCheque["state"]="disabled"
        self.botonEliminarCheque["state"]="disabled"
        self.botonModificarCheque["state"]="disabled"
        self.botonAgregarEnte["state"]="disabled"
        self.boton_actualizar["state"]="disabled"
        self.botonDepositarCheque["state"]="disabled"
        self.botonCursoCheque['state']="disabled"
        self.botonVerGraficaSaldo['state']="disabled"
        self.botonVerGraficaSaldoAc['state']="disabled"
        self.botonActuaizar['state']="disabled"
        self.botonListarCheques['state']="disabled"
        
    
    def ver_entidades(self):
        if self.flag_modificando == 1:
            return
        else:
            for elementos in self.listaEntidades.get_children():
                self.listaEntidades.delete(elementos)
            if self.cuit_var.get()==1:
                self.e_cuit['state']="normal"
                self.e_nombre.delete(0,"end")
                self.e_nombre['state']="disabled"
                self.e_cuit.focus_force()
                query = "SELECT id_entidad, entidad, cuit FROM entidades WHERE cuit LIKE '%{}%' ORDER BY entidad ASC".format(self.cadena2.get())
            else:
                self.e_cuit.delete(0,"end")
                self.e_cuit['state']="disabled"
                self.e_nombre['state']="normal"
                self.e_nombre.focus_force()
                query = "SELECT id_entidad, entidad, cuit FROM entidades WHERE entidad LIKE '%{}%' ORDER BY entidad ASC".format(self.cadena.get())
                
            entidades = self.run_query(query)
            for entes in entidades:
                if str(entes[2])=="None" or str(entes[2])=="NONE":
                    self.listaEntidades.insert("", "end", text=entes[0],values=[entes[1],""])
                else:    
                    self.listaEntidades.insert("", "end", text=entes[0],values=[entes[1],entes[2]])
    
    def ventana_nueva_entidad(self):
        self.flag_modificando = 0
        ventana = tk.Toplevel()
        ventana.title("Entidades")
        ventana.resizable(0,0)
        ventana.iconbitmap("icono_genericos.ico")
        self.inhabilitar_botones()
        
        fr_nueva_entidad = tk.LabelFrame(ventana,text="Entidades",foreground="green",font=("Arial",12,"bold"))
        fr_nueva_entidad.grid(row=0,column=0,padx=5,pady=5,sticky="NSEW")
        
        tk.Label(fr_nueva_entidad,text="Razón Social / Nombre",font=("Arial",10,"bold")).grid(row=0,column=0,padx=5,pady=5)
        
        self.cadena = tk.StringVar()
        
        self.e_nombre = tk.Entry(fr_nueva_entidad,fg="blue", font=("Arial",10,"bold"),width=40,textvariable=self.cadena)
        self.e_nombre.grid(row=0,column=1,padx=5,pady=5)
        self.e_nombre.focus_force()
        
        tk.Label(fr_nueva_entidad,text="CUIT",font=("Arial",10,"bold")).grid(row=1,column=0,padx=5,pady=5)
        
        self.cadena2 = tk.StringVar()
        self.e_cuit = tk.Entry(fr_nueva_entidad,fg="blue", font=("Arial",10,"bold"),width=40,textvariable=self.cadena2)
        self.e_cuit.grid(row=1,column=1,padx=5,pady=5)
        
                
        self.cuit_var = tk.IntVar()
        check=tk.Checkbutton(fr_nueva_entidad,command=self.ver_entidades,text="Buscar por CUIT",variable=self.cuit_var,onvalue=1,offvalue=0)
        self.cuit_var.set(0)
        check.grid(row=1,column=2,padx=5,pady=5)
        
        # Lista entidades
        
        def fixed_map(option):

            # Fix for setting text colour for Tkinter 8.6.9
            # From: https://core.tcl.tk/tk/info/509cafafae
            #
            # Returns the style map for 'option' with any styles starting with
            # ('!disabled', '!selected', ...) filtered out.
            # style.map() returns an empty list for missing options, so this
                                    
            # should be future-safe.
            return [elm for elm in style.map('Treeview', query_opt=option) if
                elm[:2] != ('!disabled', '!selected')]
        
        style = Style()
        style.theme_use("clam")
        style.configure("Treeview",background="silver",foreground="black",fieldbackground="silver", highlightthickness=0, bd=0, font=('Arial', 9,"bold")) # Modify the font of the body
        
        style.configure("Treeview.Heading", font=('Arial', 10,'bold')) # Modify the font of the headings
        style.map('Treeview', foreground=fixed_map('foreground'), background=fixed_map('background'))
        style.map("Treeview",background=[("selected","midnight blue")])
        
        scrol = Scrollbar(fr_nueva_entidad)
        scrol.grid(row=2,column=3,sticky="NSEW")
        
        self.listaEntidades = Treeview(fr_nueva_entidad,yscrollcommand=scrol.set,height=10,style="Treeview",columns=["#1","#2"])
        self.listaEntidades.grid(row=2,column=0,columnspan=3,sticky="NSEW")
        
        scrol.config(command=self.listaEntidades.yview)
        
        self.listaEntidades.column("#0",width=0,minwidth=0,stretch= 0)
        self.listaEntidades.column("#1",width=350,anchor=tk.CENTER,minwidth=150)
        self.listaEntidades.column("#2",width=150,anchor=tk.CENTER,minwidth=150)

        
        self.listaEntidades.heading("#1", text="Entidad")
        self.listaEntidades.heading("#2", text="CUIT")
        
        self.listaEntidades["displaycolumns"]=[0,1]
        
        
        
        frameBotonesEntes = tk.Frame(fr_nueva_entidad)
        frameBotonesEntes.grid(row=3,column=0,columnspan=3,sticky="NS")
        
        def inhabilitar_widgets_entes(m):
            self.botonAgregarEnte2['state']="disabled"
            self.botonEliminarEnte['state']="disabled"
            self.botonModificarEnte['state']="disabled"
            self.botonGuardarEnte['state']="normal"
            if m == "n":
                self.botonGuardarEnte['command'] = guardar_nuevo_ente
            else:
                self.botonGuardarEnte['command'] = guardar_actualizar_ente
            self.botonCancelarEnte['state']='normal'


            
        def habilitar_widgets_entes():
            self.botonAgregarEnte2['state']="normal"
            self.botonEliminarEnte['state']="normal"
            self.botonModificarEnte['state']="normal"
            self.botonGuardarEnte['state']="disabled"
            self.botonCancelarEnte['state']='disabled'

            
        
        def cargar_nueva_entidad():
            inhabilitar_widgets_entes(m="n")
            self.flag_modificando = 1
            self.e_nombre['state']="normal"
            self.e_cuit['state']="normal"
            self.e_nombre.delete(0,"end")
            self.e_cuit.delete(0,"end")
            self.e_nombre.focus_force()
            self.e_nombre["background"]="spring green"
            self.e_cuit["background"]="spring green"
            
        def cancelar():
            op = messagebox.askokcancel(parent=ventana,title="Confirmar",message="¿Desea cancelar la operación?\n Todos los cambios se perderan")
            if op == True:
                self.e_nombre.delete(0,"end")
                self.e_cuit.delete(0,"end")
                self.e_nombre["background"]="white"
                self.e_cuit["background"]="white"
                habilitar_widgets_entes()
                self.flag_modificando = 0
                self.ver_entidades()
                mensaje["text"]=""
                
        def guardar_actualizar_ente():
            texto = ""
            if len(self.e_nombre.get().strip()) == 0:
                texto = texto + "Debes ingresar un nombre.\n"
                mensaje["text"]= texto
            
            cuit_actual = self.listaEntidades.item(self.listaEntidades.get_children()[0])
            cuit_a = cuit_actual['values'][1]
            id_actual = cuit_actual['text']
            
            query = "SELECT cuit FROM entidades WHERE cuit is not null"
            cuits = self.run_query(query)
            for i in cuits:
                if str(i[0]) == str(cuit_a):
                    continue
                if str(self.e_cuit.get()).strip() == str(i[0]):
                    texto = texto + "Ya existe otra entidad ingresada con el mismo CUIT."
                    mensaje["text"] = texto
            if texto == "":
                if len(self.e_cuit.get().strip()) == 0:
                    query = "UPDATE entidades SET entidad ='{}' WHERE id_entidad = {}".format(str(self.e_nombre.get()).strip().upper(),id_actual)
                else:
                    query = "UPDATE entidades SET entidad = '{}', cuit = '{}' WHERE id_entidad = {}".format(str(self.e_nombre.get()).strip().upper(),str(self.e_cuit.get()).strip().upper(),id_actual)
                self.run_query(query)
                self.e_cuit.delete(0,"end")
                self.e_nombre.delete(0,"end")
                if self.cuit_var.get() == 1:
                    self.e_cuit['state']="normal"
                    self.e_nombre['state']="disabled"
                else:
                    self.e_cuit['state']="disabled"
                    self.e_nombre['state']="normal"
                
                self.e_nombre["background"]="white"
                self.e_cuit["background"]="white"
                habilitar_widgets_entes() 
                mensaje["text"]=texto
                messagebox.showinfo(parent=ventana,title="",message="El registro se actualizó correctamente.")
                self.flag_modificando = 0
                self.ver_entidades()
                
        
        def guardar_nuevo_ente():
            texto = ""
            if len(self.e_nombre.get().strip()) == 0:
                texto = texto + "Debes ingresar un nombre.\n"
                mensaje["text"]= texto
            
            query = "SELECT entidad FROM entidades"
            entidades = self.run_query(query)
            for i in entidades:
                if (str(i[0]) == str(self.e_nombre.get()).strip().upper()):
                    texto = texto + "El nombre o razón social ingresado ya fue cargado.\n"
                    mensaje["text"] = texto

            query = "SELECT cuit FROM entidades WHERE cuit is not null"
            cuits = self.run_query(query)
            for i in cuits:
                if str(self.e_cuit.get()).strip() == str(i[0]):
                    texto = texto + "Ya existe otra entidad ingresada con el mismo CUIT."
                    mensaje["text"] = texto
            if texto == "":
                if len(self.e_cuit.get().strip()) == 0:
                    query = "INSERT INTO entidades(entidad) VALUES ('{}')".format(str(self.e_nombre.get()).strip().upper())
                else:
                    query = "INSERT INTO entidades(entidad,cuit) VALUES ('{}','{}')".format(str(self.e_nombre.get()).strip().upper(),str(self.e_cuit.get()).strip().upper())
                self.run_query(query)
                self.e_cuit.delete(0,"end")
                self.e_nombre.delete(0,"end")
                if self.cuit_var.get() == 1:
                    self.e_cuit['state']="normal"
                    self.e_nombre['state']="disabled"
                else:
                    self.e_cuit['state']="disabled"
                    self.e_nombre['state']="normal"
                self.e_nombre["background"]="white"
                self.e_cuit["background"]="white"
                habilitar_widgets_entes()
                mensaje["text"]=texto
                self.flag_modificando = 0
                self.ver_entidades()
                messagebox.showinfo(parent=ventana,title="",message="El registro se guardó correctamente.")
                
                
        def eliminar():
            if self.listaEntidades.selection():
                item = self.listaEntidades.item(self.listaEntidades.selection()[0])
                op = messagebox.askyesno(parent=ventana,title="Confirmar",message="¿Desea eliminar el registro?")
                if op == True:
                    query = "DELETE FROM entidades WHERE id_entidad = {}".format(item['text'])
                    self.run_query(query)
                    query = "DELETE FROM cheques WHERE id_entidad = {}".format(item['text'])
                    self.run_query(query)
                    self.ver_entidades()
                    messagebox.showinfo(parent=ventana,title="",message="El registro se eliminó correctamente.")
                    self.obtener_periodos()
                    self.mostrar_detalle()
            else:
                messagebox.showerror(parent=ventana,title="Error",message="Primero debes seleccionar la entidad que vas a eliminar")
        
        def modificar():
            if self.listaEntidades.selection():
                self.flag_modificando = 1
                item = self.listaEntidades.item(self.listaEntidades.selection()[0]) 
                id__ = item["text"]
                nombre = item["values"][0]
                cuit = item["values"][1]
                for i in self.listaEntidades.get_children():
                    self.listaEntidades.delete(i)
                self.listaEntidades.insert("", 0, text= id__, values=[nombre,cuit])
                item = self.listaEntidades.item(self.listaEntidades.get_children()[0])
                inhabilitar_widgets_entes(m="e")
                self.e_nombre['state']="normal"
                self.e_cuit['state']="normal"
                self.e_nombre.delete(0,"end")
                self.e_cuit.delete(0,"end")
                self.e_nombre.insert(0,item['values'][0])
                if (str(item['values'][1]).strip() == str("NONE") or str(item['values'][1]).strip() == str("None")):
                    self.e_cuit.insert(0,"")
                else:
                    self.e_cuit.insert(0,str(item['values'][1]).strip())
                
                self.e_nombre.focus_force()
                self.e_nombre["background"]="green yellow"
                self.e_cuit["background"]="green yellow"
                
            else:
                messagebox.showerror(parent=ventana,title="Error",message="Primero debes seleccionar la entidad que vas a actualizar")
            
        
        self.botonAgregarEnte2 = tk.Button(frameBotonesEntes,command=cargar_nueva_entidad,text="Nuevo",width=10,bg = "greenyellow",font=("Arial",8,"bold"),activebackground="green")
        self.botonAgregarEnte2.grid(row=0,column=0,padx=15,pady=10)
        
        self.botonEliminarEnte = tk.Button(frameBotonesEntes,command=eliminar,text="Eliminar",width=10, bg= "tomato",font=("Arial",8,"bold"),activebackground="red")
        self.botonEliminarEnte.grid(row=0,column=1,padx=15,pady=10)
        
        self.botonModificarEnte = tk.Button(frameBotonesEntes,command=modificar,text="Modificar",width=10, bg = "gold",font=("Arial",8,"bold"),activebackground="yellow")
        self.botonModificarEnte.grid(row=0,column=2,padx=15,pady=10)
        
        self.botonGuardarEnte = tk.Button(frameBotonesEntes,state="disabled",text="Guardar",width=10, bg = "lightblue",font=("Arial",8,"bold"),activebackground="blue")
        self.botonGuardarEnte.grid(row=0,column=3,padx=15,pady=10)
        
        self.botonCancelarEnte = tk.Button(frameBotonesEntes,command=cancelar,state="disabled",text="Cancelar",width=10, bg = "lightblue",font=("Arial",8,"bold"),activebackground="blue")
        self.botonCancelarEnte.grid(row=0,column=4,padx=15,pady=10)
        
        mensaje = tk.Label(fr_nueva_entidad,text="",foreground="red")
        mensaje.grid(row=4,column=0,columnspan=4,padx=5,pady=5)
        
        self.cadena.trace("w",lambda x,y,z: self.ver_entidades())
        self.cadena2.trace("w",lambda x,y,z: self.ver_entidades())
        self.ver_entidades()
        
        def on_closing():
            self.habilitar_botones()
            ventana.destroy()
            
        ventana.protocol("WM_DELETE_WINDOW", on_closing)
        
        self.center(ventana)
        

        
    def modificar_cheque(self):
        if self.listaDetalle.selection():
            item = self.listaDetalle.item(self.listaDetalle.selection()[0])
            id__ = str(item["text"])
            f1 = item['values'][0]
            numerocheque = item['values'][1]
            entidad = item['values'][2]
            f2 = item['values'][3]
            imp = item['values'][6]
            echeck = item['values'][4]
            
            
            
            ventana = tk.Toplevel()
            ventana.title("Modificar cheque")
            ventana.iconbitmap("icono_genericos.ico")
            ventana.resizable(0,0)
            self.inhabilitar_botones()
            
            fr_mod_cheque=tk.LabelFrame(ventana,text="Nuevo Cheque",foreground="green",font=("Arial",12,"bold"))
            fr_mod_cheque.grid(row=0,column=0,padx=5,pady=5,sticky="NSEW")
            
            tk.Label(fr_mod_cheque,text="Número de cheque",font=("Arial",10,"bold")).grid(row=0,column=0,padx=5,pady=5)
            
            e_numero = tk.Entry(fr_mod_cheque,fg="blue",font=("Arial",10,"bold"),bg="gold")
            e_numero.grid(row=0,column=1,padx=5,pady=5)
            e_numero.focus_force()
            e_numero.insert(0,numerocheque)
            
            tk.Label(fr_mod_cheque,text="Fecha de Emisión",font=("Arial",10,"bold")).grid(row=1,column=0,padx=5,pady=5)
            
            e_fe = tk.Entry(fr_mod_cheque,fg="blue",font=("Arial",10,"bold"),bg="gold")
            e_fe.grid(row=1,column=1,padx=5,pady=5)
            e_fe.insert(0, f1)
            
            def validar1():
                if e_fe.get() == "dd/mm/aaaa":
                    e_fe.delete("0","end")
            
            def validar2():
                if len(e_fe.get()) == 0:
                    e_fe.insert(0, "dd/mm/aaaa")
    
            
            e_fe.bind("<FocusIn>",lambda e: validar1())
            e_fe.bind("<FocusOut>",lambda d: validar2())
            
            tk.Label(fr_mod_cheque,text="Entidad",font=("Arial",10,"bold")).grid(row=2,column=0,padx=5,pady=5)
            
            query = "SELECT entidad FROM entidades"
            lista_entidades = self.run_query(query)
            entidades = []
            for i in lista_entidades:
                entidades.append(i[0])
            def check_input(event):
                value = event.widget.get()
            
                if value == '':
                    combo['values'] = entidades
                else:
                    data = []
                    for item in entidades:
                        if value.lower() in item.lower():
                            data.append(item)

                    combo['values'] = data
            
            combo = Combobox(fr_mod_cheque)
            combo["values"] = entidades
            combo.bind("<KeyRelease>",check_input)
            combo.grid(row=2,column=1)
            combo.set(entidad)
            
            tk.Label(fr_mod_cheque,text="Fecha de Vencimiento",font=("Arial",10,"bold")).grid(row=3,column=0,padx=5,pady=5)
            
            e_vto = tk.Entry(fr_mod_cheque,fg="blue",font=("Arial",10,"bold"),bg="gold")
            e_vto.grid(row=3,column=1,padx=5,pady=5)
            e_vto.insert(0, f2)
            
            def validar12():
                if e_vto.get() == "dd/mm/aaaa":
                    e_vto.delete("0","end")
            
            def validar22():
                if len(e_vto.get()) == 0:
                    e_vto.insert(0, "dd/mm/aaaa")

            
            e_vto.bind("<FocusIn>",lambda e: validar12())
            e_vto.bind("<FocusOut>",lambda d: validar22())
            
            tk.Label(fr_mod_cheque,text="Importe ($)",font=("Arial",10,"bold")).grid(row=4,column=0,padx=5,pady=5)
            
            importe = tk.Entry(fr_mod_cheque,fg="blue",font=("Arial",10,"bold"),bg="gold")
            importe.grid(row=4,column=1,padx=5,pady=5)
            importe.insert(0,imp)
            
            
            self.echeq_si_no= tk.IntVar()
                
            c_e_cheq = tk.Checkbutton(fr_mod_cheque,text="E-CHEQ",variable=self.echeq_si_no,onvalue=1,offvalue=0)
            c_e_cheq.grid(row=5,column=1,padx=[0,5],pady=5,sticky="W")
            
            if str(echeck) == str("ECHEQ"):
                print("es echeck")
                self.echeq_si_no.set(value=1)
            else:
                print("no es echeck")
                self.echeq_si_no.set(value=0)
            
            
            mensaje = tk.Label(fr_mod_cheque,text="",foreground="red")
            mensaje.grid(row=7,column=0,columnspan=2,padx=5,pady=5)
            
            def actualizar_nuevo_cheque(id_,num,fe,ent,vto,imp):
                texto = ""
                query = "SELECT numero FROM cheques"
                numeros = self.run_query(query)
                for i in numeros:
                    if (str(i[0]) == str(numerocheque)):
                        continue
                    if (str(i[0]) == num):
                        texto = texto + "El número de cheque ingresado ya fue cargado.\n"
                        mensaje["text"] = texto
                if (fe.count("/")!=2 or vto.count("/")!=2 or len(fe.strip()) != 10 or len(fe.strip()) != 10 ):
                    texto = texto + "Las fechas ingresadas deben respetar el formato 'dd/mm/aaaa'.\n"
                    mensaje["text"] = texto 
                try: 
                    fecha_emi = datetime.strptime(fe, "%d/%m/%Y")
                except:
                    texto = texto + "Ingrese una fecha de emisión válida.\n"
                    mensaje["text"]= texto
                try:
                    fecha_vto = datetime.strptime(vto, "%d/%m/%Y")
                except:
                    texto = texto + "Ingrese una fecha de vencimiento válida.\n"
                    mensaje["text"] = texto
                try:
                    if (fecha_vto < fecha_emi):
                        texto = texto + "La fecha de vencimiento es menor a la fecha de emisión.\n"
                        mensaje["text"]= texto
                except:
                    pass
                if combo.get().strip() not in entidades:
                        texto = texto + "Debe elegir una Entidad que haya sido cargada.\n"
                        mensaje["text"] = texto
                try:
                    imp=round(float(imp),2)
                except:
                    texto = texto + "Ingrese el importe en valor numérico.\nUtiliza punto (.) en lugar de coma (,) para indicar la parte decimal.\n"
                    mensaje["text"] = texto
                mensaje["text"]= texto
                self.center(ventana)
                

                
                if texto == "":      
                    fecha_emi = fecha_emi.strftime("%Y-%m-%d")
                    fecha_vto = fecha_vto.strftime("%Y-%m-%d")
                    #Carga de datos
                    query = "SELECT id_entidad FROM entidades WHERE entidad = '{}'".format(combo.get().strip())
                    parametros = (combo.get().strip())
                    id_ = self.run_query(query)
                    for i in id_:
                        query = "UPDATE cheques SET numero='{}',fecha_e='{}',fecha_v='{}',id_entidad={},importe={},pendiente=1, echeq={} WHERE id={}".format(num,fecha_emi,fecha_vto,i[0],imp,int(self.echeq_si_no.get()),id__)
                        self.run_query(query)
                        
                    self.obtener_periodos()
                    self.mostrar_detalle()
                        
                    self.habilitar_botones()
                    ventana.destroy()
                    messagebox.showinfo(title="Actualización exitosa",message="El registro se actualizó correctamente.")
            
            boton = tk.Button(fr_mod_cheque,text="Cargar",bg="lightblue",font=("Arial",10,"bold"),activebackground="blue",
                              command=lambda: actualizar_nuevo_cheque(int(id__),e_numero.get(),e_fe.get(),combo.get(),e_vto.get(),importe.get()))
            boton.grid(row=6,column=0,columnspan=2,padx=5,pady=5)
            
            boton.bind("<Return>",lambda e: actualizar_nuevo_cheque(int(id__),e_numero.get(),e_fe.get(),combo.get(),e_vto.get(),importe.get()))
            
            def on_closing():
                self.habilitar_botones()
                ventana.destroy()
                
            ventana.protocol("WM_DELETE_WINDOW", on_closing)

            self.center(ventana)
        
        else:
            messagebox.showerror(title="Error",message="Primero debés seleccionar el cheque que querés modificar.")
        

    

        
        
    def ventana_nuevo_cheque(self):
        ventana = tk.Toplevel()
        ventana.title("Cargar nuevo cheque")
        ventana.resizable(0,0)
        ventana.iconbitmap("icono_genericos.ico")
        
        self.inhabilitar_botones()
        
        fr_nuevo_cheque=tk.LabelFrame(ventana,text="Nuevo Cheque",foreground="green",font=("Arial",12,"bold"))
        fr_nuevo_cheque.grid(row=0,column=0,padx=5,pady=5,sticky="NSEW")
        
        
        tk.Label(fr_nuevo_cheque,text="Número de cheque",font=("Arial",10,"bold")).grid(row=0,column=0,padx=5,pady=5)
        
        e_numero = tk.Entry(fr_nuevo_cheque,fg="blue",font=("Arial",10,"bold"))
        e_numero.grid(row=0,column=1,padx=5,pady=5)
        e_numero.focus_force()
        
        tk.Label(fr_nuevo_cheque,text="Fecha de Emisión",font=("Arial",10,"bold")).grid(row=1,column=0,padx=5,pady=5)
        
        e_fe = tk.Entry(fr_nuevo_cheque,fg="blue",font=("Arial",10,"bold"))
        e_fe.grid(row=1,column=1,padx=5,pady=5)
        fecha_actual = datetime.now()
        fecha_actual = fecha_actual.strftime("%d/%m/%Y")
        e_fe.insert(0, fecha_actual)
        
        def validar1():
            if e_fe.get() == "dd/mm/aaaa":
                e_fe.delete("0","end")
        
        def validar2():
            if len(e_fe.get()) == 0:
                e_fe.insert(0, "dd/mm/aaaa")

        
        e_fe.bind("<FocusIn>",lambda e: validar1())
        e_fe.bind("<FocusOut>",lambda d: validar2())
        
        tk.Label(fr_nuevo_cheque,text="Entidad",font=("Arial",10,"bold")).grid(row=2,column=0,padx=5,pady=5)
        
        query = "SELECT entidad FROM entidades"
        lista_entidades = self.run_query(query)
        entidades = []
        for i in lista_entidades:
            entidades.append(i[0])
        def check_input(event):
            value = event.widget.get()
        
            if value == '':
                combo['values'] = entidades
            else:
                data = []
                for item in entidades:
                    if value.lower() in item.lower():
                        data.append(item)

                combo['values'] = data
        
        combo = Combobox(fr_nuevo_cheque)
        combo["values"] = entidades
        combo.bind("<KeyRelease>",check_input)
        combo.grid(row=2,column=1)
        
        tk.Label(fr_nuevo_cheque,text="Fecha de Vencimiento",font=("Arial",10,"bold")).grid(row=3,column=0,padx=5,pady=5)
        
        e_vto = tk.Entry(fr_nuevo_cheque,fg="blue",font=("Arial",10,"bold"))
        e_vto.grid(row=3,column=1,padx=5,pady=5)
        e_vto.insert(0, "dd/mm/aaaa")
        
        def validar12():
            if e_vto.get() == "dd/mm/aaaa":
                e_vto.delete("0","end")
        
        def validar22():
            if len(e_vto.get()) == 0:
                e_vto.insert(0, "dd/mm/aaaa")

        
        e_vto.bind("<FocusIn>",lambda e: validar12())
        e_vto.bind("<FocusOut>",lambda d: validar22())
        
        tk.Label(fr_nuevo_cheque,text="Importe ($)",font=("Arial",10,"bold")).grid(row=4,column=0,padx=5,pady=5)
        
        importe = tk.Entry(fr_nuevo_cheque,fg="blue",font=("Arial",10,"bold"))
        importe.grid(row=4,column=1,padx=5,pady=5)
        
        es_echeq= tk.IntVar()
        
        c_e_cheq = tk.Checkbutton(fr_nuevo_cheque,text="E-CHEQ",variable=es_echeq,onvalue=1,offvalue=0)
        c_e_cheq.grid(row=5,column=1,padx=[0,5],pady=5,sticky="W")
        
        es_echeq.set(0)
        
        mensaje = tk.Label(fr_nuevo_cheque,text="",foreground="red")
        mensaje.grid(row=7,column=0,columnspan=2,padx=5,pady=5)
        

        
        def cargar_nuevo_cheque(num,fe,ent,vto,imp,echeq):
            texto = ""
            query = "SELECT numero FROM cheques"
            numeros = self.run_query(query)
            for i in numeros:
                if (str(i[0]) == num):
                    texto = texto + "El número de cheque ingresado ya fue cargado.\n"
                    mensaje["text"] = texto
            if (fe.count("/")!=2 or vto.count("/")!=2 or len(fe.strip()) != 10 or len(vto.strip()) != 10 ):
                texto = texto + "Las fechas ingresadas deben respetar el formato 'dd/mm/aaaa'.\n"
                mensaje["text"] = texto 
            try: 
                fecha_emi = datetime.strptime(fe, "%d/%m/%Y")
            except:
                texto = texto + "Ingrese una fecha de emisión válida.\n"
                mensaje["text"]= texto
            try:
                fecha_vto = datetime.strptime(vto, "%d/%m/%Y")
            except:
                texto = texto + "Ingrese una fecha de vencimiento válida.\n"
                mensaje["text"] = texto
            try:
                if (fecha_vto < fecha_emi):
                    texto = texto + "La fecha de vencimiento es menor a la fecha de emisión.\n"
                    mensaje["text"]= texto
            except:
                pass
            if combo.get().strip() not in entidades:
                    texto = texto + "Debe elegir una Entidad que haya sido cargada.\n"
                    mensaje["text"] = texto
            try:
                imp=round(float(imp),2)
            except:
                texto = texto + "Ingrese el importe en valor numérico.\nUtiliza punto (.) en lugar de coma (,) para indicar la parte decimal.\n"
                mensaje["text"] = texto
            mensaje["text"]= texto
            self.center(ventana)
            

            
            if texto == "":      
                fecha_emi = fecha_emi.strftime("%Y-%m-%d")
                fecha_vto = fecha_vto.strftime("%Y-%m-%d")
                #Carga de datos
                query = "SELECT id_entidad FROM entidades WHERE entidad = '{}'".format(combo.get().strip())
                #parametros = (combo.get().strip())
                id_ = self.run_query(query)
                for i in id_:
                    query = "INSERT INTO cheques VALUES (NULL,'{}','{}','{}',{},{},1,{},0)".format(num,fecha_emi,fecha_vto,i[0],imp,int(echeq))
                    self.run_query(query)
                    
                    self.mostrar_detalle()
                    self.obtener_periodos()
                    
                    
                    self.habilitar_botones()
                    ventana.destroy()
                    messagebox.showinfo(title="Carga Exitosa",message="El cheque se cargó correctamente.")
        
        boton = tk.Button(fr_nuevo_cheque,text="Cargar",bg="lightblue",font=("Arial",10,"bold"),activebackground="blue",
                          command=lambda: cargar_nuevo_cheque(e_numero.get(),e_fe.get(),combo.get(),e_vto.get(),importe.get(),es_echeq.get()))
        boton.grid(row=6,column=0,columnspan=2,padx=5,pady=5)
        boton.bind("<Return>", lambda e: cargar_nuevo_cheque(e_numero.get(),e_fe.get(),combo.get(),e_vto.get(),importe.get(),es_echeq.get()) )
        
        
        

        
        def on_closing():
            self.habilitar_botones()
            ventana.destroy()
            
        ventana.protocol("WM_DELETE_WINDOW", on_closing)
        self.center(ventana)
        
        
        
        
    
    def center(self,win):
        try:
            win.eval('tk::PlaceWindow . center')
        except:
            win.update_idletasks()
            width = win.winfo_width()
            height = win.winfo_height()
            x = (win.winfo_screenwidth() // 2) - (width // 2)
            y = (win.winfo_screenheight() // 2) - (height // 2)
            win.geometry('+{}+{}'.format(x, y))
    




if __name__ == "__main__":

    ventana = tk.Tk()
    App = Aplicacion(ventana)
    ventana.mainloop()
    

