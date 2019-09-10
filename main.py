import os
import threading
import openpyxl
import kivy
from kivy.app import App
from kivy.uix.label import Label
from kivy.uix.textinput import TextInput
from kivy.uix.button import Button
from kivy.uix.floatlayout import FloatLayout
from kivy.uix.popup import Popup
from kivy.properties import ObjectProperty
from kivy.core.window import Window
import excel
import onedrive

CONF_FILE = "TecMartz.conf"
NUMERO_DISENOS = 0
PATH_EXCEL = ""
URL_EXCEL = ""
MODELOS = {}


class PopupResult(Popup):
    lblResult = ObjectProperty()

    def __init__(self, result, busqueda, **kwargs):
        super(Popup, self).__init__(**kwargs)
        self.lblResult.text = result
        self.title += " \"" + busqueda + "\""


class PopupConf(Popup):
    txtNumDisenos = ObjectProperty()
    txtLocalExcel = ObjectProperty()
    txtRemoteExcel = ObjectProperty()

    def __init__(self, **kwargs):
        super(Popup, self).__init__(**kwargs)
        self.txtNumDisenos.text = str(NUMERO_DISENOS)
        self.txtLocalExcel.text = PATH_EXCEL
        self.txtRemoteExcel.text = URL_EXCEL

    def on_dismiss(self):
        global NUMERO_DISENOS, PATH_EXCEL, URL_EXCEL
        if NUMERO_DISENOS != int(self.txtNumDisenos.text) \
            or PATH_EXCEL != self.txtLocalExcel.text \
            or URL_EXCEL != self.txtRemoteExcel.text:
            #print("saved")
            """ print(NUMERO_DISENOS, self.txtNumDisenos.text)
            print(PATH_EXCEL, self.txtLocalExcel.text)
            print(URL_EXCEL,self.txtRemoteExcel.text)
 """
            NUMERO_DISENOS = int(self.txtNumDisenos.text)
            PATH_EXCEL = self.txtLocalExcel.text
            URL_EXCEL  = self.txtRemoteExcel.text

            f = open(CONF_FILE, "w")
            f.write("\n".join([str(NUMERO_DISENOS), PATH_EXCEL, URL_EXCEL]) + "\n")
            f.close()


class MyGrid(FloatLayout):
    txtModelo = ObjectProperty()
    lblEstado = ObjectProperty()
    btnUpdate = ObjectProperty()
    btnSearch = ObjectProperty()

    def inicializar_globales(self):
        global NUMERO_DISENOS, PATH_EXCEL, URL_EXCEL
        if os.path.isfile(CONF_FILE):
            #print("Loading from file")
            f = open(CONF_FILE, "r")
            d = f.readlines()
            f.close()

            NUMERO_DISENOS = int(d[0][:-1])
            PATH_EXCEL = d[1][:-1]
            URL_EXCEL  = d[2]
        else:
            #print("No conf file found")
            NUMERO_DISENOS = 8
            PATH_EXCEL = "Existencias actual 070719.xlsx"
            URL_EXCEL  = "" #CENSORED
            f = open(CONF_FILE, "w")
            f.write("\n".join([str(NUMERO_DISENOS), PATH_EXCEL, URL_EXCEL]) + "\n")
            f.close()

            
    def descargar_y_leer_excel(self, force_download=False):
        self.btnUpdate.disabled = True
        global MODELOS
        
        if not os.path.isfile(PATH_EXCEL) or force_download:
            self.lblEstado.text = "Descargando excel..."
            try:
                onedrive.download_excel(URL_EXCEL, PATH_EXCEL)
            except Exception as err:
                self.lblEstado.text = "Error descargando excel: " + str(err.__repr__())
                self.btnUpdate.disabled = False
                return

        self.lblEstado.text = "Leyendo excel..."
        try:
            MODELOS, errores = excel.get_modelos(PATH_EXCEL, NUMERO_DISENOS)
            self.lblEstado.text = "\n".join(errores) + "\n\nListo"
        except Exception as err:
            self.lblEstado.text = "Error leyendo excel: " + str(err.__repr__())
            self.btnUpdate.disabled = False
            return

        self.btnUpdate.disabled = False


    def btnSearch_onclick(self):
        threading.Thread(target=self._btnSearch_onclick).start()

    def btnUpdate_onclick(self):
        threading.Thread(target=self._btnUpdate_onclick).start()

    def btnConf_onclick(self):
        popup = PopupConf()
        popup.open()

    def _btnUpdate_onclick(self):
        self.descargar_y_leer_excel(True)

    def _btnSearch_onclick(self):
        #print("Leyendo excel...")
        busqueda = self.txtModelo.text
        if not busqueda:
            self.lblEstado.text = "Por favor introduzca el modelo a buscar."
            return

        self.btnSearch.disabled = True
        self.lblEstado.text = "Realizando la búsqueda..."
        
        try:
            result = excel.buscar_modelos(MODELOS, busqueda)
        except Exception as err:
            self.lblEstado.text = "Error buscando el modelo: " + str(err.__repr__())
            self.btnSearch.disabled = False
            return
        result_str = excel.sprint_modelos(result)

        popup_msg = result_str if result_str else "No se ha encontrado ningún modelo."
        popup = PopupResult(popup_msg, busqueda)
        popup.open()
        
        self.lblEstado.text = "Listo"
        self.btnSearch.disabled = False


class MyApp(App):
    def build(self):
        self.title = "Klecko fucking boss"
        #Window.size = (360, 640)
        Window.clearcolor = (1,1,1,1)
        return MyGrid()

    def on_start(self):
        self.root.inicializar_globales()
        self.root.descargar_y_leer_excel()

if __name__ == '__main__':
    MyApp().run()
