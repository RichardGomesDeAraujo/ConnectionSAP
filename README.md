<img src="HOD.png" align="Center" alt="Hands On Data" style="height: 180px; width:260px;"/>

<p>  <br>
  </p>

# How to Connect SAP with Python 
<p>  <br>
  </p>

###### by [Richard Gomes de Araújo](https://github.com/RichardGomesDeAraujo) - 14/02/2024
[![Github Badge](https://img.shields.io/badge/-Github-000?style=flat-square&logo=Github&logoColor=white&link=https://github.com/RichardGomesDeAraujo)](https://github.com/RichardGomesDeAraujo)
[![Linkedin Badge](https://img.shields.io/badge/-LinkedIn-blue?style=flat-square&logo=Linkedin&logoColor=white&link=https://www.linkedin.com/in/richardaraujoanalistadedados/)](https://www.linkedin.com/in/richardaraujoanalistadedados/)
[![Youtube Badge](https://img.shields.io/badge/-YouTube-ff0000?style=flat-square&labelColor=ff0000&logo=youtube&logoColor=white&link=https://www.youtube.com/channel/UCc_jlqHut_GkXc8ahgQHOOw)](https://www.youtube.com/channel/UCc_jlqHut_GkXc8ahgQHOOw)
<p>  <br>
  </p>
  
# Índice
- [**Option 1**](README.md#Option-1)
- [**Option 2**](README.md#Option-2)
<p>  <br>
  </p>
  
>### Option 1

```Python
# criar uma ambiente virtual = python -n venv venv
# entrar no venv = .\venv\Scripsts\activate
# instalar o win32 no venv = pip install pywin32
# Começar o código em Python
## Importar as biblotecas necessárias

import win32com.client
import subprocess
import sys
import time
from tkinter import *
from tkinter import messagebox

# criar uma classe
class SapGui(object):
    def __init__(self):
        # informar o caminha onde está instalado o SAP
        self.path = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
        # o subprocess vai abrir o SAP do local informado
        subprocess.Popen(self.path)
        
        # criar uma variável para instanciar a aplicação
        self.SapGuiAuto = win32com.client.GetObject("SAPGUI")
        application = self.SapGuiAuto.GetScriptingEngine
        
        # Criar a conexão com o SAP
        ## Colocar o nome da conexão obtido pelo sistema em Conexões, Características, Parâmetros Ligação ao sistema, Descrição
        self.connection = application.OpenConnection("BMF [ecc1.ddns.net]", True)
        # criar um time para o Python abrir o sistema e maximizar a tela
        time.sleep(3)
        self.session = self.connection.Children(0)
        self.session.findById("wnd[0]").maximize
        
    def saplogin(self):
        try:
            # client é uma informação que está na tela de login do SAP
            self.session.findById("wnd[0]/usr/txtRSYST-MANDT").text = "800"
            # informar os dados de login, senha e idioma
            self.session.findById("wnd[0]/usr/txtRSYST-BNAME").text = "usuario"
            self.session.findById("wnd[0]/usr/txtRSYST-BCODE").text = "senha"
            self.session.findById("wnd[0]/usr/txtRSYST-LANGU").text = "PT"
            # comando equivalente ao Enter do sistema
            self.session.findById("wnd[0]").sendVKey(0)
            
        # Em caso de erro o comando irá imprimir através do sys    
        except:
            print(sys.exc_info()[0])
        # comando para abrir a mensagem depois do método de login    
        messagebox.showinfo("showinfo", "Login Realizado com Sucesso!")

# comando para executar o arquivo. Se for o arquivo principal (main) ele executa           
if __name__ == '__main__':
    window = Tk()
    window.geometry("300x70")
    # utiliza a expressão lambda para o programa não rodar executando
    botao = Button(window, text="Login SAP", command= lambda :SapGui().sapLogin())
    botao.pack()
    # comando para executar a janela do login do TkInter
    mainloop()

```

###### [⏪](README.md#Índice)
<p>  <br>
  </p>

  
>### Option 2

```Python
#  Script para Conexão SAP 
#  Criar um ambiente virtual para instalar os pacotes 
#  Comando para criar o ambiente virtual no terminal = python -m venv venv 
#  Comando para ativar a venv = .\venv\Scripts\activate 
#  Comando para liberar execução de Scripts = Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope CurrentUser
#  Comando para instalar o pacote pywin32 = pip install pywin32 

import win32com.client

sapguiauto = win32com.client.GetObject("SAPGUI")
application = sapguiauto.GetScriptingEngine
connection = application.Children(0)
session = connection.Children(0)

print(type(session))

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "iw38"

# Verificar os seguintes passos no SAP:
# Canto direito no ícone de Opções, Opções, Acessibilidade & scripting, Scripting:
# Desmarcar a opção "Notificar se um script se vincular a um SAP GUI" e
# Desmarcar a opção "Notificar quando um script abre uma ligação"
 
# If Not IsObject(application) Then
#    Set SapGuiAuto = GetObject("SAPGUI")
#    Set application = SapGuiAuto.GetScriptingEngine
# End if

# If Not IsObject(connection) Then
#    Set connection = application.Children(0)
# End if

# If Not IsObject(session) Then
#    Set session = connection.Children(0)
# End if

#**** Esta parte não é utilizada ****#
# If IsObject(WScript) Then
#    WScript.ConnectObject session, "on"
#    WScript.ConnectObject application, "on"
# End if
#************************************#

# session.findById("wnd[0]").maximize
# session.findById("wnd[0]/tbar[0]/okcd").text = "iw38"
# session.findById("wnd[0]").sendVKey 0
# session.findById("wnd[0]/tbar[1]/btn[8]").press 
# session.findById("wnd[0]/tbar[0]/btn[3]").press
# session.findById("wnd[0]/tbar[0]/btn[3]").press

```

###### [⏪](README.md#Índice)
<p>  <br>
  </p>
