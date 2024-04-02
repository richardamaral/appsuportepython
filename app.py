import tkinter as tk
import pyperclip
import os
import subprocess
import webbrowser
from tkinter import messagebox
import pygetwindow as gw
import pyautogui
import pytesseract
import smtplib
from PIL import Image, ImageTk
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from tkinter import filedialog
from email.mime.text import MIMEText
import wmi
import psutil
import platform
import socket


def capturar_info_sistema():
    def get_system_info():
        system_info = {}
        # Informações do Sistema Operacional
        system_info["Sistema Operacional"] = platform.platform()

        # Informações de IP Interno
        ips = []
        try:
            # Obter todos os IPs de todas as interfaces de rede
            for interface, addresses in psutil.net_if_addrs().items():
                for address in addresses:
                    if address.family == socket.AF_INET and address.address != '127.0.0.1' and not address.address.startswith('169.254'):
                        ips.append(address.address)
        except:
            ips.append("N/A")

        system_info["IPs Internos"] = ips

        # Informações da Conexão de Rede (Wi-Fi ou Ethernet)
        ethernet_up = False
        wifi_up = False
        interfaces = psutil.net_if_stats()
        for interface, data in interfaces.items():
            if interface.startswith('Ethernet') and data.isup:
                ethernet_up = True
            elif interface.startswith('Wi-Fi') and data.isup:
                wifi_up = True

        if ethernet_up:
            system_info["Tipo de Conexão"] = "Cabeada"
        elif wifi_up:
            system_info["Tipo de Conexão"] = "Wi-Fi"
        else:
            system_info["Tipo de Conexão"] = "Desconhecido"

        return system_info

    def get_motherboard_info():
        motherboard_info = {}
        c = wmi.WMI()

        # Informações da Placa Mãe
        for board in c.Win32_BaseBoard():
            motherboard_info["Fabricante"] = board.Manufacturer
            motherboard_info["Modelo"] = board.Product

        # Informações da CPU
        for cpu in c.Win32_Processor():
            motherboard_info["CPU"] = cpu.Name

        # Informações de Memória RAM
        mem = psutil.virtual_memory()
        motherboard_info["Memória RAM Total (GB)"] = round(mem.total / (1024 ** 3), 2)
        motherboard_info["Memória RAM Usada (GB)"] = round(mem.used / (1024 ** 3), 2)
        motherboard_info["Memória RAM Disponível (GB)"] = round(mem.available / (1024 ** 3), 2)

        return motherboard_info

    # Obtenha as informações do sistema e da placa mãe
    system_info = get_system_info()
    motherboard_info = get_motherboard_info()

    # Retorne as informações capturadas
    return system_info, motherboard_info




imagem_selecionada = None
def selecionar_imagem():
    filepath = filedialog.askopenfilename(initialdir=os.getcwd(), title="Selecione uma imagem",
                                          filetypes=(("Arquivos de Imagem", "*.png;*.jpg;*.jpeg"), ("Todos os arquivos", "*.*")))
    if filepath:

        global imagem_selecionada
        imagem_selecionada = filepath
def on_entry_click(entry, placeholder_text):
    if entry.get() == placeholder_text:
        entry.delete(0, tk.END)
        entry.config(fg='white')

    elif entry.get() == "":
        entry.insert(0, placeholder_text)
        entry.config(fg='grey')

def nome_on_click(event):
    on_entry_click(nome_entry, "Digite seu Nome..")

def anydesk_on_click(event):
    on_entry_click(anydesk_entry, "Digite o Anydesk..")

def telefone_on_click(event):
    on_entry_click(telefone_entry, "Digite o Telefone..")
def assunto_on_click(event):
    on_entry_click(assunto_entry, "Digite o Assunto do Chamado..")


def resize_image(image_path, width, height):
    original_image = Image.open(image_path)
    resized_image = original_image.resize((width, height))
    return ImageTk.PhotoImage(resized_image)


def abrir_anydesk_ou_download():
    # Lista de possíveis caminhos para o executável do AnyDesk
    caminhos_anydesk = [
        "C:\\Program Files\\AnyDesk\\AnyDesk.exe",
        "C:\\Program Files (x86)\\AnyDesk\\AnyDesk.exe",
        "C:\\Arquivos de Programas\\AnyDesk\\AnyDesk.exe",
        "C:\\Arquivos de Programas (x86)\\AnyDesk\\AnyDesk.exe",
        "C:\\Program Files (x86)\\AnyDesk\\AnyDesk.exe",
        os.path.join(os.path.expanduser("~"), "Downloads", "AnyDesk.exe")
        # Adicione mais caminhos se necessário
        #C:\Users\ADMINISTRATOR\Downloads
    ]

    # Tentar abrir o AnyDesk a partir de cada caminho na lista
    anydesk_encontrado = False
    for caminho in caminhos_anydesk:
        try:
            subprocess.Popen([caminho])
            anydesk_encontrado = True
            break
        except FileNotFoundError:
            continue

    # Se o AnyDesk não for encontrado, redirecionar para a página de download
    if not anydesk_encontrado:
        webbrowser.open_new("https://anydesk.com/pt/downloads/thank-you?dv=win_exe")

def capturar_anydesk_codigo():
    # Tente obter o código do AnyDesk da área de transferência
    codigo_anydesk = pyperclip.paste()

    # Se o código do AnyDesk estiver presente na área de transferência, coloque-o no campo anydesk_entry
    if codigo_anydesk.startswith("AnyDesk"):
        anydesk_entry.delete(0, 'end')
        anydesk_entry.insert(0, codigo_anydesk)


def atualizar_visibilidade_anexar_imagem():
    if imagem_selecionada_checkbox_var.get() == 1:
        selecionar_imagem_label.grid(row=7, column=0, padx=(45, 0), pady=10)
        selecionar_imagem_button.grid(row=7, column=0, columnspan=2, padx=(50, 0), pady=10)
    else:
        selecionar_imagem_label.grid_forget()
        selecionar_imagem_button.grid_forget()


def limitar_caracteres(event):
    # Obtém o texto atual no telefone_entry
    telefone = telefone_entry.get()

    # Limita o número de caracteres permitidos (por exemplo, 14 caracteres)
    if len(telefone) > 14:
        # Corta o texto se ultrapassar o limite
        telefone_entry.delete(14, tk.END)

def formatar_telefone(event):
    # Obtém o texto atual no telefone_entry
    telefone = telefone_entry.get()

    # Remove todos os caracteres que não sejam dígitos
    telefone_formatado = ''.join(filter(str.isdigit, telefone))

    # Verifica se o telefone possui mais de 10 dígitos (sem o DDD)
    if len(telefone_formatado) > 10:
        # Formata o telefone com código de área (DDD)
        telefone_formatado = f"({telefone_formatado[:2]}) {telefone_formatado[2:7]}-{telefone_formatado[7:]}"
    else:
        # Formata o telefone sem código de área (DDD)
        telefone_formatado = f"{telefone_formatado[:5]}-{telefone_formatado[5:]}"

    # Atualiza o texto no telefone_entry com o telefone formatado
    telefone_entry.delete(0, tk.END)
    telefone_entry.insert(0, telefone_formatado)

def limitar_caracteres_telefone(event):
    # Obtém o texto atual no telefone_entry
    telefone = telefone_entry.get()

    # Limita o número de caracteres do telefone para 15 (incluindo caracteres especiais)
    if len(telefone) > 14:
        # Corta o texto se ultrapassar o limite
        telefone_entry.delete(14, tk.END)



def obter_ajuda():
    webbrowser.open_new("https://www.youtube.com/watch?v=Ugi0dbW7PMY&ab_channel=rxck011")


def obter_regras():
    mensagem = ("INSTALAÇÃO DE SOFTWARES (WORD, EXCEL, OUTLOOK).\n\n"
                "INSTALAÇÃO E CONFIGURAÇÃO DE IMPRESSORAS/SCANNERS.\n\n"
                "MAL FUNCIONAMENTO DE INFRAESTRUTURA. (IMPRESSORAS COM TONER, SCANNER E COMPUTADORES.)\n\n"
                "CONFIGURAÇÃO DO E-MAIL @ATENDIMENTOPRO OU @CEDUSP NO APLICATIVO OUTLOOK.\n\n"
                "CONFIGURAÇÃO (ONE DRIVE) DE PASTA EM NUVEM SINCRONIZADA.\n\n")

    messagebox.showinfo("[SERVIÇOS CORPNET]",mensagem)

def formatar_anydesk(event):
    # Obtém o texto atual no anydesk_entry
    anydesk = anydesk_entry.get()

    # Remove todos os caracteres que não sejam dígitos
    anydesk_formatado = ''.join(filter(str.isdigit, anydesk))

    # Limita o número de caracteres do AnyDesk para 10
    anydesk_formatado = anydesk_formatado[:10]

    # Atualiza o texto no anydesk_entry com o AnyDesk formatado
    anydesk_entry.delete(0, tk.END)
    anydesk_entry.insert(0, anydesk_formatado)


def limitar_caracteres_anydesk(event):
    # Obtém o texto atual no anydesk_entry
    anydesk = anydesk_entry.get()

    # Limita o número de caracteres do AnyDesk para 10
    if len(anydesk) > 
        # Corta o texto se ultrapassar o limite
        anydesk_entry.delete(10, tk.END)
def enviar_email():
    system_info, motherboard_info = capturar_info_sistema()
    # Obtendo os valores dos campos de entrada


    nome_string = nome_entry.get()
    anydesk_num = anydesk_entry.get()
    telefone_num = telefone_entry.get()
    descricao_texto = descricao_text.get("1.0", "end-1c")
    assunto_string = assunto_entry.get()

    if nome_string == "Digite seu Nome..":
        nome_string = ""
    if anydesk_num == "Digite o Anydesk..":
        anydesk_num = ""
    if telefone_num == "Digite o Telefone..":
        telefone_num = ""
    if assunto_string == "Digite o Assunto do Chamado..":
        assunto_string = ""

    assunto = f'PRO OCUPACIONAL - {assunto_string}'

    if imagem_selecionada is not None:
        with open(imagem_selecionada, 'rb') as f:
            img_data = f.read()
        img_part = MIMEImage(img_data, name=os.path.basename(imagem_selecionada))





    smtp_server = 'smtp.office365.com' ## SERVIDOR PRA SE CONECTAR AO E-MAIL REMETENTE, NO CASO O MEU ERA OFFICE 365 ENTÃO UTILIZEI ESSE COM A PORTA 587.
    porta = 587
    remetente = 'email@email.com.br'
    destinatario = 'email@email.com'
    senha = 'XXXXXXXXX' ## SENHA DO E-MAIL REMETENTE


    msg = MIMEMultipart()
    if imagem_selecionada is not None:
        msg.attach(img_part)
    else:
        None


    msg['From'] = remetente
    msg['To'] = destinatario
    msg['Subject'] = assunto

    html = f"""
       <html>
         <body>
         <b>[SOLICITAÇÃO DE SUPORTE]</b>
         <br>
         <br>
        <p>Nome: <b>{nome_string}</b></p>
       <p>Anydesk: <b>{anydesk_num}</b></p>
       <p>Telefone: <b>{telefone_num}</b></p>
       <br>
       <p>Descrição: <b>{descricao_texto}</b></p>
       <br>
       <br>
             
       """
    html += """
                 <b>Informações do Sistema:</b>
               """

    for key, value in system_info.items():
        html += f"<li>{key}: {value}</li>"

    for key, value in motherboard_info.items():
        html += f"<li>{key}: {value}</li>"

    html += """
             <br>
             <p><b>Atenciosamente, Pro Ocupacional.</b></p>
         </body>
       </html>
           """

    msg.attach(MIMEText(html, 'html'))
    # Enviando o e-mail
    try:
        # Enviando o e-mail
        with smtplib.SMTP(smtp_server, porta) as server:
            server.starttls()
            server.login(remetente, senha)
            server.sendmail(remetente, destinatario, msg.as_string())

        # Exibir mensagem de sucesso
        messagebox.showinfo("Sucesso", "Chamado criado com sucesso, agora aguarde com o AnyDesk aberto a solicitação de alguns de nossos profissionais!")

        # Limpar os campos de entrada
        nome_entry.delete(0, 'end')
        anydesk_entry.delete(0, 'end')
        telefone_entry.delete(0, 'end')
        assunto_entry.delete(0, 'end')
        descricao_text.delete(1.0, 'end')

    except smtplib.SMTPRecipientsRefused as e:
        # Exibir mensagem de erro
        messagebox.showerror("Erro", "Endereço de e-mail inválido, por favor digite novamente!")

        # Limpar o campo de e-mail
        anydesk_entry.delete(0, 'end')







def ao_clicar_dev(event):
    webbrowser.open_new("https://github.com/richardamaral")



# Criando a janela principal
root = tk.Tk()
root.title("PRO OCUPACIONAL - SUPORTE T.I")

root.geometry("500x870")
root.configure(bg="#427367")
root.resizable(False, False)

root.iconbitmap("img/suporte-tecnico.ico")

logo_image = resize_image("img/logo-pro.png", 300, 100)

imagem_selecionada_checkbox_var = tk.IntVar()


imagem_selecionada_checkbox = tk.Checkbutton(root, text="< Anexar Imagem", variable=imagem_selecionada_checkbox_var, command=atualizar_visibilidade_anexar_imagem, font=("monospace", 12, "bold"), bg="#427367", fg="white", selectcolor="#427367")

# Posicionando o Checkbutton
imagem_selecionada_checkbox.grid(row=6, column=1, padx=10, pady=10)




anydesk_image = tk.PhotoImage(file="img/anydesk_image.png")
anydesk_image = resize_image("img/anydesk_image.png", 200, 50)


dev_image = tk.PhotoImage(file="img/codificacao.png")
dev_image = resize_image("img/codificacao.png", 50, 50)

# Criando os widgets
nome_label = tk.Label(root, text="Nome:", font=("monospace", 12, "bold"),  bg="#427367", fg="white")
nome_entry = tk.Entry(root, width=40, bg="black", font=("monospace", 12, "bold"),  fg="#4D4E50")


anydesk_label = tk.Label(root, text="Anydesk:", font=("monospace", 12, "bold"),  bg="#427367", fg="#8E0505")
anydesk_entry = tk.Entry(root, width=40, bg="black", font=("monospace", 12, "bold"),  fg="#4D4E50")


telefone_label = tk.Label(root, text="Telefone:", font=("monospace", 12, "bold"),  bg="#427367", fg="white")
telefone_entry = tk.Entry(root, width=40, bg="black", font=("monospace", 12, "bold"),  fg="#4D4E50")

assunto_label = tk.Label(root, text="Assunto:", font=("monospace", 12, "bold"),  bg="#427367", fg="white")
assunto_entry = tk.Entry(root, width=40, bg="black", font=("monospace", 12, "bold"),  fg="#4D4E50")


descricao_label = tk.Label(root, text="Motivo:", font=("monospace", 12, "bold"),  bg="#427367", fg="white")
descricao_text = tk.Text(root, height=8, width=40, font=("monospace", 12, "bold"), bg="#717171", fg="black")


selecionar_imagem_label = tk.Label(root, text="Imagem:", font=("monospace", 12, "bold"),  bg="#427367", fg="white")
selecionar_imagem_button = tk.Button(root, text="Enviar Arquivo", command=selecionar_imagem, font=("monospace", 12, "bold"),  bg='black', fg='white', width='20')


video_explicativo_button = tk.Button(root, text="Ajuda", command=obter_ajuda,  font=("monospace", 12, "bold"),  bg='black', fg='cyan', width='10')

info_button = tk.Button(root, text="Regras Importantes", command=obter_regras,  font=("monospace", 12, "bold"),  bg='black', fg='red', width='20')




anydesk_button = tk.Button(root, command=abrir_anydesk_ou_download, image=anydesk_image, font=("monospace", 12, "bold"), fg='white', width='250')
enviar_button = tk.Button(root, text="Abrir Chamado", command=enviar_email, font=("monospace", 20, "bold"),  bg='black', fg='#2AF655', width='20')



logo_label = tk.Label(root, image=logo_image, bg="#427367")
logo_label.grid(row=0, column=0, columnspan=2, padx=10, pady=10)

nome_label.grid(row=1, column=0, padx=10, pady=5)
nome_entry.grid(row=1, column=1, padx=10, pady=5)
nome_entry.insert(0, "Digite seu Nome..")
nome_entry.bind('<FocusIn>', nome_on_click)
nome_entry.bind('<FocusOut>', nome_on_click)

anydesk_label.grid(row=2, column=0, padx=10, pady=5)
anydesk_entry.grid(row=2, column=1, padx=10, pady=5)
anydesk_entry.insert(0, "Digite o Anydesk..")
anydesk_entry.bind('<FocusIn>', anydesk_on_click)
anydesk_entry.bind('<FocusOut>', anydesk_on_click)
anydesk_entry.bind('<FocusOut>', formatar_anydesk)
anydesk_entry.bind('<Key>', limitar_caracteres_anydesk)

telefone_label.grid(row=3, column=0, padx=10, pady=5)
telefone_entry.grid(row=3, column=1, padx=10, pady=5)
telefone_entry.insert(0, "Digite o Telefone..")
telefone_entry.bind('<FocusIn>', telefone_on_click)
telefone_entry.bind('<FocusOut>', telefone_on_click)
telefone_entry.bind('<FocusOut>', formatar_telefone)
telefone_entry.bind('<Key>', limitar_caracteres_telefone)

assunto_label.grid(row=4, column=0, padx=10, pady=5)
assunto_entry.grid(row=4, column=1, padx=10, pady=5)
assunto_entry.insert(0, "Digite o Assunto do Chamado..")
assunto_entry.bind('<FocusIn>', assunto_on_click)
assunto_entry.bind('<FocusOut>', assunto_on_click)

descricao_label.grid(row=5, column=0, padx=10, pady=5)
descricao_text.grid(row=5, column=1, padx=10, pady=5)


imagem_selecionada_checkbox.grid(row=6, column=1, padx=10, pady=10)

selecionar_imagem_label.grid(row=7, column=0, padx=(45, 0), pady=10)
selecionar_imagem_button.grid(row=7, column=0, columnspan=2, padx=(50, 0), pady=10)


enviar_button.grid(row=8, column=0, columnspan=2, padx=10, pady=(30, 0))
anydesk_button.grid(row=9, column=0, columnspan=2, padx=10, pady=(50, 0))



video_explicativo_button.grid(row=10, column=0, columnspan=2, padx=(5, 0), pady=(45, 0))

info_button.grid(row=11, column=0, columnspan=2, padx=(5, 0), pady=(15, 0))

dev_label = tk.Label(root, image=dev_image, bg="#427367")
dev_label.grid(row=12, column=0, columnspan=2, padx=10, pady=(15, 0))



dev_label.bind("<Button-1>", ao_clicar_dev)


atualizar_visibilidade_anexar_imagem()
root.mainloop()

