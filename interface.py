import main
import logs
import functions
import users
import threading
import traceback
import json
import os
import ctypes
import signal
import customtkinter as ctk
from tkinter import filedialog
from PIL import Image

# Inciar aplicação ao apertar enter no campo de senha
def on_enter(_) -> None:
    start_app()

# Iniciar aplicação criando uma thread para rodar main.start() sem travar a interface
def start_app() -> None:
    if not entry_excel.get():
        label_feedback.configure(text="Anexe o arquivo Excel antes de inciar.", text_color="green")
        app.after(5000, clear_feedback)
    elif not entry_user.get():
        label_feedback.configure(text="Digite sua matrícula antes de inciar.", text_color="green")
        app.after(5000, clear_feedback)
    elif entry_user.get() not in users.allowed:
        label_feedback.configure(text="Acesso negado!", text_color="red")
        app.after(5000, clear_feedback)
    else:
        label_feedback.configure(text="Registros em andamento, aguarde!", text_color="green")
        thread = threading.Thread(target=run_app, daemon=True)
        thread.start()

# Iniciar main.start() em segundo plano
def run_app() -> None:
    try:
        feedback_text, text_color, feedback_time = main.start(entry_user.get(), entry_password.get(), entry_excel.get(), run_mode.get())
        app.after(0, lambda: label_feedback.configure(text=feedback_text, text_color=text_color))
        app.after(feedback_time, clear_feedback)
    except:
        error_message = traceback.format_exc()
        logs.logging.info(f"Erro: {error_message}")
        app.after(0, lambda: label_feedback.configure(text="Ocorreu algum problema!", text_color="red"))
        app.after(5000, clear_feedback)

# Iniciar extracão criando uma thread para rodar run_extract() sem travar a interface
def start_extract() -> None:
    if not entry_excel.get():
        label_feedback.configure(text="Anexe o arquivo Excel antes de inciar.", text_color="green")
        app.after(5000, clear_feedback)
        app.update_idletasks()
    elif not entry_folder.get():
        label_feedback.configure(text="Informe o nome da pasta antes inciar.", text_color="green")
        app.after(5000, clear_feedback)
    elif not entry_user.get():
        label_feedback.configure(text="Digite sua matrícula antes de inciar.", text_color="green")
        app.after(5000, clear_feedback)
    elif entry_user.get() not in users.allowed:
        label_feedback.configure(text="Acesso negado!", text_color="red")
        app.after(5000, clear_feedback)
    else:
        label_feedback.configure(text="Extração em andamento, aguarde!", text_color="green")
        thread = threading.Thread(target=run_extract, daemon=True)
        thread.start()

# Iniciar functions.extract() em segundo plano
def run_extract() -> None:
    try:
        feedback_text, text_color, feedback_time = functions.extract(entry_folder.get(), entry_excel.get())
        app.after(0, lambda: label_feedback.configure(text=feedback_text, text_color=text_color))
        app.after(feedback_time, clear_feedback)
    except:
        error_message = traceback.format_exc()
        logs.logging.info(f"Erro: {error_message}")
        app.after(0, lambda: label_feedback.configure(text="Ocorreu algum problema!", text_color="red"))
        app.after(5000, clear_feedback)

# Encerrar a aplicação corretamente ao clicar no 'X' ou pressionar Ctrl + C
def exit_app(*args) -> None:
    save_data()
    label_feedback.configure(text=f"Encerrando aplicação.", text_color="red")
    app.after(1000, app.quit)

# Selecionar o arquivo Excel a ser utilizado na aplicação
def select_file():
    file_path = filedialog.askopenfilename(title="Selecione um arquivo Excel", filetypes=[("Planilhas do Excel", "*.xlsx")])
    if file_path:
        entry_excel.delete(0, "end")
        entry_excel.insert(0, file_path)
        entry_excel.xview_moveto(1)

# Alternar visibilidade da senha
def toggle_password() -> None:
    if entry_password.cget("show") == "*":
        entry_password.configure(show="")
        vis_button.configure(image=vis_icon_show)
    else:
        entry_password.configure(show="*")
        vis_button.configure(image=vis_icon_hide)

# Salvar os campos de entrada no arquivo JSON
def save_data() -> None:
    # Carrega os dados existentes, se houver
    if os.path.exists('.\\cfg\\config.json'):
        with open('.\\cfg\\config.json', 'r') as f:
            data = json.load(f)
    else:
        data = {}
    data["Registros Max"] = {
        "User": entry_user.get(),
        "Folder": entry_folder.get(),
        "Background execution": run_mode.get(),
        "Excel file path": entry_excel.get()
    }
    # Salva o JSON atualizado
    with open('.\\cfg\\config.json', 'w') as f:
        json.dump(data, f, indent=4)

# Carregar os campos de entrada do arquivo JSON
def load_data() -> None:
    try:
        with open('.\\cfg\\config.json', 'r') as f:
            data = json.load(f)
            reg_max = data.get("Registros Max", {})

            user = reg_max.get("User", '')
            folder = reg_max.get("Folder")
            run_background = reg_max.get("Background execution")
            excel_path = reg_max.get("Excel file path")
            # Apenas insere se houver um valor salvo
            if user:
                entry_user.insert(0, user)
            if folder:
                entry_folder.insert(0, folder)
            if run_background:
                run_mode.set(run_background)
            if excel_path:
                entry_excel.insert(0, excel_path)
    except FileNotFoundError:
        pass

# Mudar página da interface
def show_frame(frame: ctk.CTkFrame) -> None:
    frame.tkraise()

# Limpa o label_feedback
def clear_feedback():
    label_feedback.configure(text="")
    app.update_idletasks()

# Garante que a interface será centralizada ao iniciar
ctypes.windll.shcore.SetProcessDpiAwareness(1)

# Tema escuro
ctk.set_appearance_mode("dark")

# Configurar o sinal para capturar Ctrl + C e encerrar corretamente
signal.signal(signal.SIGINT, exit_app)

# Janela Principal
app = ctk.CTk()
app.title("Registros Max")
app.iconbitmap(".\\cfg\\icon.ico")
app.resizable(False, False)
app.configure(fg_color="#2b2b2b")

# Criando um container para as páginas
container = ctk.CTkFrame(app, fg_color="transparent")
container.pack(fill="both", expand=True)

# Criando as páginas como frames dentro do container
page1 = ctk.CTkFrame(container, fg_color="#2b2b2b")
page2 = ctk.CTkFrame(container, fg_color="#2b2b2b")

for frame in (page1, page2):
    frame.place(relx=0, rely=0, relwidth=1, relheight=1)

# Centralizar janela
win_width, win_height = 260, 220
screen_width, screen_height = app.winfo_screenwidth(), app.winfo_screenheight()
pos_x = (screen_width - win_width) // 2
pos_y = (screen_height - win_height) // 2
app.geometry(f"{win_width}x{win_height}+{pos_x}+{pos_y}")

# Capturar fechamento da janela
app.protocol("WM_DELETE_WINDOW", exit_app)

# Label com versào do programa
label_version = ctk.CTkLabel(page1, text="v.1.2.1")
label_version.pack(padx=(220,0))

# Campo de entrada da matrícula
entry_user = ctk.CTkEntry(page1, placeholder_text="Matrícula")
entry_user.pack(pady=(2,5))

# Frame para agrupar Entry de senha e botão de visibilidade
frame_password = ctk.CTkFrame(page1, fg_color="transparent")
frame_password.pack(pady=(0,5), fill="x")
entry_password = ctk.CTkEntry(frame_password, placeholder_text="Senha", show="*")
entry_password.pack(side="left", fill="x", padx=(60,5))
entry_password.bind('<Return>', on_enter)
vis_icon_hide = ctk.CTkImage(dark_image=Image.open(".\\cfg\\hide.png"))
vis_icon_show = ctk.CTkImage(dark_image=Image.open(".\\cfg\\show.png"))
vis_button = ctk.CTkButton(frame_password,image=vis_icon_hide, text="", width=30, command=toggle_password)
vis_button.pack(side="left", fill="x")

# Frame para agrupar o botão "Extrair E-mails" e configuração
frame_extract = ctk.CTkFrame(page1, fg_color="transparent")
frame_extract.pack(pady=(0,5), fill="x")
ext_button = ctk.CTkButton(frame_extract, text="Extrair E-mails", font=("Roboto",13,"bold"), command=start_extract)
ext_button.pack(side="left", fill="x", padx=(60,5))
cfg_icon = ctk.CTkImage(dark_image=Image.open(".\\cfg\\cfg.png"))
cfg_button = ctk.CTkButton(frame_extract,image=cfg_icon, text="", width=30, command=lambda: show_frame(page2))
cfg_button.pack(side="left", fill="x")

# Botão iniciar
start_button = ctk.CTkButton(page1, text="Iniciar", font=("Roboto",13,"bold"), command=start_app)
start_button.pack(pady=(0,5))

# Criar switch para executar em Run Mode
run_mode = ctk.BooleanVar()
toggle = ctk.CTkSwitch(page1, text="Executar em segundo plano", variable=run_mode)
toggle.pack()

# Campo de entrada do nome da pasta
entry_folder = ctk.CTkEntry(page2, placeholder_text="Nome da pasta")
entry_folder.pack(pady=(30,0))

# Frame para agrupar campo de entrada do arquivo e botão de selecionar arquivo
frame_file = ctk.CTkFrame(page2, fg_color="transparent")
frame_file.pack(pady=(5,5), fill="x")
entry_excel = ctk.CTkEntry(frame_file, placeholder_text="Selecionar arquivo")
entry_excel.pack(side="left", fill="x", padx=(60,5))
attach_icon = ctk.CTkImage(dark_image=Image.open(".\\cfg\\attach-file.png"))
attach_button = ctk.CTkButton(frame_file, image=attach_icon, text="", width=30, command=select_file)
attach_button.pack(side="left", fill="x")

# Botão Salvar
back_button = ctk.CTkButton(page2, text="Salvar", font=("Roboto",13,"bold"), command=lambda: show_frame(page1))
back_button.pack(pady=(0,12))

# Feedback do processo
label_feedback = ctk.CTkLabel(page1, text="")
label_feedback.pack()

# Informações
frame_info = ctk.CTkFrame(page2, fg_color="transparent")
frame_info.pack(pady=(5,5), fill="x")
label_info_1 = ctk.CTkLabel(frame_info, text="Informações e atualizações:")
label_info_1.pack(side="left", fill="x", padx=(28,2))
label_info_2 = ctk.CTkLabel(frame_info, text="GitHub", font=("Roboto",13,"bold"), text_color="#3B8ED0", cursor="hand2")
label_info_2.pack(side="left", fill="x")
label_info_2.bind("<Button-1>", functions.abrir_link)
label_info_3 = ctk.CTkLabel(page2, text="Desenvolvido por: José Ribeiro Neto")
label_info_3.pack()

# Carregar dados
load_data()
entry_excel.xview_moveto(1)

# Exibir a página inicial
show_frame(page1)

# Iniciar aplicação
app.mainloop()