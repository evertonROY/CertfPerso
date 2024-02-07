import tkinter as tk
from tkinter import filedialog, ttk
import openpyxl
from PIL import Image, ImageDraw, ImageFont
from ttkthemes import ThemedStyle

# Função para fechar a janela
def fechar_janela():
    janela.destroy()

arquivo = False
diretorio = False

# Função para selecionar um arquivo
def selecionar_arquivo():
    global arquivo
    arquivo = filedialog.askopenfilename(filetypes=(("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*")))
    arquivo_entry.delete(0, tk.END)
    arquivo_entry.insert(0, arquivo)

# Função para escolher o diretório de salvamento
def escolher_diretorio():
    global diretorio
    diretorio = filedialog.askdirectory()
    if not diretorio.endswith('/'):
            diretorio += '/'
    diretorio_entry.delete(0, tk.END)
    diretorio_entry.insert(0, diretorio)
    
    

# Função para começar o processo
def comecar_processo():
    if not arquivo:
       resultado_text.config(state=tk.NORMAL)  # Ativando a edição temporariamente
       resultado_text.delete(1.0, tk.END)
       resultado_text.insert(tk.END, "Adicione um arquivo xlsx!\n")
       resultado_text.config(state=tk.DISABLED)  # Desativando a edição novamente
    elif not diretorio:
       resultado_text.config(state=tk.NORMAL)  # Ativando a edição temporariamente
       resultado_text.delete(1.0, tk.END)
       resultado_text.insert(tk.END, "Adicione um diretório para salvar!\n")
       resultado_text.config(state=tk.DISABLED)  # Desativando a edição novamente
    else:
       workbook_alunos = openpyxl.load_workbook(arquivo)
       global sheet_alunos
       sheet_alunos = workbook_alunos['Sheet1']

       # Simulando um processo e exibindo o resultado na área de texto
       resultado_text.config(state=tk.NORMAL)  # Ativando a edição temporariamente
       resultado_text.insert(tk.END, "========Gerando certificados========\n")
       for indice, linha in enumerate(sheet_alunos.iter_rows(min_row=2, max_row=15)):
           resultado_text.config(state=tk.NORMAL)  # Ativando a edição temporariamente
           #cada célula que contém a informação que precisamos
           nome_curso = linha[0].value
           nome_participante = linha[1].value
           tipo_participacao = linha[2].value
           data_inicio = linha[3].value
           data_termino = linha[4].value
           carga_horaria = linha[5].value
           data_emissao = linha[6].value
           print(indice + 1, nome_participante)
           resultado_text.insert(tk.END, f"{indice + 1} {nome_participante} ok\n")
           resultado_text.config(state=tk.DISABLED)  # Desativando a edição novamente
           resultado_text.update_idletasks()
           resultado_text.see(tk.END)  # Força a atualização da interface gráfica
           #Transferir os dados da planilha para a imagem do certificado
           font_nome = ImageFont.truetype('./tahomabd.ttf',80)
           font_geral = ImageFont.truetype('./tahoma.ttf',75)
           font_hora = ImageFont.truetype('./tahoma.ttf',60)
           image = Image.open('./certificado_padrao.jpg')
           desenhar = ImageDraw.Draw(image)

           desenhar.text((1020,832), nome_participante, fill='black', font=font_nome)
           desenhar.text((1070,955), nome_curso, fill='black', font=font_geral)
           desenhar.text((1440,1070), tipo_participacao, fill='black', font=font_geral)
           desenhar.text((1500,1192), str(carga_horaria), fill='black', font=font_geral)

           desenhar.text((730,1770), data_inicio, fill='black', font=font_hora)
           desenhar.text((730,1920), data_termino, fill='black', font=font_hora)

           desenhar.text((2210,1920), data_emissao, fill='black', font=font_hora)

           image.save(f'{diretorio}{indice} {nome_participante} certificado.png')
       resultado_text.config(state=tk.NORMAL)
       resultado_text.insert(tk.END, f"========Finalizado========\n")
       resultado_text.config(state=tk.DISABLED)
       resultado_text.update_idletasks()
       resultado_text.see(tk.END)

# Criando a janela principal
janela = tk.Tk()
janela.title("Criação de Certificado")

# Definindo o tamanho da janela
largura_janela = 400
altura_janela = 530
largura_tela = janela.winfo_screenwidth()
altura_tela = janela.winfo_screenheight()
pos_x = (largura_tela / 2) - (largura_janela / 2)
pos_y = (altura_tela / 2) - (altura_janela / 2)
janela.geometry('%dx%d+%d+%d' % (largura_janela, altura_janela, pos_x, pos_y))
style = ThemedStyle(janela)
style.set_theme("arc")
estilo_botao = ttk.Style()

# Adicionando widgets
titulo_label = ttk.Label(janela, text="Criação de Certificado", font=("Arial", 16))
titulo_label.pack(pady=(10, 20))

selecionar_arquivo_button = ttk.Button(janela, text="Procurar Arquivo .xlsx", takefocus=False, command=selecionar_arquivo)
selecionar_arquivo_button.pack(pady=0)

titulo_label = ttk.Label(janela, text="ou digite", font=("Arial", 8))
titulo_label.pack(pady=0)

arquivo_entry = ttk.Entry(janela, width=40)
arquivo_entry.pack(pady=0)

selecionar_diretorio_button = ttk.Button(janela, text="Selecionar Diretório", takefocus=False, command=escolher_diretorio)
selecionar_diretorio_button.pack(pady=(20, 0))

titulo_label = ttk.Label(janela, text="ou digite", font=("Arial", 8))
titulo_label.pack(pady=0)

diretorio_entry = ttk.Entry(janela, width=40)
diretorio_entry.pack(pady=0)

comecar_button = ttk.Button(janela, text="Começar", command=comecar_processo)
comecar_button.pack(pady=20)

resultado_text = tk.Text(janela, height=10, width=50)
resultado_text.pack(pady=10)
resultado_text.config(state=tk.DISABLED)  # Configurando a área de texto como somente leitura

fechar_button = ttk.Button(janela, text="Fechar", command=fechar_janela)
fechar_button.pack(pady=10)

# Loop principal da aplicação
janela.mainloop()
