import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk
import pandas as pd
import zipfile
import os

def load_csv():
    global df
    csv_file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
    if csv_file_path:
        try:
            # Tenta carregar o arquivo CSV
            df = pd.read_csv(csv_file_path, encoding='utf-8')  # ou 'latin-1', dependendo do arquivo
            
            # Verifica se o dataframe foi carregado corretamente
            if df is not None and not df.empty:
                # Verifica se o CSV contém as colunas esperadas
                if 'cnpj' in df.columns and 'nome' in df.columns and 'cpf' in df.columns:
                    # Converter a coluna 'cpf' para string para garantir a manipulação correta
                    df['cpf'] = df['cpf'].astype(str)
                    
                    # Exibir o conteúdo do CSV diretamente na interface
                    text_box.delete(1.0, tk.END)
                    text_box.insert(tk.END, df.to_string(index=False))
                    messagebox.showinfo("CSV Carregado", "Arquivo CSV carregado com sucesso.")
                    return df
                else:
                    messagebox.showerror("Erro", "O CSV deve conter as colunas 'cnpj', 'nome' e 'cpf'.")
            else:
                messagebox.showerror("Erro", "O arquivo CSV está vazio ou não pôde ser lido.")
        except pd.errors.EmptyDataError:
            messagebox.showerror("Erro", "O arquivo CSV está vazio.")
        except pd.errors.ParserError:
            messagebox.showerror("Erro", "Erro ao processar o arquivo CSV. Verifique o formato do arquivo.")
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível carregar o arquivo CSV: {str(e)}")
    else:
        messagebox.showwarning("Aviso", "Nenhum arquivo CSV selecionado.")
    return None
    
def generate_files():
    if df is None:
        messagebox.showerror("Erro", "Nenhum arquivo CSV foi carregado.")
        return
    
    start_date = start_date_entry.get()
    end_date = end_date_entry.get()
    
    if not start_date or not end_date:
        messagebox.showerror("Erro", "As Datas de inicio e fim devem ser preenchidas.")
        return
        
    output_dir = filedialog.askdirectory()
    if not output_dir:
        return
    
    try:
        # Abrir o arquivo modelo e ler o conteúdo apenas uma vez
        with open(r'C:\Users\marxe\Documents\MODELO_ECF.txt', 'r', encoding='utf-8') as model_file:  # ou 'latin-1'
            model_content = model_file.read()
            
        with zipfile.ZipFile(os.path.join(output_dir, 'EFD_Arquivos.zip'), 'w') as zipf:
            for index, row in df.iterrows():
                filename = f"{row['cnpj']}_ECF.txt"
                filepath = os.path.join(output_dir, filename)
                
                # Substituir as variáveis no conteúdo do modelo
                content = model_content.replace('VARIAVEL_01', str(row['cnpj']))
                content = content.replace('VARIAVEL_02', '0') # 0
                content = content.replace('VARIAVEL_03', '0') # 0
                content = content.replace('VARIAVEL_04', '0') # 0
                content = content.replace('VARIAVEL_05', '0') # 0
                content = content.replace('VARIAVEL_06', '0') # 0
                content = content.replace('VARIAVEL_07', '0') # 0
                content = content.replace('VARIAVEL_08', '0') # 0
                content = content.replace('VARIAVEL_09', '0') # 0
                content = content.replace('VARIAVEL_10', '0') # 0
                content = content.replace('VARIAVEL_11', '0') # 0
                content = content.replace('VARIAVEL_12', '0') # 0
                content = content.replace('VARIAVEL_13', '0') # 0
                content = content.replace('VARIAVEL_14', '0') # 0
                cpf_ajustado = str(row['cpf'])[1:]
                content = content.replace('VARIAVEL_15', cpf_ajustado)
                content = content.replace('VARIAVEL_16', str(row['nome']))
                content = content.replace('VARIAVEL_datain', str(start_date))
                content = content.replace('VARIAVEL_dataend', str(end_date))
                
                # Criar e escrever no novo arquivo TXT        
                with open(filepath, 'w') as txt_file:
                    txt_file.write(content)
                
                # Adicionar o arquivo ao zip e depois removê-lo    
                zipf.write(filepath, filename)
                os.remove(filepath)
                
        messagebox.showinfo("Sucesso", "Arquivos TXT gerados e compactados com sucesso.")
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro ao gerar os arquivos: {str(e)}")

# Configurando a janela principal
app = ctk.CTk()
app.title("Gerador de Arquivos ECF")
app.geometry("700x500")

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("green")

df = None # variavel global para armazenar o dataframe

# Elementos da interface
title_label = ctk.CTkLabel(app, text="Gerador de Arquivos ECF", font=("Arial", 20))
title_label.pack(pady=20)

upload_button = ctk.CTkButton(app, text="Carregar CSV", command=load_csv)
upload_button.pack(pady=10)

text_box = ctk.CTkTextbox(app, height=150)
text_box.pack(pady=10)

data_frame = ctk.CTkFrame(app)
data_frame.pack(pady=20)

start_date_label = ctk.CTkLabel(data_frame, text="Data de Início:")
start_date_label.grid(row=0, column=1, padx=10)

start_date_entry = ctk.CTkEntry(data_frame)
start_date_entry.grid(row=0, column=1, padx=10)

end_date_entry = ctk.CTkLabel(data_frame, text="Data de Fim:")
end_date_entry.grid(row=0, column=2, padx=10) 

end_date_entry = ctk.CTkEntry(data_frame)
end_date_entry.grid(row=0, column=3, padx=10)            

genrate_button = ctk.CTkButton(app, text="Gerar Arquivos", command=generate_files)
genrate_button.pack(pady=20)

app.mainloop()