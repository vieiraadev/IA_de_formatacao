import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import re
def carregar_arquivo():
    global df
    caminho_arquivo = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx")])
    if caminho_arquivo:

        try:
            df = pd.read_excel(caminho_arquivo)
            atualizar_treeview("Arquivo carregado", df)
            messagebox.showinfo("Sucesso", "Arquivo carregado com sucesso!")

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar o arquivo: {e}")

def exportar_para_excel(df, comando):
    caminho_arquivo = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Arquivos Excel", "*.xlsx")])
    if caminho_arquivo:

        try:
            df.to_excel(caminho_arquivo, index=False)
            messagebox.showinfo("Sucesso", f"Arquivo exportado com sucesso!\nComando: {comando}")

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao exportar o arquivo: {e}")

def atualizar_treeview(comando, data_frame):
    frame_interacao = tk.Frame(frame_chat, bg="#f0f0f0")
    frame_interacao.pack(pady=10, fill='both', expand=True)
    label_comando = tk.Label(frame_interacao,
                             text=f"Comando: {comando}",
                             font=("Arial", 12, "bold"),
                             bg="#f0f0f0",
                             fg="#333")
    label_comando.pack(anchor="w", padx=5)
    tree = ttk.Treeview(frame_interacao)
    tree.pack(pady=5, fill='x', expand=True)
    tree["columns"] = list(data_frame.columns)
    tree["show"] = "headings"
    for col in data_frame.columns:
        tree.heading(col, text=col)
        tree.column(col, width=100)
    for indice, linha in data_frame.iterrows():
        tree.insert("", "end", values=list(linha))
    btn_exportar = tk.Button(frame_interacao,
                             text="Exportar para Excel",
                             command=lambda: exportar_para_excel(data_frame, comando),
                             bg="#4CAF50",
                             fg="white",
                             font=("Arial", 10))

    btn_exportar.pack(pady=5)
    canvas.update_idletasks()
    canvas.configure(scrollregion=canvas.bbox("all"))
    canvas.yview_moveto(1)

def processar_comando():
    global df
    comando = entry_comando.get().lower()
    entry_comando.delete(0, 'end')
    if comando.startswith("delete a coluna"):
        coluna = comando.replace("delete a coluna", '').strip()
        colunas_lower = {col.lower(): col for col in df.columns}
        if coluna.lower() in colunas_lower:
            coluna_real = colunas_lower[coluna.lower()]
            df.drop(columns=[coluna_real], inplace=True)
            atualizar_treeview(comando, df)

        else:
            messagebox.showerror("Erro", f"A coluna '{coluna}' não existe no DataFrame.")

    elif comando.startswith("renomear a coluna"):
        partes = comando.replace("renomear a coluna", '').strip().split(" para ")
        if len(partes) == 2:
            coluna_atual, novo_nome = partes[0].strip(), partes[1].strip()
            colunas_lower = {col.lower(): col for col in df.columns}
            if coluna_atual.lower() in colunas_lower:
                coluna_real = colunas_lower[coluna_atual.lower()]
                df.rename(columns={coluna_real: novo_nome}, inplace=True)
                atualizar_treeview(comando, df)

            else:
                messagebox.showerror("Erro", f"A coluna '{coluna_atual}' não existe no DataFrame.")
    elif comando.startswith("filtrar na coluna"):
        partes = comando.replace("filtrar na coluna", '').strip().split(" pelo valor ")
        if len(partes) == 2:
            coluna, valor = partes[0].strip(), partes[1].strip()
            colunas_lower = {col.lower(): col for col in df.columns}

            if coluna.lower() in colunas_lower:
                coluna_real = colunas_lower[coluna.lower()]
                df[coluna_real] = df[coluna_real].astype(str).str.strip().str.lower()
                valor = valor.strip().lower()
                df_filtrado = df[df[coluna_real] == valor]
                if not df_filtrado.empty:
                    atualizar_treeview(comando, df_filtrado)

                else:
                    messagebox.showwarning("Atenção",
                                           f"Nenhum dado encontrado para '{valor}' na coluna '{coluna_real}'.")

            else:

                messagebox.showerror("Erro", f"A coluna '{coluna}' não existe no DataFrame.")

    elif comando.startswith("ordenar o dataframe pela coluna"):

            coluna = comando.replace("ordenar o dataframe pela coluna", '').strip()
            colunas_lower = {col.lower(): col for col in df.columns}

            if coluna.lower() in colunas_lower:
                coluna_real = colunas_lower[coluna.lower()]
                df.sort_values(by=coluna_real, ascending=True, inplace=True)
                atualizar_treeview(comando, df)


            else:
                messagebox.showerror("Erro", f"A coluna '{coluna}' não existe no DataFrame.")

    elif comando.startswith("preencher valores nulos na coluna"):
        partes = comando.replace("preencher valores nulos na coluna", '').strip().split(" com ")

        if len(partes) == 2:
            coluna, valor = partes[0].strip(), partes[1].strip()
            colunas_lower = {col.lower(): col for col in df.columns}

            if coluna.lower() in colunas_lower:
                coluna_real = colunas_lower[coluna.lower()]
                df[coluna_real].fillna(value=valor, inplace=True)
                atualizar_treeview(comando, df)

            else:
                messagebox.showerror("Erro", f"A coluna '{coluna}' não existe no DataFrame.")

    elif comando.startswith("mostrar as primeiras"):

        try:
            numeros = re.findall(r'\d+', comando)

            if numeros:
                n = int(numeros[0])
                df = df.head(n)  # Atualiza o DataFrame global com apenas as primeiras 'n' linhas.
                atualizar_treeview(comando, df)

            else:
                messagebox.showerror("Erro", "Número de linhas não especificado ou inválido.")

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao mostrar linhas: {e}")

    elif comando.startswith("mostrar as últimas"):

        try:
            numeros = re.findall(r'\d+', comando)

            if numeros:
                n = int(numeros[0])
                df = df.tail(n)
                atualizar_treeview(comando, df)

            else:
                messagebox.showerror("Erro", "Número de linhas não especificado ou inválido.")

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao mostrar linhas: {e}")

    elif comando.startswith("mostrar o") and "que" in comando and "vendeu na coluna de" in comando:

        try:
            padrao = r"mostrar o (.+) que (mais|menos) vendeu na coluna de (.+)"
            match = re.match(padrao, comando)
            if match:
                coluna_grupo = match.group(1).strip()
                mais_ou_menos = match.group(2).strip()
                coluna_vendas = match.group(3).strip()
                colunas_lower = {col.lower().strip(): col.strip() for col in df.columns}
                coluna_grupo_key = coluna_grupo.lower().strip()
                coluna_vendas_key = coluna_vendas.lower().strip()
                if coluna_grupo_key in colunas_lower and coluna_vendas_key in colunas_lower:
                    coluna_grupo_real = colunas_lower[coluna_grupo_key]
                    coluna_vendas_real = colunas_lower[coluna_vendas_key]
                    df[coluna_vendas_real] = pd.to_numeric(df[coluna_vendas_real], errors='coerce')
                    grupo_vendas = df.groupby(coluna_grupo_real)[coluna_vendas_real].sum()
                    if mais_ou_menos == 'mais':
                        grupo_selecionado = grupo_vendas.idxmax()
                        vendas = grupo_vendas.max()

                    else:
                        grupo_selecionado = grupo_vendas[grupo_vendas == grupo_vendas.min()]
                        vendas = grupo_vendas.min()
                    df_grupo = df[df[coluna_grupo_real].isin(grupo_selecionado.index)] if mais_ou_menos == 'menos' else df[df[coluna_grupo_real] == grupo_selecionado]
                    atualizar_treeview(f"{coluna_grupo_real} que {mais_ou_menos} vendeu: {', '.join(grupo_selecionado.index) if mais_ou_menos == 'menos' else grupo_selecionado} (Total de vendas: {vendas})", df_grupo)

                else:
                    colunas_disponiveis = ', '.join(df.columns)
                    messagebox.showerror("Erro", f"Uma das colunas especificadas não existe no DataFrame.\nColunas disponíveis: {colunas_disponiveis}")

            else:
                messagebox.showwarning("Erro", "Comando mal formatado. Tente novamente.")

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao mostrar o {coluna_grupo_real} que {mais_ou_menos} vendeu: {e}")

    elif comando.startswith("mostrar") and "ordenados por" in comando and "na coluna de" in comando:

        try:
            padrao = r"mostrar (.+) ordenados por (.+) na coluna de (.+)"
            match = re.match(padrao, comando)
            if match:
                coluna_grupo = match.group(
                    1).strip()
                coluna_vendas = match.group(
                    3).strip()
                colunas_lower = {col.lower().strip(): col.strip() for col in df.columns}
                coluna_grupo_real = None
                coluna_vendas_real = None

                for col in df.columns:

                    if col.lower().strip() == coluna_grupo.lower().strip():
                        coluna_grupo_real = col
                        break
                for col in df.columns:
                    if col.lower().strip() == coluna_vendas.lower().strip():
                        coluna_vendas_real = col
                        break

                if coluna_grupo_real and coluna_vendas_real:
                    df[coluna_vendas_real] = pd.to_numeric(df[coluna_vendas_real], errors='coerce')
                    grupo_vendas = df.groupby(coluna_grupo_real)[coluna_vendas_real].sum().reset_index()
                    grupo_vendas.sort_values(by=coluna_vendas_real, ascending=False, inplace=True)
                    atualizar_treeview( f"{coluna_grupo_real.capitalize()} ordenados por vendas na coluna {coluna_vendas_real}", grupo_vendas)

                else:
                    colunas_disponiveis = ', '.join(df.columns)
                    messagebox.showerror("Erro", f"Uma das colunas especificadas não existe no DataFrame.\nColunas disponíveis: {colunas_disponiveis}")

            else:
                messagebox.showerror("Erro", "Comando mal formatado. Tente novamente.")

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao mostrar {coluna_grupo_real} ordenados por vendas: {e}")

    else:

        print("Comando não encontrado!")

def mostrar_dicas():
    janela_dicas = tk.Toplevel(janela_principal)
    janela_dicas.title("Dicas de Comandos")
    janela_dicas.geometry("600x400")
    texto_dicas = """Exemplos de Comandos:

1. delete a coluna Meta
2. renomear a coluna Vendedor para Vendedor_Principal
3. filtrar na coluna Meta pelo valor 50000
4. ordenar o DataFrame pela coluna Meta
5. preencher valores nulos na coluna Total de Vendas com 100
6. mostrar as primeiras 10 linhas
7. mostrar as últimas 5 linhas
8. mostrar o Vendedor que mais vendeu na coluna de Total de Vendas
9. mostrar o Produto que menos vendeu na coluna de Total de Vendas
10. mostrar Vendedor ordenados por vendas na coluna de Total de Vendas
11. mostrar Produto ordenados por vendas na coluna de Total de Vendas
"""

    texto = tk.Text(janela_dicas,
                    wrap='word',
                    height=20)
    texto.insert('1.0', texto_dicas)
    texto.pack(expand=True, fill='both')

    def copiar_dicas():
        janela_principal.clipboard_clear()
        janela_principal.clipboard_append(texto_dicas)
        messagebox.showinfo("Copiado", "Dicas copiadas para a área de transferência!")
    btn_copiar = tk.Button(janela_dicas,
                           text="Copiar Dicas",
                           command=copiar_dicas,
                           bg="#4CAF50",
                           fg="white",
                           font=("Arial", 10))
    btn_copiar.pack(pady=10)
janela_principal = tk.Tk()
janela_principal.title("Manipulação de DataFrame com Comandos")
frame_principal = tk.Frame(janela_principal)
frame_principal.pack(pady=10,
                     fill='both',
                     expand=True)
canvas = tk.Canvas(frame_principal,
                   bg="white")
canvas.pack(side="left",
            fill="both",
            expand=True)
scrollbar_chat = ttk.Scrollbar(frame_principal,
                               orient="vertical",
                               command=canvas.yview)
scrollbar_chat.pack(side="right", fill="y")
canvas.configure(yscrollcommand=scrollbar_chat.set)
frame_chat = tk.Frame(canvas, bg="white")
canvas.create_window((0, 0), window=frame_chat, anchor="nw")
def on_frame_configure(event):
    canvas.configure(scrollregion=canvas.bbox("all"))
frame_chat.bind("<Configure>", on_frame_configure)
frame_inferior = tk.Frame(janela_principal, bg="#333")
frame_inferior.pack(side="bottom",
                    fill="x",
                    padx=10,
                    pady=5)
entry_comando = tk.Entry(frame_inferior,
                         width=60,
                         font=("Arial", 12),
                         fg="#333")

entry_comando.pack(side="left", padx=5)
btn_comando = tk.Button(frame_inferior,
                        text="Executar",
                        command=processar_comando,
                        font=("Arial", 12),
                        bg="#4CAF50",
                        fg="white")
btn_comando.pack(side="left", padx=5)
btn_carregar = tk.Button(frame_inferior,
                         text="Carregar Arquivo Excel",
                         command=carregar_arquivo,
                         font=("Arial", 12),
                         bg="#2196F3",
                         fg="white")

btn_carregar.pack(side="left", padx=5)
btn_dicas = tk.Button(frame_inferior,
                      text="Dicas",
                      command=mostrar_dicas,
                      font=("Arial", 12),
                      bg="#FF5722",
                      fg="white")

btn_dicas.pack(side="left", padx=5)

df = pd.DataFrame()

janela_principal.geometry("900x600")
janela_principal.mainloop()