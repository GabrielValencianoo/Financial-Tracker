import tkinter as tk
import ttkbootstrap as tb
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from datetime import datetime
from ofxparse import OfxParser
import os
import json
import random

# Variáveis globais
df_global = None
arquivo_excel = None
tree_widget = None

banco_id = {
    '001': 'Banco do Brasil',
    '033': 'Santander',
    '104': 'Caixa Econômica Federal',
    '237': 'Bradesco',
    '0260': 'Nubank',
    '0341': 'Itau',
    '356': 'Real',
    '399': 'HSBC',
    '422': 'Safra',
    '453': 'Banco Rural',
    '633': 'Banco Rendimento',
    '652': 'Itaú Unibanco Holding S.A.',
    '745': 'Citibank',
    '756': 'Bancoob'
    
}

contas = []
categorias = {}
dict_mapeamento = {}

VALORES_PADRAO = {
    'Conta': 'Desconhecido',
    'Categoria': 'Outros',
    'Subcategoria': 'Despesa desconhecida',
    'Data': datetime.now().strftime("%Y-%m-%d"),
    'Descrição': "Sem descrição",
    'Valor': 0.0,
    'Tipo': "Despesa"
}


def criar_excel_padrao():
    """Cria um DataFrame padrão se não houver arquivo"""
    return pd.DataFrame(columns=['Conta','Categoria','Subcategoria', 'Valor', 'Tipo', 'Descrição','Data'])
    

def carregar_excel():
    """Carrega arquivo Excel"""
    global df_global, arquivo_excel
    
    filename = filedialog.askopenfilename(
        title="Selecione o arquivo Excel",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    
    if filename:
        try:
            df_global = pd.read_excel(filename)
            arquivo_excel = filename
            atualizar_tabela()
            messagebox.showinfo("Sucesso", f"Arquivo carregado: {os.path.basename(filename)}")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar arquivo: {str(e)}")

def criar_novo_excel():
    """Cria um novo arquivo Excel"""
    global df_global, arquivo_excel
    
    filename = filedialog.asksaveasfilename(
        title="Criar novo arquivo Excel",
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")]
    )
    
    if filename:
        df_global = criar_excel_padrao()
        arquivo_excel = filename
        salvar_excel()
        atualizar_tabela()
        messagebox.showinfo("Sucesso", "Novo arquivo Excel criado!")

def salvar_excel():
    """Salva o DataFrame no arquivo Excel"""
    global df_global, arquivo_excel
    
    if df_global is None:
        messagebox.showwarning("Aviso", "Nenhum dado para salvar!")
        return
    
    if arquivo_excel is None:
        arquivo_excel = filedialog.asksaveasfilename(
            title="Salvar arquivo Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
    
    if arquivo_excel:
        try:
            df_global.to_excel(arquivo_excel, index=False)
            messagebox.showinfo("Sucesso", "Arquivo salvo com sucesso!")
        except Exception as e:
            print(str(e))
            messagebox.showerror("Erro", f"Erro ao salvar arquivo: {str(e)}")
            

def atualizar_tabela():
    """Atualiza a visualização da tabela"""
    global tree_widget, df_global
    
    # Limpar tabela
    for item in tree_widget.get_children():
        tree_widget.delete(item)
    
    if df_global is not None and not df_global.empty:
        for idx, row in df_global.iterrows():
            tree_widget.insert('', 'end', values=(idx, *row.values), tags=(df_global.iloc[idx].Conta,))
    
    # print(df_global.head())
    # print(df_global.describe())
    # print(df_global.info())
    print(df_global["Conta"].unique())
    print(df_global.groupby('Conta').size())
    print(df_global.groupby('Categoria').size())
    print(df_global.groupby('Subcategoria').size())
    print(df_global.groupby('Tipo').size())    
    print(df_global["Valor"].sum())
    print(df_global.groupby('Conta')["Valor"].sum().round(2))

def adicionar_registro():
    """Abre janela para adicionar novo registro"""
    global df_global
    
    if df_global is None:
        df_global = criar_excel_padrao()
    
    janela_add = tk.Toplevel()
    janela_add.title("Adicionar Registro")
    janela_add.geometry("700x700+1200+200")
    
    # Campos
    tk.Label(janela_add, text="Conta:").pack(pady=5)
    entry_conta = ttk.Combobox(janela_add, values=contas, width=30,height = 50)
    entry_conta.pack()

    tk.Label(janela_add, text="Categoria:").pack(pady=5)
    entry_categoria = ttk.Combobox(janela_add, values=list(categorias.keys()), width=30,height = 50)
    entry_categoria.pack()   

    tk.Label(janela_add, text="Subcategoria:").pack(pady=5)
    entry_subcategoria = ttk.Combobox(janela_add, postcommand=lambda: entry_subcategoria.config(values=categorias[entry_categoria.get()]), width=30,height = 50)
    entry_subcategoria.pack() 

    tk.Label(janela_add, text="Valor:").pack(pady=5)
    entry_valor = tk.Entry(janela_add, width=30)
    entry_valor.pack()

    tk.Label(janela_add, text="Tipo:").pack(pady=5)
    combo_tipo = ttk.Combobox(janela_add, values=["Receita", "Despesa"], width=28)
    combo_tipo.pack()

    tk.Label(janela_add, text="Descrição:").pack(pady=5)
    entry_desc = tk.Entry(janela_add, width=30)
    entry_desc.pack()
    
    tk.Label(janela_add, text="Data (AAAA-MM-DD):").pack(pady=5)
    entry_data = tb.DateEntry(janela_add, dateformat = "%Y-%m-%d",width=30)
    entry_data.pack()

    
    def salvar_novo():
        try:
            nova_linha = {
                'Conta': entry_conta.get(),
                'Categoria': entry_categoria.get(),
                'Subcategoria': entry_subcategoria.get(),
                'Data': entry_data.get_date().strftime("%Y-%m-%d"),
                'Descrição': entry_desc.get(),
                'Valor': float(entry_valor.get()),
                'Tipo': combo_tipo.get()
            }
            df_global.loc[len(df_global)] = nova_linha
            atualizar_tabela()
            janela_add.destroy()
            messagebox.showinfo("Sucesso", "Registro adicionado!")
        except ValueError:
            messagebox.showerror("Erro", "Valor inválido! Use apenas números.")
    
    tk.Button(janela_add, text="Salvar", command=salvar_novo, bg="green", fg="white").pack(pady=20)

def atualizar_registro(event):
    """Atualiza o registro selecionado"""
    global df_global, tree_widget
    
    selecionado = tree_widget.selection()
    if not selecionado:
        messagebox.showwarning("Aviso", "Selecione um registro para atualizar!")
        return
    
    item = tree_widget.item(selecionado[0])
    valores = item['values']
    idx = valores[0]
    
    janela_edit = tk.Toplevel()
    janela_edit.title("Atualizar Registro")
    janela_edit.geometry("700x700+1200+200")

    tk.Label(janela_edit, text="Conta:").pack(pady=5)
    entry_conta = ttk.Combobox(janela_edit, values=contas, width=30,height = 50)
    entry_conta.set(valores[1])
    entry_conta.pack()

    tk.Label(janela_edit, text="Categoria:").pack(pady=5)
    entry_categoria = ttk.Combobox(janela_edit, values=list(categorias.keys()), width=30,height = 50)
    entry_categoria.set(valores[2])
    entry_categoria.pack()   

    tk.Label(janela_edit, text="Subcategoria:").pack(pady=5)
    entry_subcategoria = ttk.Combobox(janela_edit, values=categorias[entry_categoria.get()], postcommand=lambda: entry_subcategoria.config(values=categorias[entry_categoria.get()]), width=30,height = 50)
    entry_subcategoria.set(valores[3])
    entry_subcategoria.pack() 

    tk.Label(janela_edit, text="Valor:").pack(pady=5)
    entry_valor = tk.Entry(janela_edit, width=30)
    entry_valor.insert(0, valores[4])
    entry_valor.pack()

    tk.Label(janela_edit, text="Tipo:").pack(pady=5)
    combo_tipo = ttk.Combobox(janela_edit, values=["Receita", "Despesa"], width=28)
    combo_tipo.set(valores[5])
    combo_tipo.pack()

    tk.Label(janela_edit, text="Descrição:").pack(pady=5)
    entry_desc = tk.Entry(janela_edit, width=30)
    entry_desc.insert(0, valores[6])
    entry_desc.pack()
    
    tk.Label(janela_edit, text="Data (AAAA-MM-DD):").pack(pady=5)
    entry_data = tb.DateEntry(janela_edit,dateformat = "%Y-%m-%d", width=30)
    entry_data.set_date(datetime.strptime(valores[7], "%Y-%m-%d"))
    entry_data.pack()
    
    
    def salvar_alteracao():
        try:
            df_global.at[idx, 'Conta'] = entry_conta.get()
            df_global.at[idx, 'Categoria'] = entry_categoria.get()
            df_global.at[idx, 'Subcategoria'] = entry_subcategoria.get()
            df_global.at[idx, 'Data'] = entry_data.get_date().strftime("%Y-%m-%d")
            df_global.at[idx, 'Descrição'] = entry_desc.get()
            df_global.at[idx, 'Valor'] = float(entry_valor.get())
            df_global.at[idx, 'Tipo'] = combo_tipo.get()
            atualizar_tabela()
            janela_edit.destroy()
            messagebox.showinfo("Sucesso", "Registro atualizado!")
        except ValueError:
            messagebox.showerror("Erro", "Valor inválido!")
    
    tk.Button(janela_edit, text="Salvar", command=salvar_alteracao, bg="blue", fg="white").pack(pady=20)


def duplicar_registro():
    """Duplica o registro selecionado"""
    global df_global, tree_widget
    
    selecionado = tree_widget.selection()
    if not selecionado:
        messagebox.showwarning("Aviso", "Selecione um registro para duplicar!")
        return
    
    item = tree_widget.item(selecionado[0])
    valores = item['values']
    idx = valores[0]
    
    nova_linha = df_global.loc[idx].copy()
    df_global.loc[len(df_global)] = nova_linha
    atualizar_tabela()
    messagebox.showinfo("Sucesso", "Registro duplicado!")

def deletar_registro():
    """Deleta o registro selecionado"""
    global df_global, tree_widget
    
    selecionado = tree_widget.selection()
    if not selecionado:
        messagebox.showwarning("Aviso", "Selecione um registro para deletar!")
        return
    
    resposta = messagebox.askyesno("Confirmar", "Deseja realmente deletar este registro?")
    if resposta:
        item = tree_widget.item(selecionado[0])
        idx = item['values'][0]
        df_global = df_global.drop(index=idx).reset_index(drop=True)
        atualizar_tabela()
        messagebox.showinfo("Sucesso", "Registro deletado!")

def deletar_tabela():
    """Deleta todos os registros da tabela"""
    global df_global
    
    if df_global is None or df_global.empty:
        messagebox.showwarning("Aviso", "Nenhum registro para deletar!")
        return
    
    resposta = messagebox.askyesno("Confirmar", "Deseja realmente deletar TODOS os registros?")
    if resposta:
        df_global = criar_excel_padrao()
        atualizar_tabela()
        messagebox.showinfo("Sucesso", "Todos os registros foram deletados!")



def importar_csv():
    """Importa dados de um arquivo CSV"""
    global df_global,df_importado
    
    if df_global is None:
        df_global = criar_excel_padrao()
    
    filename = filedialog.askopenfilename(
        title="Selecione o arquivo CSV",
        filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
    )


    print(filename)
    if filename is None or filename == "":
        messagebox.showwarning("Aviso", "Nenhum arquivo selecionado!")
        return
    
    try:
        df_importado = pd.read_csv(filename)
        colunas_csv = df_importado.columns.tolist()
        janela_csv = tk.Toplevel()
        janela_csv.title("Importar CSV")
        janela_csv.geometry("700x700+1200+200")
        for col_prog in df_global.columns:
            lbl = tk.Label(janela_csv, text=f"{col_prog}:", font=('Arial', 10, 'bold'))
            lbl.pack(anchor="w", pady=(5, 0))

            if col_prog == "Conta":
                combo = ttk.Combobox(janela_csv, values=["IGNORAR"] + colunas_csv, state="readonly", width=30)
            # elif col_prog == "Categoria":
            #     combo = ttk.Combobox(janela_csv, values=["IGNORAR"] + colunas_csv + list(categorias.keys()), state="readonly", width=30)
            else:
                combo = ttk.Combobox(janela_csv, values=["IGNORAR"] + colunas_csv, state="readonly", width=30)
            combo.pack(fill="x", pady=2)
            dict_mapeamento[col_prog] = combo
        


    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao importar CSV: {str(e)}")

    def realizar_concatenacao():
        """Executa o de/para e concatena os dataframes removendo NaNs."""
        global df_importado,df_global
        
        try:
            
            df_novo = pd.DataFrame("", index=range(len(df_importado)), columns=df_global.columns)

            for col_prog, combo in dict_mapeamento.items():
                escolha = combo.get()

                if escolha == "IGNORAR":
                    # Cria uma coluna com o valor padrão repetido para todas as linhas do CSV
                    valor_fixo = VALORES_PADRAO.get(col_prog) # Usa "" se não achar no dicionário
                    df_novo[col_prog] = valor_fixo
                else:
                    # Mapeia a coluna do CSV para o nome da sua coluna original
                    df_novo[col_prog] = df_importado[escolha].values
                

            print(df_novo.head())
            # 1. Concatena o df original com o novo processado
            df_final = pd.concat([df_global, df_novo], ignore_index=True)
            
            # 2. A MÁGICA: Substitui todos os valores NaN por string vazia ''
            df_global = df_final.fillna('')

            atualizar_tabela()
            janela_csv.destroy()
            messagebox.showinfo("Sucesso", f"{len(df_importado)} registros importados do CSV!")
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro no processamento: {e}")
    
    btn_processar = tk.Button(janela_csv, text="3. Concatenar e Finalizar", 
                    command=realizar_concatenacao, bg="#4caf50", fg="white", font=('Arial', 10, 'bold'))
    btn_processar.pack(pady=10, padx=20, fill="x")

def importar_ofx():
    """Importa dados de arquivo OFX usando ofxparse"""
    global df_global
    
    if df_global is None:
        df_global = criar_excel_padrao()
    
    filename = filedialog.askopenfilename(
        title="Selecione o arquivo OFX",
        filetypes=[("OFX files", "*.ofx"), ("All files", "*.*")]
    )
    
    if filename:
        try:
            # Abre e parseia o arquivo OFX
            with open(filename, 'rb') as ofx_file:
                ofx = OfxParser.parse(ofx_file)
            
            transacoes_adicionadas = 0
            
            # Percorre todas as contas no arquivo OFX
            for account in ofx.accounts:
                # Percorre todas as transações da conta
                for transaction in account.statement.transactions:
                    # Formata a data
                    data_formatada = transaction.date.strftime("%Y-%m-%d")
                    
                    # Obtém descrição (memo ou payee)
                    descricao = transaction.memo or transaction.payee or "Sem descrição"
                    
                    # Valor e tipo
                    valor = float(transaction.amount)
                    
                    tipo = "Receita" if valor > 0 else "Despesa"

                    try:
                        conta = banco_id[account.routing_number]
                    except KeyError:
                        conta = "Desconhecido"


                    
                    
                    nova_linha = {
                        'Conta': conta,
                        'Categoria': 'Outros',
                        'Subcategoria': 'Despesa desconhecida',
                        'Data': data_formatada,
                        'Descrição': descricao,
                        'Valor': valor,
                        'Tipo': tipo                       
                    }
                    
                    df_global.loc[len(df_global)] = nova_linha
                    transacoes_adicionadas += 1
            
            atualizar_tabela()
            messagebox.showinfo("Sucesso", f"{transacoes_adicionadas} transações importadas do arquivo OFX!")
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao importar OFX: {str(e)}")


def read_txt_settings():
    """Lê configurações de um arquivo .txt (exemplo para futuras customizações)"""
    global contas, categorias
    jsonSettings = {}
    try:
        with open("Financial_Settings.json", "r", encoding="utf-8") as file:
            jsonSettings = json.load(file)
    except FileNotFoundError:
        print("Arquivo de configuração não encontrado. Usando configurações padrão.")
    except Exception as e:
        print(f"Erro ao ler configurações: {str(e)}")
    

    contas = jsonSettings.get("contas", [])
    categorias = jsonSettings.get("categorias", {})

def criar_interface():
    """Cria a interface principal"""
    global tree_widget
    
    root = tk.Tk()
    root.title("Gerenciador de Finanças - Excel e OFX")
    root.geometry("1500x600+1200+200")
    
    # Frame superior - Botões
    frame_botoes = tk.Frame(root)
    frame_botoes.pack(pady=10)
    
    tk.Button(frame_botoes, text="Novo Excel", command=criar_novo_excel, 
              bg="#4CAF50", fg="white", width=12).grid(row=0, column=0, padx=5)
    tk.Button(frame_botoes, text="Abrir Excel", command=carregar_excel, 
              bg="#2196F3", fg="white", width=12).grid(row=0, column=1, padx=5)
    tk.Button(frame_botoes, text="Salvar Excel", command=salvar_excel, 
              bg="#FF9800", fg="white", width=12).grid(row=0, column=2, padx=5)
    tk.Button(frame_botoes, text="Importar OFX", command=importar_ofx, 
              bg="#9C27B0", fg="white", width=12).grid(row=0, column=3, padx=5)
    tk.Button(frame_botoes, text="Importar CSV", command=importar_csv, 
              bg="#009688", fg="white", width=12).grid(row=0, column=4, padx=5)
    
    # Frame CRUD
    frame_crud = tk.Frame(root)
    frame_crud.pack(pady=10)
    
    tk.Button(frame_crud, text="Adicionar", command=adicionar_registro, 
              bg="#4CAF50", fg="white", width=12).grid(row=0, column=0, padx=5)
    tk.Button(frame_crud, text="Atualizar", command= lambda: atualizar_registro(None), 
              bg="#2196F3", fg="white", width=12).grid(row=0, column=1, padx=5)
    tk.Button(frame_crud, text="Deletar", command=deletar_registro, 
              bg="#F44336", fg="white", width=12).grid(row=0, column=2, padx=5)
    tk.Button(frame_crud, text="Deletar Tudo", command=deletar_tabela, 
              bg="#F44336", fg="white", width=12).grid(row=0, column=3, padx=5)  
    tk.Button(frame_crud, text="Duplicar", command=duplicar_registro, 
              bg="#FFC107", fg="white", width=12).grid(row=0, column=4, padx=5)  
    

    # Frame da tabela
    frame_tabela = tk.Frame(root)
    frame_tabela.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
    
    # Scrollbar
    scrollbar = ttk.Scrollbar(frame_tabela)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    
    # Treeview
    colunas =('ID','Conta','Categoria','Subcategoria', 'Valor', 'Tipo', 'Descrição','Data')
    tree_widget = ttk.Treeview(frame_tabela, columns=colunas, show='headings', 
                               yscrollcommand=scrollbar.set)

    tree_widget.heading('ID', text='ID')
    tree_widget.heading('Conta', text='Conta')
    tree_widget.heading('Categoria', text='Categoria')
    tree_widget.heading('Subcategoria', text='Subcategoria')
    tree_widget.heading('Valor', text='Valor')
    tree_widget.heading('Tipo', text='Tipo')
    tree_widget.heading('Descrição', text='Descrição')
    tree_widget.heading('Data', text='Data')

    tree_widget.column('ID', width=50)
    tree_widget.column('Conta', width=100)
    tree_widget.column('Categoria', width=100)
    tree_widget.column('Subcategoria', width=100)
    tree_widget.column('Valor', width=100)
    tree_widget.column('Tipo', width=100)
    tree_widget.column('Descrição', width=400)
    tree_widget.column('Data', width=100)

    tree_widget.pack(fill=tk.BOTH, expand=True)
    scrollbar.config(command=tree_widget.yview)

    tree_widget.bind('<Double-1>', atualizar_registro )
    read_txt_settings()

    for conta in contas:
        tree_widget.tag_configure(conta, background=f"#{random.randint(0, 0xFFFFFF):06x}", foreground=f"#{random.randint(0, 0xFFFFFF):06x}")
        
    # Create a tag named 'colored_row' with a red background
    
    
    root.mainloop()

if __name__ == "__main__":
    criar_interface()