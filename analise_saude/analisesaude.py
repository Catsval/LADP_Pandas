import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox

df = None  # Variável global para armazenar o dataframe

def carregar_dados():
    global df
    file_path = 'C:/Users/catav/Downloads/analise_saude/ESaude2022.xlsx'
    if file_path:
        try:
            # Especificar o motor openpyxl para arquivos .xlsx
            df_q1_1 = pd.read_excel(file_path, sheet_name='Q1.1', engine='openpyxl')
            df_q1_2 = pd.read_excel(file_path, sheet_name='Q1.2', engine='openpyxl')

            # Limpar e preparar os dados
            def clean_dataframe(df):
                df_cleaned = df.drop([0, 2, 3, 4])
                df_cleaned.columns = df_cleaned.iloc[0]
                df_cleaned = df_cleaned.drop(1).reset_index(drop=True)
                return df_cleaned

            df_q1_1_cleaned = clean_dataframe(df_q1_1)
            df_q1_2_cleaned = clean_dataframe(df_q1_2)

            # Mesclar os dois dataframes
            df = pd.concat([df_q1_1_cleaned, df_q1_2_cleaned])

            # Converter os anos para strings
            df.columns = df.columns.astype(str)

            # Salvar o dataframe mesclado em um novo arquivo Excel
            merged_file_path = 'C:/Users/catav/Downloads/analise_saude/ESaude2022_merged.xlsx'
            df.to_excel(merged_file_path, index=False)

            messagebox.showinfo("Sucesso", f"Dados carregados, limpos e salvos em {merged_file_path} com sucesso!")

            # Exibir as categorias, subcategorias, subcategorias 2 e detalhes disponíveis
            categorias = df['Categoria'].unique().tolist()
            subcategorias = df['Subcategoria'].unique().tolist()
            subcategorias_2 = df['Subcategoria 2'].unique().tolist()
            detalhes = df['Detalhe'].unique().tolist()

            combobox_categoria['values'] = categorias
            combobox_subcategoria['values'] = subcategorias
            combobox_subcategoria_2['values'] = subcategorias_2
            combobox_detalhe['values'] = detalhes

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar dados: {e}")

def atualizar_subcategorias(event=None):
    if df is not None:
        categoria = combobox_categoria.get()
        subcategorias = df[df['Categoria'] == categoria]['Subcategoria'].unique().tolist()
        combobox_subcategoria['values'] = subcategorias
        combobox_subcategoria.set('')
        combobox_subcategoria_2.set('')
        combobox_detalhe.set('')
        atualizar_subcategorias_2()

def atualizar_subcategorias_2(event=None):
    if df is not None:
        categoria = combobox_categoria.get()
        subcategoria = combobox_subcategoria.get()
        subcategorias_2 = df[(df['Categoria'] == categoria) & (df['Subcategoria'] == subcategoria)]['Subcategoria 2'].dropna().unique().tolist()
        detalhes = df[(df['Categoria'] == categoria) & (df['Subcategoria'] == subcategoria)]['Detalhe'].dropna().unique().tolist()
        combobox_subcategoria_2['values'] = subcategorias_2
        combobox_detalhe['values'] = detalhes

def calcular_estatisticas(df, operacao):
    try:
        categoria = combobox_categoria.get()
        subcategoria = combobox_subcategoria.get()
        subcategoria_2 = combobox_subcategoria_2.get()
        detalhe = combobox_detalhe.get()
        ano = entry_ano.get()

        # Aplicar os filtros opcionais
        if categoria:
            df = df[df['Categoria'] == categoria]
        if subcategoria:
            df = df[df['Subcategoria'] == subcategoria]
        if subcategoria_2:
            df = df[df['Subcategoria 2'] == subcategoria_2]
        if detalhe:
            df = df[df['Detalhe'] == detalhe]
        if ano:
            if ano in df.columns:
                df = df[['Categoria', 'Subcategoria', 'Subcategoria 2', 'Detalhe', ano]]
            else:
                messagebox.showerror("Erro", f"Ano {ano} não encontrado nos dados")
                return

        if operacao == 'Mostrar Dados Brutos':
            exibir_dados_brutos(df)
            return

        dados_filtrados = df.iloc[:, 4:].apply(pd.to_numeric, errors='coerce')

        if operacao == 'Média':
            resultado = dados_filtrados.mean()
        elif operacao == 'Mediana':
            resultado = dados_filtrados.median()
        elif operacao == 'Moda':
            resultado = dados_filtrados.mode().iloc[0]

        exibir_resultados(resultado)
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao calcular estatísticas: {e}")

def exibir_dados_brutos(dados):
    window = tk.Toplevel(root)
    window.title("Dados Brutos")

    text = tk.Text(window)
    text.pack(fill=tk.BOTH, expand=True)
    text.insert(tk.END, dados.to_string(index=False))

def exibir_resultados(resultado):
    window = tk.Toplevel(root)
    window.title("Resultados")
    text = tk.Text(window)
    text.pack(fill=tk.BOTH, expand=True)
    
    # Converter os índices de ano para strings
    resultado.index = resultado.index.astype(str)
    text.insert(tk.END, resultado.to_string())

def analisar_dados():
    if df is not None:
        escolha = combobox_operacao.get()
        calcular_estatisticas(df, escolha)
    else:
        messagebox.showwarning("Aviso", "Carregue os dados primeiro.")

root = tk.Tk()
root.title("Análise de Dados de Saúde")

frame = ttk.Frame(root, padding="10")
frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

ttk.Button(frame, text="Carregar Dados", command=carregar_dados).grid(row=0, column=0, pady=5)

ttk.Label(frame, text="Selecione a operação:").grid(row=1, column=0, pady=5)
combobox_operacao = ttk.Combobox(frame, values=["Mostrar Dados Brutos", "Média", "Mediana", "Moda"])
combobox_operacao.grid(row=2, column=0, pady=5)
combobox_operacao.current(0)

ttk.Label(frame, text="Categoria:").grid(row=3, column=0, pady=5)
combobox_categoria = ttk.Combobox(frame)
combobox_categoria.grid(row=4, column=0, pady=5)
combobox_categoria.bind("<<ComboboxSelected>>", atualizar_subcategorias)

ttk.Label(frame, text="Subcategoria:").grid(row=5, column=0, pady=5)
combobox_subcategoria = ttk.Combobox(frame)
combobox_subcategoria.grid(row=6, column=0, pady=5)
combobox_subcategoria.bind("<<ComboboxSelected>>", atualizar_subcategorias_2)

ttk.Label(frame, text="Subcategoria 2:").grid(row=7, column=0, pady=5)
combobox_subcategoria_2 = ttk.Combobox(frame)
combobox_subcategoria_2.grid(row=8, column=0, pady=5)

ttk.Label(frame, text="Detalhe:").grid(row=9, column=0, pady=5)
combobox_detalhe = ttk.Combobox(frame)
combobox_detalhe.grid(row=10, column=0, pady=5)

ttk.Label(frame, text="Ano:").grid(row=11, column=0, pady=5)
entry_ano = ttk.Entry(frame)
entry_ano.grid(row=12, column=0, pady=5)

ttk.Button(frame, text="Analisar Dados", command=analisar_dados).grid(row=13, column=0, pady=5)

root.mainloop()
