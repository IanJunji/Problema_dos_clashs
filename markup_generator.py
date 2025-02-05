import tkinter as tk
from tkinter import filedialog, messagebox
import os
import openpyxl
from openpyxl import load_workbook

# Variáveis globais para armazenar os caminhos dos arquivos/diretórios
clash_file_path = None
matrix_file_path = None
output_dir = None

def select_clash_file():
    filename = filedialog.askopenfilename(filetypes=[("Text files", "*.txt")])
    if filename:
        clash_file_path.set(filename)

def select_matrix_file():
    filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if filename:
        matrix_file_path.set(filename)

def select_output_dir():
    dirname = filedialog.askdirectory()
    if dirname:
        output_dir.set(dirname)

def process_clash_file(filepath):
    clashs = []
    current_clash = {}
    
    with open(filepath, 'r', encoding='utf-8') as arquivo:
        linhas = arquivo.readlines()
    
    for linha in linhas:
        linha = linha.strip()
        
        if linha.startswith('Name:'):
            if current_clash:  # Se já houver um clash em progresso, adiciona-o à lista
                clashs.append(current_clash)
            current_clash = {}  # Inicia um novo clash
            _, valor = linha.split(':', 1)
            current_clash['name'] = valor.strip()
        
        elif linha.startswith('Image Location:'):
            _, valor = linha.split(':', 1)
            _, nome = valor.split('-', 1)
            nome_objs, _ = nome.split('-', 1)
            obj1, obj2 = nome_objs.split('X', 1)
            _, pre_id = linha.split('\\', 1)
            id_val, _ = pre_id.split('.', 1)
            current_clash['objetos'] = nome_objs.strip()
            current_clash['image_loc'] = valor.strip()
            current_clash['obj_1'] = obj1.strip()
            current_clash['obj_2'] = obj2.strip()
            current_clash['id'] = id_val.strip()

        elif linha.startswith('HardStatus:'):
            _, valor = linha.split(':', 1)
            current_clash['HardStatus'] = valor.strip()
        
        elif linha.startswith('Clash Point:'):
            _, valor = linha.split(':', 1)
            cord_sem_m = valor.replace('m', '').strip()
            x, resto = cord_sem_m.split(',', 1)
            y, z = resto.split(',', 1)
            current_clash['coord_x'] = x.strip()
            current_clash['coord_y'] = y.strip()
            current_clash['coord_z'] = z.strip()
            current_clash['coordinates'] = cord_sem_m
        
        elif linha.startswith('Date Created:'):
            _, valor = linha.split(':', 1)
            current_clash['criacao'] = valor.strip()
        
        elif linha.startswith('Entity Handle:'):
            if 'entity_1' not in current_clash:
                _, valor = linha.split(':', 1)
                current_clash['entity_1'] = valor.strip()
            else:
                _, valor = linha.split(':', 1)
                current_clash['entity_2'] = valor.strip()
        
        elif linha.startswith('Layer:'):
            if 'layer_1' not in current_clash:
                _, valor = linha.split(':', 1)
                current_clash['layer_1'] = valor.strip()
            else:
                _, valor = linha.split(':', 1)
                current_clash['layer_2'] = valor.strip()

    if current_clash:
        clashs.append(current_clash)
    
    return clashs

def process_matrix(clashs, matrix_path):
    # Carrega a planilha e as abas
    planilha_matriz = openpyxl.load_workbook(matrix_path)
    aba_matriz = planilha_matriz['Matriz']
    aba_listagem = planilha_matriz['Listagem']
    clashs_aprovados = []

    # Dicionários para armazenar os valores de classificação e mapeamento para o nome da coluna
    classification_values = {
        "Topografia": {"Pontual": set(), "X": set(), "Y": set()},
        "Drenagem": {"Linear": set(), "Enterrado": set(), "Pontual": set()},
        "Interferências": {"Enterrada": set(), "Aerea": set(), "Pontual": set()},
        "Sinalização Vertical": {"Pontual": set(), "X": set(), "Y": set()},
        "Dispositivo de Segurança": {"Linear": set(), "X": set(), "Y": set()},
        "OAEs": {"Linear": set(), "X": set(), "Y": set()},
        "Contenções": {"Paramento": set(), "Grampo": set(), "Dreno": set()},
        "Pavimentação": {"Camadas": set(), "X": set(), "Y": set()},
        "Iluminação": {"Engastado": set(), "X": set(), "Y": set()},
        "Estruturas Especiais": {"X": set(), "Y": set(), "Z": set()},
    }

    column_names = {
        "Topografia": {"Pontual": "topografia_a", "X": "topografia_b", "Y": "topografia_c"},
        "Drenagem": {"Linear": "drenagem_a", "Enterrado": "drenagem_b", "Pontual": "drenagem_c"},
        "Interferências": {"Enterrada": "interferencia_a", "Aerea": "interferencia_b", "Pontual": "interferencia_c"},
        "Sinalização Vertical": {"Pontual": "sinalizacao_a", "X": "sinalizacao_b", "Y": "sinalizacao_c"},
        "Dispositivo de Segurança": {"Linear": "disp_seg_a", "X": "disp_seg_b", "Y": "disp_seg_c"},
        "OAEs": {"Linear": "oae_a", "X": "oae_b", "Y": "oae_c"},
        "Contenções": {"Paramento": "contencoes_a", "Grampo": "contencoes_b", "Dreno": "contencoes_c"},
        "Pavimentação": {"Camadas": "pavimentacao_a", "X": "pavimentacao_b", "Y": "pavimentacao_c"},
        "Iluminação": {"Engastado": "iluminacao_a", "X": "iluminacao_b", "Y": "iluminacao_c"},
        "Estruturas Especiais": {"X": "est_esp_a", "Y": "est_esp_b", "Z": "est_esp_c"},
    }

    # Preenche o dicionário classification_values com os dados da aba "Listagem"
    # As colunas: 1 => disciplina, 2 => classificação e 3 => valor
    for row in range(2, aba_listagem.max_row + 1):
        discipline = aba_listagem.cell(row=row, column=1).value
        if discipline in classification_values:
            criteria = aba_listagem.cell(row=row, column=2).value
            value = aba_listagem.cell(row=row, column=3).value
            if criteria in classification_values[discipline]:
                classification_values[discipline][criteria].add(value)

    # Pré-computa um dicionário para os cabeçalhos da aba "Matriz" (linha 2, colunas a partir da 3)
    matriz_header = {}
    for col in aba_matriz.iter_cols(min_row=2, max_row=2, min_col=3):
        for cel in col:
            matriz_header[cel.value] = cel.column

    # Pré-computa um dicionário para mapeamento das linhas da aba "Matriz" (coluna 2, a partir da linha 3)
    matriz_rows = {}
    for col in aba_matriz.iter_cols(min_col=2, max_col=2, min_row=3):
        for cel in col:
            matriz_rows[cel.value] = cel.row

    # Processa cada clash e determina se é aprovado (célula na intersecção diferente de 'O')
    for clash in clashs:
        discipline = clash.get('obj_1', None)
        obj2 = clash.get('obj_2', None)
        layer1 = clash.get('layer_1', None)
        if discipline in classification_values and layer1 is not None:
            nome_coluna = None
            # Percorre cada critério para a disciplina e verifica se layer1 está nos valores permitidos
            for criteria, allowed_values in classification_values[discipline].items():
                if layer1 in allowed_values:
                    nome_coluna = column_names[discipline][criteria]
                    break

            if nome_coluna is None:
                continue  # Não foi possível determinar a coluna para este clash

            # Obtém o índice da coluna e da linha utilizando os mapeamentos pré-computados
            coluna = matriz_header.get(nome_coluna, None)
            linha = matriz_rows.get(obj2, None)

            if linha is not None and coluna is not None:
                if aba_matriz.cell(row=linha, column=coluna).value != 'O':
                    clashs_aprovados.append(clash)

    return clashs_aprovados

def criar_txts_por_disciplina(clashs_aprovados, diretorio_saida):
    os.makedirs(diretorio_saida, exist_ok=True)
    
    for clash in clashs_aprovados:
        nome = clash.get('name', '')
        objetos = clash.get('objetos', '')
        id_clash = clash.get('id', '')
        entity1 = clash.get('entity_1', '')
        entity2 = clash.get('entity_2', '')
        layer1 = clash.get('layer_1', '')
        layer2 = clash.get('layer_2', '')
        coord_x = clash.get('coord_x', '')
        coord_y = clash.get('coord_y', '')
        coord_z = clash.get('coord_z', '')
        
        conteudo = f"X: {coord_x}\n" \
                   f"Y: {coord_y}\n" \
                   f"Z: {coord_z}\n" \
                   f"Objetos: {objetos}\n" \
                   f"ID: {id_clash}\n" \
                   f"Entity1: {entity1}\n" \
                   f"Entity2: {entity2}\n" \
                   f"Layer1: {layer1}\n" \
                   f"Layer2: {layer2}\n" \
                   f"{'-'*40}\n"
        
        # Cria/atualiza os TXT para as disciplinas
        disciplinas = [clash.get('obj_1', ''), clash.get('obj_2', '')]
        
        for disciplina in disciplinas:
            if disciplina:
                caminho_txt = os.path.join(diretorio_saida, f"{disciplina}.txt")
                with open(caminho_txt, 'a', encoding='utf-8') as txt_file:
                    txt_file.write(conteudo)

def process_files():
    if not (clash_file_path.get() and matrix_file_path.get() and output_dir.get()):
        messagebox.showerror("Erro", "Por favor, selecione todos os arquivos necessários.")
        return
    
    try:
        clashs = process_clash_file(clash_file_path.get())
        clashs_aprovados = process_matrix(clashs, matrix_file_path.get())
        criar_txts_por_disciplina(clashs_aprovados, output_dir.get())
        messagebox.showinfo("Sucesso", "Processamento concluído com sucesso!")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro durante o processamento: {str(e)}")

def create_gui():
    global clash_file_path, matrix_file_path, output_dir
    root = tk.Tk()
    root.title("Clash Analyzer")
    root.geometry("600x400")
    
    # Inicializa as variáveis globais como StringVar's
    clash_file_path = tk.StringVar()
    matrix_file_path = tk.StringVar()
    output_dir = tk.StringVar()
    
    # Seleção do Arquivo de Clash
    tk.Label(root, text="Arquivo de Clash:").pack(pady=5)
    tk.Entry(root, textvariable=clash_file_path, width=50).pack(pady=5)
    tk.Button(root, text="Selecionar Arquivo", command=select_clash_file).pack(pady=5)
    
    # Seleção do Arquivo da Matriz
    tk.Label(root, text="Arquivo da Matriz:").pack(pady=5)
    tk.Entry(root, textvariable=matrix_file_path, width=50).pack(pady=5)
    tk.Button(root, text="Selecionar Matriz", command=select_matrix_file).pack(pady=5)
    
    # Seleção do Diretório de Saída
    tk.Label(root, text="Diretório de Saída:").pack(pady=5)
    tk.Entry(root, textvariable=output_dir, width=50).pack(pady=5)
    tk.Button(root, text="Selecionar Diretório", command=select_output_dir).pack(pady=5)
    
    # Botão para Processar os Arquivos
    tk.Button(root, text="Processar", command=process_files).pack(pady=20)
    
    root.mainloop()

if __name__ == "__main__":
    create_gui()