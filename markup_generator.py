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

def extract_disciplina(linha):
    """
    Extrai a disciplina a partir de uma linha do caminho (linha que começa com "Path:").
    Realiza uma série de splits conforme o padrão esperado da string e retorna a disciplina
    associada, ou None caso não consiga extrair.
    """
    try:
        # Primeiramente ignora as partes antes dos '>':
        _, rsp = linha.split('>', 1)
        _, rsp2 = rsp.split('>', 1)
        
        # Processa os splits consecutivos conforme o padrão:
        temp = rsp2
        _, temp = temp.split('-', 1)      # descarta a primeira parte (valrsp)
        _, temp = temp.split('-', 1)      # descarta a segunda parte (valrj)
        _, temp = temp.split('-', 1)      # descarta a terceira parte (val218)
        _, temp = temp.split('-', 1)      # descarta a quarta parte (val226)
        _, temp = temp.split('-', 1)      # descarta a quinta parte (valaca)
        _, temp = temp.split('-', 1)      # descarta a sexta parte (valexe)
        _, temp = temp.split('-', 1)      # descarta a sétima parte (valmb), 
        # Agora temp contém "valsemifinal-numdps-resto..."
        valfinal, temp = temp.split('-', 1)
        numdps, _ = temp.split('-', 1)
        
        # Limpa os espaços em branco
        valfinal = valfinal.strip()
        numdps = numdps.strip()
        
        # Mapeamento básico das siglas para as disciplinas
        disciplinas = {
            'C1': 'Topografia',
            'F1': 'Geometria',
            'G1': 'Terraplenagem',
            'H2': 'Drenagem',
            'J2': 'Dispositivo de Segurança',
            'I2': 'Pavimentação',
            'L2': 'OAEs',
            'K2': 'Iluminação',
            'L4': 'Geotecnia',
            'M1': 'Interferências',
            'Q1': 'Desapropriação',
            'N2': 'Paisagismo',
            'Z9': 'Geral'
        }

        # Tratamento especial para 'J1'
        if valfinal == 'J1' and numdps == '001':
            return 'Sinalização Vertical'
        
        return disciplinas.get(valfinal)
    except Exception as e:
        print(f"Erro durante a extração da disciplina: {e}")
        return None

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

        elif linha.startswith('Path:'):
            disciplina = extract_disciplina(linha)
            if disciplina:
                key = 'disciplina_1' if 'disciplina_1' not in current_clash else 'disciplina_2'
                current_clash[key] = disciplina

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
    workbook = openpyxl.load_workbook(matrix_path)
    aba_matriz = workbook['Matriz']
    aba_listagem = workbook['Listagem']

    # Inicialize a lista que armazenará os clashs aprovados
    clashs_semi_aprovados = []

    for clash in clashs:
        for col in aba_matriz.iter_cols(min_row=2, max_row=2, min_col=4):
            for cel in col:
                if cel.value == clash['disciplina_1']:
                    coluna = cel.column
        
        for col in aba_matriz.iter_cols(min_col=2, max_col=2, min_row=4):
            for cel in col:
                if cel.value == clash['disciplina_2']:
                    linha = cel.row
        
        if aba_matriz.cell(row=linha, column=coluna).value == 'O':
            clashs_semi_aprovados.append(clash)
    return clashs_semi_aprovados

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
        print(clashs)
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