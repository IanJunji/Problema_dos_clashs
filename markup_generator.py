import tkinter as tk
from tkinter import filedialog, messagebox, ttk  # Import ttk for themed widgets
import os
import openpyxl
from openpyxl import load_workbook
import sys
from collections import defaultdict
import threading # Importando a biblioteca de threading

import openpyxl.workbook

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
            'J2': 'Dispositivos de Segurança',
            'I2': 'Pavimentação',
            'L2': 'OAEs',
            'K2': 'Iluminação',
            'L4': 'Contenções',
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

def is_clash_complete(clash):
    """
    Verifica se o dicionário do clash contém todas as informações obrigatórias.
    Caso algum campo obrigatório esteja faltando ou com valor vazio (e para os layers,
    se o valor for 'Layer_vazio'), a função retorna False.
    """
    required_fields = [
        'name', 'id', 'coord_x', 'coord_y', 'coord_z',
        'disciplina_1', 'disciplina_2', 'entity_1', 'entity_2',
        'layer_1', 'layer_2'
    ]
    for field in required_fields:
        if field not in clash or not clash[field]:
            return False
        if field in ['layer_1', 'layer_2'] and clash[field] == 'Layer_vazio':
            return False
    return True

def process_clash_file(filepath):
    """
    Processa o arquivo de clash e separa os registros válidos dos problemáticos.
    Os registros válidos são retornados em 'clashs' e os problemáticos em 'clashs_problematicos'.
    Também retorna a lista de disciplinas encontradas.
    """
    lista_disciplinas = []
    clashs = []               # Registros completos
    clashs_problematicos = [] # Registros com informações faltantes
    current_clash = {}
    
    with open(filepath, 'r', encoding='utf-8') as arquivo:
        linhas = arquivo.readlines()
    
    for i in range(len(linhas)):
        linha = linhas[i].strip()
        
        if linha.startswith('Name:'):
            if current_clash:  # Se já houver um clash em progresso, faz a verificação
                if is_clash_complete(current_clash):
                    clashs.append(current_clash)
                else:
                    clashs_problematicos.append(current_clash)
            current_clash = {}  # Inicia um novo clash
            _, valor = linha.split(':', 1)
            current_clash['name'] = valor.strip()
        
        elif linha.startswith('Image Location:'):
            _, valor = linha.split(':', 1)
            _, pre_id = linha.split('\\', 1)
            id_val, _ = pre_id.split('.', 1)
            current_clash['image_loc'] = valor.strip()
            current_clash['id'] = id_val.strip()

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
                if disciplina not in lista_disciplinas:
                    lista_disciplinas.append(disciplina)
                key = 'disciplina_1' if 'disciplina_1' not in current_clash else 'disciplina_2'
                current_clash[key] = disciplina

        elif linha.startswith('Entity Handle:'):
            if 'entity_1' not in current_clash:
                _, valor = linha.split(':', 1)
                current_clash['entity_1'] = valor.strip()
            else:
                _, valor = linha.split(':', 1)
                current_clash['entity_2'] = valor.strip()
        
        elif linha.startswith('Item 1'):
            if linhas[i+1].startswith('Layer:'):
                _, valor = linhas[i+1].split(':', 1)
                current_clash['layer_1'] = valor.strip()
            else:
                current_clash['layer_1'] = 'Layer_vazio'

        elif linha.startswith('Item 2'):
            if linhas[i+1].startswith('Layer:'):
                _, valor = linhas[i+1].split(':', 1)
                current_clash['layer_2'] = valor.strip()
            else:
                current_clash['layer_2'] = 'Layer_vazio'

    # Ao final, verifica o último clash processado
    if current_clash:
        if is_clash_complete(current_clash):
            clashs.append(current_clash)
        else:
            clashs_problematicos.append(current_clash)
    return clashs, lista_disciplinas, clashs_problematicos

def process_matrix(clashs, matrix_path):
    # Carrega a planilha e as abas
    workbook = openpyxl.load_workbook(matrix_path)
    aba_matriz = workbook['Matriz']

    # Inicialize a lista que armazenará os clashs aprovados
    clashs_semi_aprovados = []

    for clash in clashs:
        # Inicializa as variáveis para cada clash
        coluna = None  # Coluna correspondente à disciplina_1
        row = None     # Linha correspondente à disciplina_2
        
        # Procura o valor de 'disciplina_1' na primeira linha (colunas 4 em diante)
        for col in aba_matriz.iter_cols(min_row=2, max_row=2, min_col=4):
            for cel in col:
                if cel.value == clash['disciplina_1']:
                    coluna = cel.column
                    break  # Encerra o loop interno
            if coluna is not None:
                break  # Encerra o loop externo se encontrado
        
        # Procura o valor de 'disciplina_2' na segunda coluna (linhas a partir da 4)
        for col in aba_matriz.iter_cols(min_col=2, max_col=2, min_row=4):
            for cel in col:
                if cel.value == clash['disciplina_2']:
                    row = cel.row
                    break  # Encerra o loop interno
            if row is not None:
                break  # Encerra o loop externo se encontrado
                
        # Verifica se ambos 'coluna' e 'row' foram definidos e se a célula correspondente na matriz é 'O'
        if coluna is not None and row is not None:
            if aba_matriz.cell(row=row, column=coluna).value == 'O':
                clashs_semi_aprovados.append(clash)
    return clashs_semi_aprovados

def separacao_de_excecao(clashs, matriz_path):
    planilha = openpyxl.load_workbook(matriz_path)
    aba = planilha['exceções']
    clashs_aprovados = []
    for clash in clashs:
        # Obtemos os layers do clash e garantimos que sejam strings sem espaços a mais
        layer1 = str(clash.get('layer_1', '')).strip()
        layer2 = str(clash.get('layer_2', '')).strip()
        is_excecao = False  # Flag que identifica se o clash está na lista de exceções
        
        # Percorre todas as linhas da aba de exceções.
        # Use range(1, aba.max_row + 1) para iterar corretamente sobre as linhas
        for i in range(1, aba.max_row + 1):
            cell_ex1 = aba.cell(row=i, column=2)
            cell_ex2 = aba.cell(row=i, column=4)
            # Obtemos o valor das células (convertendo para string e removendo espaços, se não forem None)
            layer_ex1 = str(cell_ex1.value).strip() if cell_ex1.value is not None else ''
            layer_ex2 = str(cell_ex2.value).strip() if cell_ex2.value is not None else ''
            
            if (layer1 == layer_ex1 and layer2 == layer_ex2) or (layer1 == layer_ex2 and layer2 == layer_ex1):
                is_excecao = True
                break  # Já encontrou a exceção, não é necessário continuar verificando
        
        if not is_excecao:
            clashs_aprovados.append(clash)
    
    return clashs_aprovados

def criar_txts_por_disciplina(clashs_aprovados, diretorio_saida):
    os.makedirs(diretorio_saida, exist_ok=True)
    
    for clash in clashs_aprovados:
        disp_1 = clash.get('disciplina_1', '')
        disp_2 = clash.get('disciplina_2', '')
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
                   f"Objetos: {disp_1} X {disp_2}\n" \
                   f"ID: {id_clash}\n" \
                   f"Entity1: {entity1}\n" \
                   f"Entity2: {entity2}\n" \
                   f"Layer1: {layer1}\n" \
                   f"Layer2: {layer2}\n" \
                   f"{'-'*40}\n"
        
        # Cria/atualiza os TXT para as disciplinas
        disciplinas = [clash.get('disciplina_1', ''), clash.get('disciplina_2', '')]
        
        for disciplina in disciplinas:
            if disciplina:
                caminho_txt = os.path.join(diretorio_saida, f"{disciplina}.txt")
                with open(caminho_txt, 'a', encoding='utf-8') as txt_file:
                    txt_file.write(conteudo)

def criar_txt_defeitos(clashs_problematicos, diretorio_saida):
    """
    Cria um arquivo 'defeitos.txt' no diretório de saída com os detalhes dos clashs
    que estão faltando informações obrigatórias.
    """
    os.makedirs(diretorio_saida, exist_ok=True)
    caminho_defeitos = os.path.join(diretorio_saida, "defeitos.txt")
    
    with open(caminho_defeitos, 'w', encoding='utf-8') as txt_file:
        for clash in clashs_problematicos:
            conteudo = f"Nome: {clash.get('name', 'N/A')}\n"
            conteudo += f"ID: {clash.get('id', 'N/A')}\n"
            conteudo += f"Coord X: {clash.get('coord_x', 'N/A')}\n"
            conteudo += f"Coord Y: {clash.get('coord_y', 'N/A')}\n"
            conteudo += f"Coord Z: {clash.get('coord_z', 'N/A')}\n"
            conteudo += f"Disciplina 1: {clash.get('disciplina_1', 'N/A')}\n"
            conteudo += f"Disciplina 2: {clash.get('disciplina_2', 'N/A')}\n"
            conteudo += f"Entity 1: {clash.get('entity_1', 'N/A')}\n"
            conteudo += f"Entity 2: {clash.get('entity_2', 'N/A')}\n"
            conteudo += f"Layer 1: {clash.get('layer_1', 'N/A')}\n"
            conteudo += f"Layer 2: {clash.get('layer_2', 'N/A')}\n"
            conteudo += "-"*40 + "\n"
            txt_file.write(conteudo)

def contagem_conflitos_totais(clashs):
    lista_conflitos = []
    contagem_conflitos_total = []
    for clash in clashs:
        if clash.get('layer_1') and clash.get('layer_2'):
            key = f"{clash['layer_1']}%{clash['layer_2']}"
            key_inv = f"{clash['layer_2']}%{clash['layer_1']}"
            if key not in lista_conflitos and key_inv not in lista_conflitos:
                lista_conflitos.append(key)
                contagem_conflitos_total.append(1)
            else:
                for i in range(len(lista_conflitos)):
                    if lista_conflitos[i] == key:
                        contagem_conflitos_total[i] += 1
    return lista_conflitos, contagem_conflitos_total

def separar_layers(clashs):
    disciplinas_layers = defaultdict(list)

    for clash in clashs:
        # Extrai todas as combinações layer-disciplina do clash
        layers_disciplinas = []
        
        # Verifica se tem disciplina_1 e layer_1
        if 'disciplina_1' in clash and 'layer_1' in clash:
            layers_disciplinas.append( (clash['layer_1'], clash['disciplina_1']) )
        
        # Verifica se tem disciplina_2 e layer_2
        if 'disciplina_2' in clash and 'layer_2' in clash:
            layers_disciplinas.append( (clash['layer_2'], clash['disciplina_2']) )

        # Adiciona ao dicionário garantindo a relação correta
        for layer, disciplina in layers_disciplinas:
            if layer not in disciplinas_layers[disciplina]:
                disciplinas_layers[disciplina].append(layer)

    return dict(disciplinas_layers)

def excel_conflitos_por_disciplina(conflitos_por_disciplina, dicionario_layer_disciplina, saida):
    workbook = openpyxl.Workbook()
    aba = workbook.active

    row_num = 1

    for disciplinas_chave, conflitos in conflitos_por_disciplina.items():
        disciplina1, disciplina2 = disciplinas_chave.split(' x ')
        total_conflitos = conflitos['total']  # Já temos o total calculado

        aba[f'A{row_num}'] = disciplina1
        aba[f'B{row_num}'] = disciplina2
        aba[f'C{row_num}'] = total_conflitos
        row_num += 1

        if conflitos['conflitos']:
            aba[f'A{row_num}'] = 'Layer 1'
            aba[f'B{row_num}'] = 'Layer 2'
            aba[f'C{row_num}'] = 'Contagem'
            row_num += 1

        for layer_conflito, contagem in conflitos['conflitos'].items():
            layer1, layer2 = layer_conflito.split('%')

            # Verifica a qual disciplina cada layer pertence
            if layer1 in dicionario_layer_disciplina.get(disciplina1, []):
                # layer1 pertence à disciplina1, layer2 pertence à disciplina2
                aba[f'A{row_num}'] = layer1
                aba[f'B{row_num}'] = layer2
            else:
                # layer2 pertence à disciplina1, layer1 pertence à disciplina2
                aba[f'A{row_num}'] = layer2
                aba[f'B{row_num}'] = layer1

            aba[f'C{row_num}'] = contagem
            row_num += 1

        row_num += 1

    workbook.save(f'{saida}/lista_conflitos_disciplinas.xlsx')

def relacionar_conflitos_disciplinas(lista_conflitos, contagem_conflitos_total, lista_disciplinas, dicionario_layer_disciplina):
    conflitos_por_disciplina = {}

    for i in range(len(lista_disciplinas)):
        for j in range(i + 1, len(lista_disciplinas)):  # Evita duplicatas (A x B e B x A)
            disciplina1 = lista_disciplinas[i]
            disciplina2 = lista_disciplinas[j]
            chave_disciplinas = f"{disciplina1} x {disciplina2}"

            conflitos_layer = {}
            total_conflitos = 0

            for k in range(len(lista_conflitos)):
                layer1, layer2 = lista_conflitos[k].split('%')
                contagem = contagem_conflitos_total[k]

                # Verifica se o par de layers pertence às disciplinas
                if (
                    (layer1 in dicionario_layer_disciplina.get(disciplina1, []) and
                     layer2 in dicionario_layer_disciplina.get(disciplina2, []))
                    or
                    (layer2 in dicionario_layer_disciplina.get(disciplina1, []) and
                     layer1 in dicionario_layer_disciplina.get(disciplina2, []))
                ):
                    conflitos_layer[lista_conflitos[k]] = contagem
                    total_conflitos += contagem

            conflitos_por_disciplina[chave_disciplinas] = {
                'conflitos': conflitos_layer,
                'total': total_conflitos
            }

    return conflitos_por_disciplina

def process_files():
    # Desabilita o botão de processamento para evitar cliques múltiplos
    process_button.config(state=tk.DISABLED)
    progress_bar['value'] = 0  # Reset progress bar
    progress_label.config(text="Processando arquivo de clash...")
    try:
        # Agora a função retorna também os clashs problemáticos
        clashs, lista_disciplinas, clashs_problematicos = process_clash_file(clash_file_path.get())
        progress_bar['value'] = 25
        progress_label.config(text="Processando matriz...")
        
        dicionario_layers_por_disciplinas = separar_layers(clashs)

        # DEBUG: Mostra as disciplinas e seus layers
        print("=== RELAÇÃO DISCIPLINA/LAYERS ===")
        for disciplina, layers in dicionario_layers_por_disciplinas.items():
            print(f"{disciplina}:")
            for layer in layers:
                print(f" - {layer}")

        lista_conflitos, contagem_conflitos_total = contagem_conflitos_totais(clashs)
        progress_bar['value'] = 50
        progress_label.config(text="Relacionando conflitos por disciplina...")
        conflitos_por_disciplina = relacionar_conflitos_disciplinas(lista_conflitos, contagem_conflitos_total, lista_disciplinas, dicionario_layers_por_disciplinas)
        excel_conflitos_por_disciplina(conflitos_por_disciplina, dicionario_layers_por_disciplinas, output_dir.get())
        progress_bar['value'] = 75
        progress_label.config(text="Processando matriz de aprovação...")
        clash_semi_aprovados = process_matrix(clashs, matrix_file_path.get())
        clashs_aprovados = separacao_de_excecao(clash_semi_aprovados, matrix_file_path.get())
        criar_txts_por_disciplina(clashs_aprovados, output_dir.get())
        
        # Cria o arquivo com os clashs problemáticos, se houver
        criar_txt_defeitos(clashs_problematicos, output_dir.get())
        
        progress_bar['value'] = 100
        progress_label.config(text="Concluído!")
        messagebox.showinfo("Sucesso", "Processamento concluído com sucesso!")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro durante o processamento: {str(e)}")
        progress_label.config(text="Erro durante o processamento.")
    finally:
        process_button.config(state=tk.NORMAL)

def start_processing():
    # Cria e inicia uma thread para a função process_files
    thread = threading.Thread(target=process_files)
    thread.start()

def create_gui():
    global clash_file_path, matrix_file_path, output_dir, process_button, progress_bar, progress_label # Adicionado progress_bar e progress_label como globais
    root = tk.Tk()
    root.title("Clash Analyzer")
    root.geometry("600x450") # Altura um pouco maior para a barra de progresso
    root.columnconfigure(0, weight=1) # Permite que a coluna 0 se expanda
    root.rowconfigure(0, weight=1)    # Permite que a linha 0 se expanda
    
    # Adiciona ícone
    if getattr(sys, 'frozen', False):
        icon_path = os.path.join(sys._MEIPASS, 'icon.ico')
    else:
        icon_path = 'icon.ico'
    try:
        root.iconbitmap(icon_path)
    except:
        pass
    
    # Inicializa as variáveis globais como StringVar's
    clash_file_path = tk.StringVar()
    matrix_file_path = tk.StringVar()
    output_dir = tk.StringVar()

    # Frame principal para organizar os widgets
    main_frame = ttk.Frame(root, padding="20") # Usando ttk.Frame e adicionando padding diretamente
    main_frame.grid(column=0, row=0, sticky=(tk.W, tk.E, tk.N, tk.S)) # Usando grid e sticky para expansão

    # Labels e Entradas - Organizado com grid
    ttk.Label(main_frame, text="Arquivo de Clash:").grid(column=0, row=0, sticky=tk.W, padx=5, pady=5)
    clash_entry = ttk.Entry(main_frame, textvariable=clash_file_path, width=50)
    clash_entry.grid(column=1, row=0, sticky=(tk.W, tk.E), padx=5, pady=5)
    ttk.Button(main_frame, text="Selecionar Arquivo", command=select_clash_file).grid(column=2, row=0, sticky=tk.W, padx=5, pady=5)

    ttk.Label(main_frame, text="Arquivo da Matriz:").grid(column=0, row=1, sticky=tk.W, padx=5, pady=5)
    matrix_entry = ttk.Entry(main_frame, textvariable=matrix_file_path, width=50)
    matrix_entry.grid(column=1, row=1, sticky=(tk.W, tk.E), padx=5, pady=5)
    ttk.Button(main_frame, text="Selecionar Matriz", command=select_matrix_file).grid(column=2, row=1, sticky=tk.W, padx=5, pady=5)

    ttk.Label(main_frame, text="Diretório de Saída:").grid(column=0, row=2, sticky=tk.W, padx=5, pady=5)
    output_entry = ttk.Entry(main_frame, textvariable=output_dir, width=50)
    output_entry.grid(column=1, row=2, sticky=(tk.W, tk.E), padx=5, pady=5)
    ttk.Button(main_frame, text="Selecionar Diretório", command=select_output_dir).grid(column=2, row=2, sticky=tk.W, padx=5, pady=5)

    # Barra de Progresso
    progress_bar = ttk.Progressbar(main_frame, orient=tk.HORIZONTAL, length=300, mode='determinate')
    progress_bar.grid(column=0, row=3, columnspan=3, sticky=(tk.W, tk.E), padx=5, pady=10)
    progress_label = ttk.Label(main_frame, text="") # Label para exibir o status do progresso
    progress_label.grid(column=0, row=4, columnspan=3, sticky=tk.W, padx=5)


    # Botão para Processar os Arquivos
    process_button = ttk.Button(main_frame, text="Processar", command=start_processing) # Chama start_processing agora
    process_button.grid(column=1, row=5, pady=20) # Centralizado na coluna 1

    root.mainloop()


if __name__ == "__main__":
    create_gui()