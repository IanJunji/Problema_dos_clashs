def ler_arquivo_txt(nome_arquivo):
    """
    Lê um arquivo de texto e retorna seu conteúdo
    
    Args:
        nome_arquivo (str): Caminho do arquivo a ser lido
        
    Returns:
        str: Conteúdo do arquivo
    """
    try:
        with open(nome_arquivo, 'r', encoding='utf-8') as arquivo:
            conteudo = arquivo.read()
        return conteudo
    except FileNotFoundError:
        print(f"Erro: O arquivo '{nome_arquivo}' não foi encontrado.")
        return None
    except Exception as e:
        print(f"Erro ao ler o arquivo: {str(e)}")
        return None
    
conteudo_txt = ler_arquivo_txt('aaa_teste_para_criacao_de_xml/02.04-Drenagem X Sinalização Vertical-desagrupado.txt')
linhas = conteudo_txt.splitlines() if conteudo_txt else []

# Adicionar prints de debug
print(f"Conteúdo do arquivo existe? {conteudo_txt is not None}")
if conteudo_txt:
    print(f"Número de linhas no arquivo: {len(linhas)}")
    print("Primeiras 5 linhas do arquivo:")
    for i, linha in enumerate(linhas[:5]):
        print(f"Linha {i+1}: {linha}")

clashs = []
clash = {}
in_clash = False
current_item = None

for linha in linhas:
    linha = linha.strip()
    print(f"Processando linha: '{linha}'")  # Debug print
    
    if linha.startswith('Name:'):
        print("-> Encontrou novo clash")  # Debug print
        in_clash = True
        if clash:  # Se já existe um clash anterior, salva ele
            clashs.append(clash.copy())
            print(f"-> Salvou clash anterior: {clash}")  # Debug print
        clash = {'name': linha[5:].strip()}
    elif in_clash and linha:
        if linha == 'Item 1':
            current_item = 'item1'
        elif linha == 'Item 2':
            current_item = 'item2'
        elif linha.startswith('Distance:'):
            clash['distance'] = linha[9:].strip()
        elif linha.startswith('HardStatus:'):
            clash['hardstatus'] = linha[11:].strip()
        elif linha.startswith('Clash Point:'):
            clash['clash_point'] = linha[12:].strip()
        elif linha.startswith('Date Created:'):
            clash['date_created'] = linha[13:].strip()
        elif linha.startswith('Date Approved:'):
            clash['date_approved'] = linha[14:].strip()
        elif linha.startswith('Approved By:'):
            clash['approved_by'] = linha[12:].strip()
        elif linha.startswith('Image Location:'):
            clash['image_location'] = linha[15:].strip()
        elif linha.startswith('Entity Handle:'):
            if current_item == 'item1':
                clash['entity_handle_item1'] = linha[14:].strip()
            elif current_item == 'item2':
                clash['entity_handle_item2'] = linha[14:].strip()
        elif linha.startswith('Layer:'):
            if current_item == 'item1':
                clash['layer_item1'] = linha[6:].strip()
            elif current_item == 'item2':
                clash['layer_item2'] = linha[6:].strip()
    elif linha == '------------------':
        print("-> Encontrou separador")  # Debug print
        if clash and in_clash:  # Se tem um clash em processamento
            clashs.append(clash.copy())
            print(f"-> Salvou clash atual: {clash}")  # Debug print
            clash = {}
            in_clash = False
            current_item = None

# Adiciona o último clash se existir
if clash and in_clash:
    clashs.append(clash.copy())
    print(f"-> Salvou último clash: {clash}")  # Debug print

print("\n=== RESULTADO FINAL ===")
print(f"Número total de clashs encontrados: {len(clashs)}")
for i, c in enumerate(clashs, 1):
    print(f"\nClash {i}:")
    for k, v in c.items():
        print(f"  {k}: {v}")

# Debug: mostra todas as chaves encontradas
print("\nTodas as chaves encontradas nos clashs:")
todas_chaves = set()
for c in clashs:
    todas_chaves.update(c.keys())
print(todas_chaves)