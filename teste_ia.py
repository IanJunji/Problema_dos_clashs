import os
import xml.etree.ElementTree as ET
import uuid
import datetime
import zipfile

def parse_txt_file(file_path):
    """
    Lê um arquivo TXT e extrai informações importantes.
    Para este exemplo, assume-se que o arquivo tem linhas do tipo "chave: valor" 
    (ex.: "title: Conflito de Elementos").
    Retorna um dicionário com os dados.
    """
    data = {}
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            for line in file:
                if ':' in line:
                    key, value = line.split(':', 1)
                    data[key.strip()] = value.strip()
    except FileNotFoundError:
        print(f"Arquivo {file_path} não encontrado.")
    return data

def generate_markup_xml(data, output_path):
    """
    Gera um arquivo XML (markup.bcf) preenchido com os dados extraídos.
    Garante os campos obrigatórios: GUID, Title, Description e Date.
    Outros campos (como Location) também podem ser incluídos.
    """
    markup = ET.Element('Markup')

    # Campo obrigatório: GUID
    guid = ET.SubElement(markup, 'GUID')
    guid.text = str(uuid.uuid4())

    # Outros campos: título, descrição e data
    title = ET.SubElement(markup, 'Title')
    title.text = data.get('title', 'No Title')
    
    description = ET.SubElement(markup, 'Description')
    description.text = data.get('description', 'No Description')
    
    date_elem = ET.SubElement(markup, 'Date')
    date_elem.text = datetime.datetime.now().isoformat()

    # Exemplo de campo extra: localização (caso exista)
    location = ET.SubElement(markup, 'Location')
    location.text = data.get('location', 'Unknown Location')

    # Se houver informações de viewpoint, elas poderiam ser referenciadas aqui.
    # Exemplo (comentado): viewpoint_ref = ET.SubElement(markup, 'Viewpoint')
    # viewpoint_ref.text = data.get('viewpoint', '')

    tree = ET.ElementTree(markup)
    tree.write(output_path, encoding='utf-8', xml_declaration=True)
    print(f"Arquivo de markup gerado em: {output_path}")

def generate_viewpoint_file(data, output_path):
    """
    Gera um arquivo de viewpoint (por exemplo, 1.bcfv) com dados de câmera, cortes, etc.
    Este exemplo cria um XML simples contendo apenas um campo 'Camera'.
    """
    viewpoint = ET.Element('Viewpoint')

    camera = ET.SubElement(viewpoint, 'Camera')
    camera.text = data.get('camera', 'Default Camera')

    tree = ET.ElementTree(viewpoint)
    tree.write(output_path, encoding='utf-8', xml_declaration=True)
    print(f"Arquivo de viewpoint gerado em: {output_path}")

def package_bcf(zip_output_path, markup_path, viewpoints_dir, snapshots_dir=None):
    """
    Organiza os arquivos e pastas (markup.bcf, pasta de Viewpoints e Snapshots) 
    e os compacta em um único arquivo .bcfzip.
    """
    with zipfile.ZipFile(zip_output_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        # Adiciona o arquivo XML do markup
        zipf.write(markup_path, arcname=os.path.basename(markup_path))
        
        # Adiciona os arquivos da pasta de Viewpoints (se existirem)
        if os.path.exists(viewpoints_dir):
            for foldername, subfolders, filenames in os.walk(viewpoints_dir):
                for filename in filenames:
                    file_path = os.path.join(foldername, filename)
                    # Garante manter a estrutura de pastas relativa
                    arcname = os.path.relpath(file_path, os.path.dirname(viewpoints_dir))
                    zipf.write(file_path, arcname=arcname)
        
        # Adiciona os arquivos da pasta de Snapshots (caso exista e seja necessária sua inclusão)
        if snapshots_dir and os.path.exists(snapshots_dir):
            for foldername, subfolders, filenames in os.walk(snapshots_dir):
                for filename in filenames:
                    file_path = os.path.join(foldername, filename)
                    arcname = os.path.relpath(file_path, os.path.dirname(snapshots_dir))
                    zipf.write(file_path, arcname=arcname)
    print(f"BCFZIP criado em: {zip_output_path}")

def parse_clash_records(file_path):
    """
    Lê um arquivo TXT contendo registros de clashes.
    
    Cada registro deve ter:
      - Uma seção "header": com campos (ex.: Name, Distance, Date Created, etc.).
      - Uma ou mais seções "item": iniciadas com "Item" e seguidas de campos específicos.
    
    Blocos de registros são assumidos estar separados por uma linha composta por hífens (ex.: "------------------").
    
    Retorna:
      Uma lista de dicionários. Cada dicionário possui:
         - "header": dicionário com os campos da parte principal.
         - "items": lista de dicionários, um para cada bloco iniciado com "Item".
    """
    records = []
    current_record = {"header": {}, "items": []}
    current_item = None

    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue  # pula linhas em branco

                # Se a linha for um separador de registro
                if line.startswith('-----'):
                    # Finaliza o item atual, se estiver em um bloco de item
                    if current_item is not None:
                        current_record["items"].append(current_item)
                        current_item = None
                    # Se o registro tiver sido preenchido, adiciona-o à lista
                    if current_record["header"] or current_record["items"]:
                        records.append(current_record)
                    # Reinicia para o próximo registro
                    current_record = {"header": {}, "items": []}
                    continue

                # Se a linha indicar o início de um item
                if line.startswith("Item"):
                    # Se já houver um item em processamento, armazena-o
                    if current_item is not None:
                        current_record["items"].append(current_item)
                    # Cria um novo dicionário para o item (aqui guarda a própria linha de identificação)
                    current_item = {"Item": line}
                    continue

                # Se a linha contiver ':' presume-se que seja uma linha de par chave-valor
                if ':' in line:
                    key, value = line.split(':', 1)
                    key = key.strip()
                    value = value.strip()
                    if current_item is not None:
                        current_item[key] = value
                    else:
                        current_record["header"][key] = value
                else:
                    # Caso a linha não contenha ':' e não seja identificador de item ou separador, pode ser tratada como dado extra
                    if current_item is not None:
                        current_item.setdefault("Extra", "")
                        current_item["Extra"] += line + " "
                    else:
                        current_record["header"].setdefault("Extra", "")
                        current_record["header"]["Extra"] += line + " "

        # Ao final do arquivo, adiciona o que estiver pendente
        if current_item is not None:
            current_record["items"].append(current_item)
        if current_record["header"] or current_record["items"]:
            records.append(current_record)
    except FileNotFoundError:
        print(f"Arquivo {file_path} não encontrado.")
    
    return records

def main():
    """
    Fluxo Principal:
      1. Leitura e extração dos dados do arquivo TXT.
      2. Geração do XML do markup (markup.bcf).
      3. (Opcional) Geração de arquivos de viewpoint (ex.: 1.bcfv).
      4. (Opcional) Inserção de snapshots (aqui apenas criamos a pasta para exemplificar).
      5. Empacotamento de tudo em um arquivo BCFZIP.
    """
    # Definição de caminhos e nomes de arquivo
    txt_file = r'D:\geoconversor\aaa_teste_para_criacao_de_xml\02.04-Drenagem X Sinalização Vertical-desagrupado.txt'                   # Arquivo TXT de entrada (deve estar no mesmo diretório ou informar o caminho completo)
    markup_file = 'markup.bcf'               # Arquivo XML de saída (markup)
    viewpoints_folder = 'Viewpoints'         # Pasta onde os arquivos de viewpoint serão gerados
    snapshots_folder = 'Snapshots'           # Pasta para snapshots (opcional)

    # Criação das pastas se não existirem
    os.makedirs(viewpoints_folder, exist_ok=True)
    os.makedirs(snapshots_folder, exist_ok=True)

    # 1. Leitura e extração do arquivo TXT
    clash_records = parse_clash_records(txt_file)
    
    if not clash_records:
        print("Nenhum registro de clash extraído do arquivo TXT. Verifique o conteúdo do arquivo.")
        return
    
    # Exemplo: processar o primeiro registro (header) do arquivo
    primeiro_clash = clash_records[0]
    print("Cabeçalho do primeiro clash:", primeiro_clash["header"])
    print("Itens do primeiro clash:", primeiro_clash["items"])
    
    # Aqui você pode iterar sobre os registros e para cada um gerar arquivos XML ou realizar outro processamento.

    # 2. Mapeamento dos dados está representado no dicionário "data"

    # 3. Geração do arquivo markup.bcf
    generate_markup_xml(primeiro_clash["header"], markup_file)

    # 4. Geração do arquivo de viewpoint (opcional)
    viewpoint_file_path = os.path.join(viewpoints_folder, '1.bcfv')
    generate_viewpoint_file(primeiro_clash["header"], viewpoint_file_path)

    # (Opcional) - Aqui você poderia copiar ou gerar snapshots na pasta 'Snapshots'
    # Exemplo: copiar uma imagem de referência, se necessário.

    # 5. Empacotamento em um arquivo BCFZIP
    bcfzip_file = 'output.bcfzip'
    package_bcf(bcfzip_file, markup_file, viewpoints_folder, snapshots_dir=snapshots_folder)

    # 6. Testes e Validação (pode ser implementado após a geração para automatizar a verificação)
    print("Processo de geração do BCFZIP concluído.")

if __name__ == '__main__':
    main()