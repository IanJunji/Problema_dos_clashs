



caminho = r'aaa_teste_para_criacao_de_xml/02.04-Drenagem X Sinalização Vertical-desagrupado.txt'  # Corrigido a mistura de barras
print(caminho)

with open(caminho, 'r', encoding='utf') as file:
    linhas = file.readlines()
    
dados = [linha.strip() for linha in linhas if linha.strip()]

# ... existing code ...

clashes = []
current_clash = None
current_item = None

for linha in linhas:
    linha = linha.strip()
    if not linha:
        continue
    
    if linha.startswith("Name:"):
        # New clash found
        current_clash = {
            "name": linha.split(":", 1)[1].strip(),
            "details": {},
            "items": []
        }
        clashes.append(current_clash)
    elif linha.startswith("Item "):
        # New item found within current clash
        current_item = {}
        current_clash["items"].append(current_item)
    elif ":" in linha and current_clash is not None:
        # Handle key-value pairs
        key, value = linha.split(":", 1)
        key = key.strip()
        value = value.strip()
        
        if current_item is not None:
            current_item[key] = value
        else:
            current_clash["details"][key] = value

# Print structured results
print("Structured Clash Data:")
for clash in clashes:
    print(f"\nClash Name: {clash['name']}")
    print("Details:")
    for k, v in clash['details'].items():
        print(f"  {k}: {v}")
    
    print("\nItems:")
    for item in clash['items']:
        print("  Item:")
        for k, v in item.items():
            print(f"    {k}: {v}")