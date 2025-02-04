import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import datetime
import os
import openpyxl
from openpyxl import load_workbook, workbook
import openpyxl.workbook

class ClashAnalyzerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Clash Analyzer")
        self.root.geometry("600x400")
        
        # Variables to store file paths
        self.clash_file_path = tk.StringVar()
        self.matrix_file_path = tk.StringVar()
        self.output_dir = tk.StringVar()
        
        self.create_widgets()
    
    def create_widgets(self):
        # Clash File Selection
        tk.Label(self.root, text="Arquivo de Clash:").pack(pady=5)
        tk.Entry(self.root, textvariable=self.clash_file_path, width=50).pack(pady=5)
        tk.Button(self.root, text="Selecionar Arquivo", command=self.select_clash_file).pack(pady=5)
        
        # Matrix File Selection
        tk.Label(self.root, text="Arquivo da Matriz:").pack(pady=5)
        tk.Entry(self.root, textvariable=self.matrix_file_path, width=50).pack(pady=5)
        tk.Button(self.root, text="Selecionar Matriz", command=self.select_matrix_file).pack(pady=5)
        
        # Output Directory Selection
        tk.Label(self.root, text="Diretório de Saída:").pack(pady=5)
        tk.Entry(self.root, textvariable=self.output_dir, width=50).pack(pady=5)
        tk.Button(self.root, text="Selecionar Diretório", command=self.select_output_dir).pack(pady=5)
        
        # Process Button
        tk.Button(self.root, text="Processar", command=self.process_files).pack(pady=20)
    
    def select_clash_file(self):
        filename = filedialog.askopenfilename(filetypes=[("Text files", "*.txt")])
        if filename:
            self.clash_file_path.set(filename)
    
    def select_matrix_file(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if filename:
            self.matrix_file_path.set(filename)
    
    def select_output_dir(self):
        dirname = filedialog.askdirectory()
        if dirname:
            self.output_dir.set(dirname)
    
    def process_files(self):
        if not all([self.clash_file_path.get(), self.matrix_file_path.get(), self.output_dir.get()]):
            messagebox.showerror("Erro", "Por favor, selecione todos os arquivos necessários.")
            return
        
        try:
            # Process clash file
            clashs = self.process_clash_file(self.clash_file_path.get())
            
            # Process matrix and get approved clashes
            clashs_aprovados = self.process_matrix(clashs, self.matrix_file_path.get())
            
            # Create discipline TXTs
            self.criar_txts_por_disciplina(clashs_aprovados, self.output_dir.get())
            
            messagebox.showinfo("Sucesso", "Processamento concluído com sucesso!")
        
        except Exception as e:
            messagebox.showerror("Erro", f"Erro durante o processamento: {str(e)}")

    def process_clash_file(self, filepath):
        clashs = []
        current_clash = {}
        
        with open(filepath, 'r', encoding='utf-8') as arquivo:
            linhas = arquivo.readlines()
        
        for linha in linhas:
            linha = linha.strip()
            
            if linha.startswith('Name:'):
                if current_clash:  # Se já tiver um clash em progresso
                    clashs.append(current_clash)
                current_clash = {}  # Começa novo clash
                _, valor = linha.split(':', 1)
                current_clash['name'] = valor.strip()
            
            elif linha.startswith('Image Location:'):
                _, valor = linha.split(':', 1)
                _, nome = valor.split('-', 1)
                nome_objs, _ = nome.split('-', 1)
                obj1, obj2 = nome_objs.split('X', 1)
                _, pre_id = linha.split('\\', 1)
                id, _ = pre_id.split('.', 1)
                current_clash['objetos'] = nome_objs.strip()
                current_clash['image_loc'] = valor.strip()
                current_clash['obj_1'] = obj1.strip()
                current_clash['obj_2'] = obj2.strip()
                current_clash['id'] = id.strip()

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
                current_clash['coordinates'] = valor.replace('m', '').strip()

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

    def process_matrix(self, clashs, matrix_path):
        planilha_matriz = openpyxl.load_workbook(matrix_path)
        aba_matriz = planilha_matriz['Matriz']
        clashs_aprovados = []
        
        for clash in clashs:
            obj1 = clash['obj_1']
            obj2 = clash['obj_2']
            for col in aba_matriz.iter_cols(min_row=2, max_row=2, min_col=4):
                for cel in col:
                    if cel.value == obj1:
                        coluna = cel.column
    
            for col in aba_matriz.iter_cols(min_col=2,max_col=2,min_row=4):
                for cel in col:
                    if cel.value == obj2:
                        linha = cel.row

            # Coluna e linha definidas
            if not aba_matriz.cell(row=linha, column=coluna).value == 'O':
                clashs_aprovados.append(clash)
        
        return clashs_aprovados

    def criar_txts_por_disciplina(self, clashs_aprovados, diretorio_saida):
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
            
            # Informações a serem escritas no TXT
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

            # Lista com os objetos para iterar
            disciplinas = [clash.get('obj_1', ''), clash.get('obj_2', '')]
            
            for disciplina in disciplinas:
                if disciplina:  # Verifica se a disciplina não está vazia
                    # Define o caminho do arquivo TXT para a disciplina
                    caminho_txt = os.path.join(diretorio_saida, f"{disciplina}.txt")
                    
                    # Abre o arquivo no modo append e escreve o conteúdo
                    with open(caminho_txt, 'a', encoding='utf-8') as txt_file:
                        txt_file.write(conteudo)

def main():
    root = tk.Tk()
    app = ClashAnalyzerGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()