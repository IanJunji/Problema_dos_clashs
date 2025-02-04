import win32com.client
import pythoncom
import sys

def criar_circulo(coordenadas):
    try:
        # Versão específica para AutoCAD 2012 (Release 19)
        acad = win32com.client.Dispatch("AutoCAD.Application.19")
        doc = acad.ActiveDocument
        model = doc.ModelSpace
        
        x, y, z = map(float, coordenadas.split(','))
        
        # Cria um círculo vermelho
        centro = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (x, y, z))
        raio = 0.5  # 0.5 metros
        circle = model.AddCircle(centro, raio)
        
        # Muda a cor para vermelho
        circle.TrueColor = win32com.client.Dispatch("AutoCAD.AcCmColor")
        circle.TrueColor.SetRGB(255, 0, 0)
        
        doc.Regen(True)  # Atualiza a tela
        acad.Visible = True
        
    except Exception as e:
        print(f"Erro: {e}")
        input("Pressione Enter para sair...")

if __name__ == "__main__":
    if len(sys.argv) > 1:
        criar_circulo(sys.argv[1]) 