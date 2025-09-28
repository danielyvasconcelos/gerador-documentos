"""
Script para criar release compacto (menor que 25MB)
"""

import os
import subprocess
import shutil
import zipfile

def criar_release_compacto():
    """Cria release apenas com executável"""
    
    print("Criando release compacto...")
    
    # Limpa builds anteriores
    if os.path.exists("dist"):
        shutil.rmtree("dist")
    if os.path.exists("build"):
        shutil.rmtree("build")
    
    # Gera executável compacto
    print("Gerando executável compacto...")
    comando = [
        "python", "-m", "PyInstaller",
        "--onefile",
        "--windowed", 
        "--name=GeradorDocumentos",
        "--strip",  # Remove símbolos de debug
        "--exclude-module=matplotlib",  # Remove módulos desnecessários
        "--exclude-module=scipy",
        "--exclude-module=pandas",
        "src/interface_grafica.py"
    ]
    
    try:
        subprocess.run(comando, check=True)
        print("Executável criado!")
    except subprocess.CalledProcessError:
        print("Erro ao criar executável")
        return
    
    # Cria ZIP apenas com executável
    print("Criando ZIP compacto...")
    with zipfile.ZipFile("GeradorDocumentos_Compacto.zip", "w", zipfile.ZIP_DEFLATED) as zipf:
        zipf.write("dist/GeradorDocumentos.exe", "GeradorDocumentos.exe")
        
        # Adiciona apenas arquivos essenciais
        zipf.writestr("LEIA-ME.txt", """
GERADOR DE DOCUMENTOS - R PONTES CONSTRUTORA

COMO USAR:
1. Execute GeradorDocumentos.exe
2. Selecione sua planilha Excel
3. Selecione seu template Word
4. Escolha onde salvar os documentos
5. Clique em "GERAR DOCUMENTOS"

REQUISITOS:
- Windows 7 ou superior
- Nenhuma instalação adicional necessária
""")
    
    tamanho = os.path.getsize("GeradorDocumentos_Compacto.zip") / 1024 / 1024
    print(f"Release compacto criado: GeradorDocumentos_Compacto.zip")
    print(f"Tamanho: {tamanho:.1f} MB")
    
    # Limpa arquivos temporários
    shutil.rmtree("build", ignore_errors=True)
    shutil.rmtree("dist", ignore_errors=True)
    
    if tamanho < 25:
        print("✅ Arquivo menor que 25MB - pode subir no GitHub!")
    else:
        print("⚠️ Ainda maior que 25MB - use GitHub Releases")

if __name__ == "__main__":
    criar_release_compacto()