"""
Script para gerar executável do Gerador de Documentos
"""

import os
import subprocess
import shutil

def criar_executavel():
    """Cria executável usando PyInstaller"""
    
    print("🔨 Criando executável...")
    
    # Comando PyInstaller
    comando = [
        "pyinstaller",
        "--onefile",                    # Arquivo único
        "--windowed",                   # Sem console
        "--name=GeradorDocumentos",     # Nome do executável
        "--icon=icon.ico",              # Ícone (opcional)
        "--add-data=templates;templates", # Inclui templates
        "src/interface_grafica.py"      # Arquivo principal
    ]
    
    try:
        # Executa PyInstaller
        subprocess.run(comando, check=True)
        
        print("✅ Executável criado com sucesso!")
        print("📁 Local: dist/GeradorDocumentos.exe")
        
        # Cria pasta de distribuição
        criar_pacote_distribuicao()
        
    except subprocess.CalledProcessError as e:
        print(f"❌ Erro ao criar executável: {e}")
    except FileNotFoundError:
        print("❌ PyInstaller não encontrado. Instale com: pip install pyinstaller")

def criar_pacote_distribuicao():
    """Cria pacote completo para distribuição"""
    
    print("\n📦 Criando pacote de distribuição...")
    
    # Cria pasta de distribuição
    dist_folder = "GeradorDocumentos_v1.0"
    if os.path.exists(dist_folder):
        shutil.rmtree(dist_folder)
    
    os.makedirs(dist_folder)
    os.makedirs(f"{dist_folder}/templates")
    os.makedirs(f"{dist_folder}/data")
    
    # Copia arquivos necessários
    shutil.copy("dist/GeradorDocumentos.exe", dist_folder)
    shutil.copy("templates/modelo_proposta.docx", f"{dist_folder}/templates/")
    shutil.copy("README.md", dist_folder)
    
    # Cria arquivo de instruções
    with open(f"{dist_folder}/COMO_USAR.txt", "w", encoding="utf-8") as f:
        f.write("""
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

SUPORTE:
- Consulte o README.md para mais detalhes
""")
    
    print(f"✅ Pacote criado: {dist_folder}/")
    print("📋 Conteúdo:")
    print("   - GeradorDocumentos.exe")
    print("   - templates/modelo_proposta.docx")
    print("   - README.md")
    print("   - COMO_USAR.txt")

if __name__ == "__main__":
    criar_executavel()