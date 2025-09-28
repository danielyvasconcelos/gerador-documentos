"""
Script para criar release do projeto
"""

import os
import shutil
import subprocess
import zipfile
from datetime import datetime

def criar_release():
    """Cria pacote completo para release"""
    
    print("Criando release...")
    
    # Versão atual
    versao = "v1.0.0"
    data = datetime.now().strftime("%Y-%m-%d")
    
    # Limpa builds anteriores
    if os.path.exists("dist"):
        shutil.rmtree("dist")
    if os.path.exists("build"):
        shutil.rmtree("build")
    
    # Gera executável
    print("Gerando executável...")
    comando = [
        "python", "-m", "PyInstaller",
        "--onefile",
        "--windowed", 
        "--name=GeradorDocumentos",
        "--distpath=release",
        "src/interface_grafica.py"
    ]
    
    try:
        subprocess.run(comando, check=True)
        print("Executável criado!")
    except subprocess.CalledProcessError:
        print("Erro ao criar executável")
        return
    
    # Cria pasta de release
    release_folder = f"GeradorDocumentos_{versao}"
    if os.path.exists(release_folder):
        shutil.rmtree(release_folder)
    
    os.makedirs(release_folder)
    os.makedirs(f"{release_folder}/templates")
    os.makedirs(f"{release_folder}/exemplos")
    
    # Copia arquivos necessários
    print("Copiando arquivos...")
    shutil.copy("release/GeradorDocumentos.exe", release_folder)
    shutil.copy("templates/modelo_proposta.docx", f"{release_folder}/templates/")
    shutil.copy("README.md", release_folder)
    shutil.copy("requirements.txt", release_folder)
    
    # Cria documentos de exemplo
    shutil.copytree("data/documentos_gerados", f"{release_folder}/exemplos", dirs_exist_ok=True)
    
    # Cria arquivo de instruções
    with open(f"{release_folder}/LEIA-ME.txt", "w", encoding="utf-8") as f:
        f.write(f"""
GERADOR DE DOCUMENTOS - R PONTES CONSTRUTORA
Versão: {versao}
Data: {data}

🚀 COMO USAR:
1. Execute GeradorDocumentos.exe
2. Selecione sua planilha Excel
3. Selecione seu template Word
4. Escolha onde salvar os documentos
5. Clique em "GERAR DOCUMENTOS"

📁 ARQUIVOS INCLUSOS:
- GeradorDocumentos.exe (aplicativo principal)
- templates/modelo_proposta.docx (template de exemplo)
- exemplos/ (documentos de exemplo gerados)
- README.md (documentação completa)

💡 REQUISITOS:
- Windows 7 ou superior
- Nenhuma instalação adicional necessária

🆘 SUPORTE:
- Consulte README.md para documentação completa
- Exemplos de uso na pasta exemplos/

Desenvolvido por: Daniely Vasconcelos
""")
    
    # Cria arquivo ZIP para distribuição
    print("Criando arquivo ZIP...")
    with zipfile.ZipFile(f"{release_folder}.zip", "w", zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, files in os.walk(release_folder):
            for file in files:
                file_path = os.path.join(root, file)
                arc_path = os.path.relpath(file_path, ".")
                zipf.write(file_path, arc_path)
    
    print(f"Release criado: {release_folder}.zip")
    print(f"Tamanho: {os.path.getsize(f'{release_folder}.zip') / 1024 / 1024:.1f} MB")
    
    # Limpa arquivos temporários
    shutil.rmtree("build", ignore_errors=True)
    shutil.rmtree("release", ignore_errors=True)
    
    print("\nRelease pronto para upload no GitHub!")
    print(f"Arquivo: {release_folder}.zip")

if __name__ == "__main__":
    criar_release()