"""
Sistema de atualização automática
"""

import requests
import os
import subprocess
from tkinter import messagebox

class Atualizador:
    def __init__(self):
        self.versao_atual = "1.0.0"
        self.url_servidor = "https://empresa.com/api/versao"
        
    def verificar_atualizacao(self):
        """Verifica se há nova versão disponível"""
        try:
            response = requests.get(self.url_servidor)
            versao_servidor = response.json()["versao"]
            
            if versao_servidor > self.versao_atual:
                return True, versao_servidor
            return False, None
            
        except:
            return False, None
    
    def baixar_atualizacao(self, versao):
        """Baixa nova versão"""
        url_download = f"https://empresa.com/downloads/GeradorDocumentos_v{versao}.exe"
        
        # Baixa arquivo
        response = requests.get(url_download)
        with open(f"GeradorDocumentos_v{versao}.exe", "wb") as f:
            f.write(response.content)
        
        return f"GeradorDocumentos_v{versao}.exe"
    
    def aplicar_atualizacao(self, arquivo_novo):
        """Aplica atualização"""
        # Cria script de atualização
        script = f"""
        @echo off
        timeout /t 2
        move "{arquivo_novo}" "GeradorDocumentos.exe"
        start GeradorDocumentos.exe
        del update.bat
        """
        
        with open("update.bat", "w") as f:
            f.write(script)
        
        # Executa atualização e fecha programa atual
        subprocess.Popen("update.bat", shell=True)
        exit()

def verificar_atualizacao_startup():
    """Verifica atualização na inicialização"""
    atualizador = Atualizador()
    tem_atualizacao, nova_versao = atualizador.verificar_atualizacao()
    
    if tem_atualizacao:
        resposta = messagebox.askyesno(
            "Atualização Disponível",
            f"Nova versão {nova_versao} disponível!\nDeseja atualizar agora?"
        )
        
        if resposta:
            arquivo = atualizador.baixar_atualizacao(nova_versao)
            atualizador.aplicar_atualizacao(arquivo)