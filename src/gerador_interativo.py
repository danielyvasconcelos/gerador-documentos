"""
Gerador Interativo - Permite selecionar planilha e template
"""

import tkinter as tk
from tkinter import filedialog, messagebox
import os
from gerador_rpontes import GeradorRPontes

def selecionar_arquivos():
    """Interface simples para seleção de arquivos"""
    root = tk.Tk()
    root.withdraw()  # Esconde janela principal
    
    print("=== GERADOR DE DOCUMENTOS R PONTES ===")
    print("Selecione os arquivos necessários...\n")
    
    # Selecionar planilha Excel
    print("1. Selecione a planilha Excel...")
    planilha = filedialog.askopenfilename(
        title="Selecionar Planilha Excel",
        filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
    )
    
    if not planilha:
        print("❌ Nenhuma planilha selecionada. Cancelando...")
        return
    
    print(f"✅ Planilha: {os.path.basename(planilha)}")
    
    # Selecionar template Word
    print("\n2. Selecione o template Word...")
    template = filedialog.askopenfilename(
        title="Selecionar Template Word",
        filetypes=[("Word files", "*.docx"), ("All files", "*.*")]
    )
    
    if not template:
        print("❌ Nenhum template selecionado. Cancelando...")
        return
    
    print(f"✅ Template: {os.path.basename(template)}")
    
    # Selecionar pasta de saída
    print("\n3. Selecione onde salvar os documentos...")
    pasta_saida = filedialog.askdirectory(
        title="Selecionar Pasta para Salvar Documentos"
    )
    
    if not pasta_saida:
        print("❌ Nenhuma pasta selecionada. Usando pasta padrão...")
        pasta_saida = "../data/documentos_gerados"
    
    print(f"✅ Pasta de saída: {pasta_saida}")
    
    # Confirmar execução
    print(f"\n📋 RESUMO:")
    print(f"   Planilha: {os.path.basename(planilha)}")
    print(f"   Template: {os.path.basename(template)}")
    print(f"   Pasta de saída: {pasta_saida}")
    
    resposta = input("\n🚀 Iniciar geração? (s/n): ").lower().strip()
    
    if resposta in ['s', 'sim', 'y', 'yes']:
        try:
            print("\n" + "="*50)
            print("INICIANDO GERAÇÃO DE DOCUMENTOS...")
            print("="*50)
            
            # Criar gerador com arquivos selecionados
            gerador = GeradorRPontes(planilha, template, pasta_saida)
            documentos = gerador.processar_clientes()
            
            print(f"\n🎉 SUCESSO! {len(documentos)} documentos gerados!")
            print(f"📁 Local: {pasta_saida}")
            
            # Perguntar se quer abrir pasta
            abrir = input("\n📂 Abrir pasta dos documentos? (s/n): ").lower().strip()
            if abrir in ['s', 'sim', 'y', 'yes']:
                os.startfile(pasta_saida)
                
        except Exception as e:
            print(f"\n❌ ERRO: {e}")
            messagebox.showerror("Erro", f"Erro durante execução:\n{e}")
    else:
        print("Operação cancelada.")
    
    input("\nPressione Enter para sair...")

if __name__ == "__main__":
    selecionar_arquivos()