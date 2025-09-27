from gerador_rpontes import GeradorRPontes

def main():
    planilha = "../data/RPONTES.xlsx"
    template = "../templates/modelo_proposta.docx"
    
    gerador = GeradorRPontes(planilha, template)
    
    print("=== GERADOR DE DOCUMENTOS R PONTES ===")
    print("1. Processar TODOS os clientes (30)")
    print("2. Processar apenas os primeiros 5")
    print("3. Sair")
    
    opcao = input("\nEscolha uma opção: ")
    
    if opcao == "1":
        print("\nProcessando TODOS os 30 clientes...")
        documentos = gerador.processar_clientes()
        print(f"\n✅ {len(documentos)} documentos gerados!")
        
    elif opcao == "2":
        print("\nProcessando primeiros 5 clientes...")
        documentos = gerador.processar_clientes(limite=5)
        print(f"\n✅ {len(documentos)} documentos gerados!")
        
    elif opcao == "3":
        print("Saindo...")
        return
    else:
        print("Opção inválida!")

if __name__ == "__main__":
    main()