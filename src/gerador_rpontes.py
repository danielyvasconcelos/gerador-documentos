"""
Gerador de Documentos - R Pontes Construtora
Automatiza a criação de propostas comerciais a partir de dados do Excel
"""

import openpyxl
from docx import Document
import os
from datetime import datetime
from validador import ValidadorDados

class GeradorRPontes:
    """Classe principal que gera documentos Word a partir de dados do Excel"""
    
    def __init__(self, caminho_planilha, caminho_template):
        self.caminho_planilha = caminho_planilha
        self.caminho_template = caminho_template
        self.pasta_saida = "../data/documentos_gerados"
        
    def ler_dados_completos(self):
        """Lê e combina dados das 4 abas da planilha RPONTES"""
        wb = openpyxl.load_workbook(self.caminho_planilha)
        
        # Carrega cada aba da planilha
        aba_clientes = wb['Dados dos Clientes']
        aba_imoveis = wb['Dados do ImóvelServiço Contrata']
        aba_financeiro = wb['Dados Financeiros e Valores']
        aba_status = wb['StatusErro']
        
        dados_completos = []
        max_row = aba_clientes.max_row
        
        # Percorre cada linha de dados (pula o cabeçalho)
        for i in range(2, max_row + 1):
            cliente = {}
            
            # Dados pessoais
            cliente['ID_PROPOSTA'] = aba_clientes.cell(i, 1).value
            cliente['NOME_CLIENTE'] = aba_clientes.cell(i, 2).value
            cliente['DOCUMENTO_CLIENTE'] = aba_clientes.cell(i, 3).value
            cliente['EMAIL_CLIENTE'] = aba_clientes.cell(i, 4).value
            cliente['DATA_PROPOSTA'] = aba_clientes.cell(i, 5).value
            
            # Dados do imóvel
            cliente['EMPREENDIMENTO'] = aba_imoveis.cell(i, 1).value
            cliente['TIPO_UNIDADE'] = aba_imoveis.cell(i, 2).value
            cliente['METRAGEM'] = aba_imoveis.cell(i, 3).value
            cliente['DESCRICAO_SERVICOS'] = aba_imoveis.cell(i, 4).value
            
            # Dados financeiros - formata valores em reais
            valor_total = aba_financeiro.cell(i, 1).value
            valor_entrada = aba_financeiro.cell(i, 2).value
            cliente['VALOR_TOTAL'] = f"R$ {valor_total:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
            cliente['VALOR_ENTRADA'] = f"R$ {valor_entrada:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
            cliente['NUMERO_PARCELAS'] = int(aba_financeiro.cell(i, 3).value)
            cliente['DATA_PRIMEIRA_PARCELA'] = aba_financeiro.cell(i, 4).value
            
            # Status de controle
            cliente['STATUS_GERACAO'] = aba_status.cell(i, 1).value
            
            dados_completos.append(cliente)
            
        wb.close()
        return dados_completos
    
    def gerar_documento(self, dados_cliente):
        """Gera documento Word personalizado substituindo os placeholders"""
        doc = Document(self.caminho_template)
        
        # Substitui placeholders no formato {CAMPO} pelos valores reais
        for paragraph in doc.paragraphs:
            texto_original = paragraph.text
            for key, value in dados_cliente.items():
                if value:
                    placeholder = f"{{{key}}}"
                    if placeholder in texto_original:
                        texto_original = texto_original.replace(placeholder, str(value))
            paragraph.text = texto_original
        
        # Cria nome do arquivo: Proposta_001_2025_Maria_Silva.docx
        nome_cliente = str(dados_cliente['NOME_CLIENTE']).replace(' ', '_')
        id_proposta = str(dados_cliente['ID_PROPOSTA']).replace('/', '_')
        nome_arquivo = f"Proposta_{id_proposta}_{nome_cliente}.docx"
        caminho_saida = os.path.join(self.pasta_saida, nome_arquivo)
        
        doc.save(caminho_saida)
        return caminho_saida
    
    def processar_clientes(self, limite=None):
        """Método principal que coordena todo o processo de geração"""
        # Lê dados da planilha
        dados_clientes = self.ler_dados_completos()
        
        if limite:
            dados_clientes = dados_clientes[:limite]
        
        # Valida os dados antes de gerar documentos
        validador = ValidadorDados()
        print("Validando dados...")
        
        clientes_validos = []
        for i, cliente in enumerate(dados_clientes, 1):
            if validador.validar_cliente(cliente, i + 1):  # +1 por causa do cabeçalho
                clientes_validos.append(cliente)
        
        print(validador.obter_relatorio_erros())
        
        if not clientes_validos:
            print("Nenhum cliente valido encontrado!")
            return []
        
        # Gera documentos para clientes válidos
        documentos_gerados = []
        print(f"Gerando documentos para {len(clientes_validos)} clientes validos...")
        
        for i, cliente in enumerate(clientes_validos, 1):
            try:
                caminho_doc = self.gerar_documento(cliente)
                documentos_gerados.append(caminho_doc)
                print(f"{i:2d}. {cliente['NOME_CLIENTE']} - OK")
            except Exception as e:
                print(f"{i:2d}. {cliente['NOME_CLIENTE']} - ERRO: {e}")
        
        return documentos_gerados

# Executa apenas quando o arquivo é rodado diretamente
if __name__ == "__main__":
    planilha = "../data/RPONTES.xlsx"
    template = "../templates/modelo_proposta.docx"
    
    gerador = GeradorRPontes(planilha, template)
    
    print("Iniciando geração de documentos...")
    documentos = gerador.processar_clientes()
    print(f"\nConcluído! {len(documentos)} documentos gerados.")