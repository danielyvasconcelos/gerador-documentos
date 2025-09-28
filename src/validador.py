"""
Validador de dados para o gerador de documentos
Verifica CPF, valores e campos obrigatórios
"""

import re
from datetime import datetime

class ValidadorDados:
    """Valida dados dos clientes antes de gerar os documentos"""
    
    def __init__(self):
        self.erros = []
    
    def validar_cpf(self, cpf):
        """Valida formato básico do CPF"""
        if not cpf:
            return False
        
        # Remove caracteres especiais
        cpf = re.sub(r'[^0-9]', '', str(cpf))
        
        # Verifica se tem 11 dígitos e não são todos iguais
        if len(cpf) != 11 or cpf == cpf[0] * 11:
            return False
            
        return True
    
    def validar_valor_monetario(self, valor):
        """Verifica se valor monetário é positivo"""
        try:
            if isinstance(valor, str):
                # Remove formatação: "R$ 1.500,50" -> 1500.50
                valor = valor.replace('R$', '').replace(' ', '')
                valor = valor.replace('.', '').replace(',', '.')
                valor = float(valor)
            return valor > 0
        except:
            return False
    
    def validar_cliente(self, cliente, linha):
        """Valida todos os dados de um cliente"""
        erros_cliente = []
        
        # Campos obrigatórios
        campos_obrigatorios = ['NOME_CLIENTE', 'DOCUMENTO_CLIENTE', 'EMPREENDIMENTO', 'VALOR_TOTAL']
        
        for campo in campos_obrigatorios:
            if not cliente.get(campo):
                erros_cliente.append(f"Campo {campo} esta vazio")
        
        # Valida CPF se preenchido
        if cliente.get('DOCUMENTO_CLIENTE'):
            if not self.validar_cpf(cliente['DOCUMENTO_CLIENTE']):
                erros_cliente.append("CPF invalido")
        
        # Valida valores monetários
        if cliente.get('VALOR_TOTAL'):
            if not self.validar_valor_monetario(cliente['VALOR_TOTAL']):
                erros_cliente.append("Valor total invalido")
        
        if cliente.get('VALOR_ENTRADA'):
            if not self.validar_valor_monetario(cliente['VALOR_ENTRADA']):
                erros_cliente.append("Valor entrada invalido")
        
        # Valida número de parcelas
        if cliente.get('NUMERO_PARCELAS'):
            try:
                parcelas = int(cliente['NUMERO_PARCELAS'])
                if parcelas <= 0:
                    erros_cliente.append("Numero de parcelas deve ser positivo")
            except:
                erros_cliente.append("Numero de parcelas invalido")
        
        # Acumula erros se houver
        if erros_cliente:
            self.erros.append({
                'linha': linha,
                'cliente': cliente.get('NOME_CLIENTE', 'N/A'),
                'erros': erros_cliente
            })
            return False
        
        return True
    
    def obter_relatorio_erros(self):
        """Gera relatório dos erros encontrados"""
        if not self.erros:
            return "Todos os dados estao validos!"
        
        relatorio = f"{len(self.erros)} cliente(s) com problemas:\n\n"
        
        for erro in self.erros:
            relatorio += f"Linha {erro['linha']} - {erro['cliente']}:\n"
            for e in erro['erros']:
                relatorio += f"  • {e}\n"
            relatorio += "\n"
        
        return relatorio