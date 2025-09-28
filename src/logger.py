"""
Sistema de Logs para o Gerador de Documentos
Registra todas as operações com timestamp para auditoria
"""

import logging
import os
from datetime import datetime

class Logger:
    """Gerencia logs detalhados do sistema"""
    
    def __init__(self):
        self.pasta_logs = "../data/logs"
        self._criar_pasta_logs()
        self._configurar_logger()
        
    def _criar_pasta_logs(self):
        """Cria pasta de logs se não existir"""
        if not os.path.exists(self.pasta_logs):
            os.makedirs(self.pasta_logs)
    
    def _configurar_logger(self):
        """Configura o sistema de logging"""
        # Nome do arquivo com data atual
        data_hoje = datetime.now().strftime("%Y%m%d")
        arquivo_log = os.path.join(self.pasta_logs, f"gerador_{data_hoje}.log")
        
        # Configuração do logger
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(arquivo_log, encoding='utf-8'),
                logging.StreamHandler()  # Também mostra no console
            ]
        )
        
        self.logger = logging.getLogger(__name__)
    
    def info(self, mensagem):
        """Registra informação geral"""
        self.logger.info(mensagem)
    
    def sucesso(self, mensagem):
        """Registra operação bem-sucedida"""
        self.logger.info(f"SUCESSO: {mensagem}")
    
    def erro(self, mensagem):
        """Registra erro"""
        self.logger.error(f"ERRO: {mensagem}")
    
    def inicio_processo(self, processo):
        """Marca início de um processo"""
        self.logger.info(f"INICIO: {processo}")
    
    def fim_processo(self, processo, duracao=None):
        """Marca fim de um processo"""
        if duracao:
            self.logger.info(f"FIM: {processo} - Duração: {duracao:.2f}s")
        else:
            self.logger.info(f"FIM: {processo}")
    
    def estatisticas(self, total, sucessos, erros):
        """Registra estatísticas finais"""
        self.logger.info(f"ESTATISTICAS: Total={total}, Sucessos={sucessos}, Erros={erros}")