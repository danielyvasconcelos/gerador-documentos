# Gerador de Documentos - R Pontes Construtora

Automatiza a criação de propostas comerciais personalizadas a partir de dados estruturados em planilhas Excel.

## Funcionalidades

- ✅ **Leitura automática** de planilhas Excel com múltiplas abas
- ✅ **Validação de dados** (CPF, valores, campos obrigatórios)
- ✅ **Geração de documentos Word** personalizados
- ✅ **Logs detalhados** com timestamps e métricas
- ✅ **Relatórios de erro** para auditoria

## Tecnologias

- **Python 3.x**
- **OpenPyXL** - Manipulação de Excel
- **python-docx** - Geração de documentos Word
- **Logging** - Sistema de logs profissional

## Estrutura do Projeto

```
gerador-documentos/
├── data/
│   ├── RPONTES.xlsx           # Planilha com dados dos clientes
│   ├── documentos_gerados/    # Documentos Word gerados
│   └── logs/                  # Arquivos de log
├── src/
│   ├── gerador_rpontes.py     # Script principal
│   ├── validador.py           # Validação de dados
│   ├── logger.py              # Sistema de logs
│   └── analisar_planilha.py   # Análise da estrutura
├── templates/
│   └── modelo_proposta.docx   # Template Word
├── requirements.txt
└── EXECUTAR.bat              # Execução rápida
```

## Como Executar

### 1. Instalar dependências
```bash
pip install -r requirements.txt
```

### 2. Executar o gerador

**Opção A - Comando direto:**
```bash
cd src
python gerador_rpontes.py
```

**Opção B - Arquivo executável:**
```bash
EXECUTAR.bat
```

### 3. Verificar resultados
- **Documentos:** `data/documentos_gerados/`
- **Logs:** `data/logs/gerador_YYYYMMDD.log`

## Exemplo de Uso

1. Coloque os dados na planilha `data/RPONTES.xlsx`
2. Execute o script
3. Documentos personalizados são gerados automaticamente

## Logs e Monitoramento

O sistema gera logs detalhados com:
- Timestamp de cada operação
- Tempo de execução por documento
- Estatísticas de sucesso/erro
- Rastreamento completo do processo

## Impacto

- **Tempo:** De horas para minutos
- **Precisão:** Eliminação de erros manuais
- **Produtividade:** Automação completa do processo
- **Auditoria:** Logs detalhados para controle