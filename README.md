# Automação da Geração de Propostas, Orçamentos e Contratos
## R Pontes Construtora

### Descrição do Projeto
Sistema de automação para geração de documentos personalizados (propostas, orçamentos e contratos) utilizando dados estruturados de planilhas Excel.

### Problema Resolvido
- Eliminação do processo manual de cópia de dados cliente por cliente
- Redução de erros humanos na criação de documentos
- Otimização do tempo da equipe de vendas

### Tecnologias Utilizadas
- **Python**: Linguagem principal
- **OpenPyXL**: Manipulação de planilhas Excel
- **python-docx**: Manipulação de documentos Word
- **OS**: Organização de arquivos

### Estrutura do Projeto
```
gerador-documentos/
├── data/
│   ├── clientes_pendentes.xlsx
│   └── documentos_gerados/
├── templates/
│   └── modelo_proposta.docx
├── src/
│   ├── gerador_documentos.py
│   └── utils.py
├── requirements.txt
└── README.md
```

### Impacto Esperado
- Redução de tempo: de horas para minutos
- Eliminação de erros manuais
- Padronização de documentos
- Maior produtividade da equipe