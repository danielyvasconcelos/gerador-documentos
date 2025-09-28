# 🚀 Gerador de Documentos - R Pontes Construtora

Automatiza a criação de propostas comerciais personalizadas a partir de dados estruturados em planilhas Excel.

## 📦 Download Rápido

**Quer usar sem instalar nada?**

➡️ **[Baixar Executável (GitHub Releases)](../../releases/latest)**

1. Clique no link acima
2. Baixe `GeradorDocumentos_v1.0.0.zip`
3. Extraia e execute `GeradorDocumentos.exe`
4. Pronto! ✅

## Funcionalidades

- ✅ **Leitura automática** de planilhas Excel com múltiplas abas
- ✅ **Validação de dados** (CPF, valores, campos obrigatórios)
- ✅ **Geração de documentos Word** personalizados
- ✅ **Seleção de arquivos** (planilha, template e pasta de saída)
- ✅ **Interface gráfica** intuitiva
- ✅ **Preservação de formatação** (logos, cores, estilos)
- ✅ **Logs detalhados** com timestamps e métricas
- ✅ **Relatórios de erro** para auditoria

## Tecnologias

- **Python 3.x**
- **OpenPyXL** - Manipulação de Excel
- **python-docx** - Geração de documentos Word
- **Tkinter** - Interface gráfica
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
│   ├── interface_grafica.py   # Interface gráfica
│   ├── gerador_interativo.py  # Seleção de arquivos
│   ├── validador.py           # Validação de dados
│   └── logger.py              # Sistema de logs
├── templates/
│   └── modelo_proposta.docx   # Template Word
├── requirements.txt
├── EXECUTAR.bat              # Execução rápida
├── INTERFACE.bat             # Interface gráfica
└── SELECIONAR_ARQUIVOS.bat   # Seleção interativa
```

## Como Executar

### 1. Instalar dependências

```bash
pip install -r requirements.txt
```

### 2. Executar o gerador

**Opção A - Interface Gráfica (Recomendado):**

```bash
INTERFACE.bat
```

**Opção B - Seleção de Arquivos:**

```bash
SELECIONAR_ARQUIVOS.bat
```

**Opção C - Comando direto:**

```bash
cd src
python gerador_rpontes.py
```

**Opção D - Arquivo executável:**

```bash
EXECUTAR.bat
```

### 3. Verificar resultados

- **Documentos:** `data/documentos_gerados/`
- **Logs:** `data/logs/gerador_YYYYMMDD.log`

## Exemplo de Uso

### **Usando Interface Gráfica:**

1. Execute `INTERFACE.bat`
2. Selecione sua planilha Excel
3. Selecione seu template Word personalizado
4. Escolha onde salvar os documentos
5. Clique em "GERAR DOCUMENTOS"

### **Usando Seleção de Arquivos:**

1. Execute `SELECIONAR_ARQUIVOS.bat`
2. Escolha planilha, template e pasta de saída
3. Confirme a geração

### **Usando arquivos padrão:**

1. Coloque os dados na planilha `data/RPONTES.xlsx`
2. Execute `EXECUTAR.bat`
3. Documentos são gerados em `data/documentos_gerados/`

## Logs e Monitoramento

O sistema gera logs detalhados com:

- Timestamp de cada operação
- Tempo de execução por documento
- Estatísticas de sucesso/erro
- Rastreamento completo do processo

## Características Avançadas

- **Preservação de Formatação:** Mantém logos, cores e estilos do template original
- **Flexibilidade:** Use qualquer planilha Excel e template Word
- **Interface Amigável:** Seleção visual de arquivos e pastas
- **Validação Inteligente:** Verifica CPF, valores e campos obrigatórios
- **Logs Profissionais:** Rastreamento completo com timestamps
- **Multiplataforma:** Funciona em qualquer sistema com Python

## 📥 Opções de Download

### **🎯 Para Usuários (Recomendado)**

**Baixar executável pronto:**
- ➡️ **[GitHub Releases](../../releases/latest)**
- Arquivo: `GeradorDocumentos_v1.0.0.zip` (33MB)
- **Não precisa instalar Python!**

### **👨‍💻 Para Desenvolvedores**

**Clonar e executar código:**
```bash
git clone https://github.com/SEU-USUARIO/gerador-documentos.git
cd gerador-documentos
pip install -r requirements.txt
python src/interface_grafica.py
```

**Gerar próprio executável:**
```bash
pip install pyinstaller
python build_release.py
```

## 🎯 Casos de Uso

### **🏢 Empresas que podem usar:**
- **Construtoras:** Propostas de imóveis
- **Escritórios:** Contratos de serviços  
- **Lojas:** Orçamentos personalizados
- **Escolas:** Certificados de alunos
- **Clínicas:** Relatórios médicos
- **Qualquer empresa** que gere documentos repetitivos

### **📊 Tipos de planilha aceitos:**
- Qualquer arquivo `.xlsx` ou `.xls`
- Uma ou múltiplas abas
- Qualquer estrutura de dados

### **📄 Tipos de template aceitos:**
- Qualquer arquivo `.docx`
- Preserva logos, cores e formatação
- Usa placeholders `{CAMPO}` para substituição

## 📈 Impacto

- **⏱️ Tempo:** De horas para minutos
- **🎯 Precisão:** Eliminação de erros manuais
- **🚀 Produtividade:** Automação completa do processo
- **🔧 Flexibilidade:** Use seus próprios templates e dados
- **📋 Auditoria:** Logs detalhados para controle

## 🤝 Contribuição

Contribuições são bem-vindas! 

1. Fork o projeto
2. Crie uma branch para sua feature
3. Commit suas mudanças
4. Push para a branch
5. Abra um Pull Request

---

**Desenvolvido por:** Daniely Vasconcelos  
**Versão:** 1.0.0  
**Status:** ✅ Estável e em produção
