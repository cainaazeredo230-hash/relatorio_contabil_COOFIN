# 📊 Automação FlexVision - SIAFE (ETL + Relatórios Financeiros)

Sistema completo de automação para extração, tratamento e consolidação de dados financeiros do sistema **FlexVision (SIAFE - RJ)**.

O projeto realiza download automatizado de relatórios, processa regras contábeis complexas e gera um relatório Excel final com múltiplas abas e análises estruturadas.

---

## 🚀 Objetivo

Eliminar o trabalho manual de:

- Download de relatórios no FlexVision
- Cruzamento de dados entre NE, NL, PD e OB
- Consolidação de valores por credor
- Geração de relatórios gerenciais em Excel

---

## ⚙️ Stack utilizada

- Python 3.10+
- Playwright (automação web)
- Pandas (ETL e transformação de dados)
- OpenPyXL (formatação avançada Excel)
- Python-dotenv (variáveis de ambiente)

---

## 🧠 Arquitetura do sistema

Fluxo principal:

- FlexVision (Web)
- ↓
- Playwright (Login + Download)
- ↓
- Arquivos Excel (RAW)
- ↓
- Pandas ETL (tratamento e cruzamento)
- ↓
- Regras financeiras e contábeis
- ↓
- Excel final formatado (OpenPyXL)


---

## 📂 Estrutura do projeto

- 📦 projeto
- ├── main.py                # Orquestrador principal
- ├── login_siafe.py        # Login automático no FlexVision
- ├── extrair_relatorio.py  # Download dos relatórios
- ├── extrair_dados.py      # ETL principal (transformação de dados)
- ├── cruzamento.py         # Funções auxiliares de cruzamento
- │
- ├── config/
- │   ├── De_para.xlsx      # Normalização de credores
- │   └── Layout.xlsx       # Estrutura base do relatório
- │
- ├── data/
- │   └── (relatórios baixados automaticamente)
- │
- ├── ARQUIVOS_RETORNO/ # Onde ficará o arquivo de relatório
- │
- └── .env                  # Credenciais do sistema


---

## 🔄 Fluxo de execução

### 1. Autenticação
- Login automático no FlexVision via Playwright

### 2. Download de relatórios
São baixados automaticamente:

- PD (Programação de Desembolso)
- NE (Nota de Empenho)
- NL (Nota de Liquidação)
- OB (Ordens Bancárias)
- Lançamentos Contábeis

---

### 3. Tratamento de dados (ETL)

O sistema realiza:

#### 📌 Normalização
- Padronização de credores via `De_para.xlsx`
- Conversão de datas

#### 📌 Regras financeiras
- PDs anuladas são removidas
- Separação de:
  - Processado e Pago
  - Parcialmente Pago
  - Erro no pagamento

#### 📌 Cálculo de valores
- Valor original
- Valor GD (abatimentos)
- Valor líquido

#### 📌 Datas inteligentes
- Últimos 7 dias
- Mês atual
- Histórico

---

### 4. Cruzamento de dados

Relaciona:

- PD ↔ OB ↔ NE ↔ NL
- Credores ↔ contratos
- Lançamentos contábeis ↔ execução orçamentária

---

### 5. Geração do Excel final

Arquivo gerado: 

DD.MM.YYYY_PROCESSOS_EMPENHADOS_LIQUIDADOS_PAGOS YYYY.xlsx

Contém:

### 📄 Abas:

- Execução_NE_NL_PD_OB (principal)
- NE_NL_PD_OB_13 Empresa
- NE_NL_PD_OB_Deodônio
- analítico_pd
- analítico_ne
- analítico_nl

---

## 🧩 Regras de negócio importantes

- PDs anuladas são ignoradas (`Status PD == 2 - Anulado`)
- Situações consideradas:
  - Processado e Pago
  - Parcialmente Pago
  - Enviado ao Banco
- GD (abatimentos) são aplicados automaticamente
- Credores são normalizados via `De_para.xlsx`
- Datas de retorno são inferidas quando ausentes

- SISTEMA AINDA NÃO CONSEGUE AS SEGUINTES INFORMAÇÕES:
    - "PRESTADORES DE SERVIÇOS QUE FALTAM EMITIR PD"
    - "RETENÇÕES QUE FALTAM EMITIR PD"
    - "EMPENHOS (NE)" e "LIQUIDAÇÕES (NL)" dos seguintes credores:
        - BARUCH (ANTIGA GLX DE ITAOCARA)
        - Diárias
        - Multas e Juros sobre retenções
        - Taxa da emissão da Guia do ISS

---

## Como instalar
rodar os seguintes comandos no seu terminal:
  - pip install -r requirements.txt
  - pip install pyinstaller
  - pyinstaller --onefile --add-data "config;config" --add-data "data;data" --add-data ".env;." main.py

---

## 🔐 Configuração (.env)

```env
USUARIO=seu_cpf
SENHA=sua_senha
LINK=https://siafe2-flexvision.fazenda.rj.gov.br/Flexvision/
