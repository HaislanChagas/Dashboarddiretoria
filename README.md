# 📊 Dashboard Diretoria

Dashboard interativo desenvolvido em Python com Streamlit para visualização de dados gerenciais para a diretoria.

## 🚀 Tecnologias Utilizadas

- **Python** — linguagem principal (81% do projeto)
- **Streamlit** — framework para criação do dashboard web
- **Pandas** — manipulação e análise de dados
- **Plotly** — gráficos interativos
- **gspread + google-auth** — integração com Google Sheets
- **openpyxl** — leitura e escrita de arquivos Excel (.xlsx)
- **CSS customizado** — estilização da interface (`style.css`)

## 📁 Estrutura do ProjetoDashboarddiretoria/
├── App.py              # Arquivo principal da aplicação
├── style.css           # Estilos customizados
├── requirements.txt    # Dependências do projeto
└── .devcontainer/      # Configuração para dev container

## ⚙️ Como Rodar Localmente

**1. Clone o repositório:**
```bash
git clone https://github.com/HaislanChagas/Dashboarddiretoria.git
cd Dashboarddiretoria
```

**2. Crie um ambiente virtual e instale as dependências:**
```bash
python -m venv venv
source venv/bin/activate  # Linux/Mac
venv\Scripts\activate     # Windows

pip install -r requirements.txt
```

**3. Execute a aplicação:**
```bash
streamlit run App.py
```

**4. Acesse no navegador:** `http://localhost:8501`

## 🔑 Configuração do Google Sheets

Para a integração com Google Sheets funcionar, é necessário configurar as credenciais da Google Cloud:

1. Crie um projeto no [Google Cloud Console](https://console.cloud.google.com/)
2. Ative a API do Google Sheets
3. Gere uma chave de serviço (Service Account)
4. Salve o arquivo `.json` de credenciais na raiz do projeto

## 📝 Licença

Este projeto é de uso interno.
