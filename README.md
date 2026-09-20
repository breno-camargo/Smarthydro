# 💧 Automação de Medição de Consumo de Água (EBO Schneider -> Excel)

Aplicação desenvolvida em Python para automatizar a extração, cálculo e geração de relatórios de consumo predial de água diretamente do banco de dados do **Schneider Electric EcoStruxure Building Operation (EBO)**.

---

## 📌 Contexto & Motivação
Anteriormente, os relatórios mensais dependiam de uma integração com um sistema legado de controle de acesso (W-Access). Após a migração para a plataforma Keyaccess, o software antigo permaneceu no ambiente sem licença e sem suporte apenas para a emissão dessa planilha mensal.

Esta ferramenta foi criada para eliminar esse ponto único de falha (*Single Point of Failure*), conectando-se diretamente à fonte dos dados e oferecendo uma interface simples para a equipe operacional.

---

## ✨ Funcionalidades
- **Conexão Direta ao Banco:** Consulta dados históricos e de medição direto no **Microsoft SQL Server** do EBO.
- **Cálculo de Consumo:** Processamento de consumo por hidrômetro/sala com validação de média dos meses anteriores (detecção de distorções/vazamentos).
- **Exportação Formatada:** Geração automática de arquivo `.xlsx` com formatação visual, resumo executivo e gráficos analíticos.
- **Interface Gráfica:** GUI amigável com seletor de datas/período para uso sem necessidade de familiaridade com terminal.

---

## 📸 Demonstração

### Interface Gráfica (Desktop)
![Interface do Sistema](assets/interface.png)

### Painel Executivo e Gráficos no Excel
![Painel Executivo e Gráficos](assets/relatorio_graficos.png)

### Planilha Detalhada de Consumo e Rateio
![Planilha de Consumo e Rateio](assets/relatorio_tabela.png)

---

## 🛠️ Tecnologias Utilizadas
- **Linguagem:** Python 3.x
- **Banco de Dados:** Microsoft SQL Server (`pyodbc`)
- **Manipulação de Dados:** `pandas`
- **Geração de Relatórios:** `openpyxl`
- **Interface Gráfica:** `Tkinter`

---

## 🚀 Como Executar

### Pré-requisitos
- Python instalado (versão 3.10 ou superior)
- Acesso à rede/instância do SQL Server do EBO
- Drivers ODBC configurados

```bash
# Clone o repositório
git clone https://github.com/seu-usuario/seu-repositorio.git

# Acesse o diretório
cd seu-repositorio

# Crie e ative o ambiente virtual
python -m venv venv
venv\Scripts\activate

# Instale as dependências
pip install -r requirements.txt

# Configure os parâmetros de conexão
copy config.example.json config.json

# Execute a aplicação
python main.py
```
