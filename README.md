<div align="center">

# 📦 Gestão de Estoque

### Aplicação desktop com Python, MySQL e Google Sheets

![Python](https://img.shields.io/badge/Python-3.10%2B-3776AB?logo=python&logoColor=white)
![Interface](https://img.shields.io/badge/Interface-Tkinter-2E5E4E)
![Banco de dados](https://img.shields.io/badge/Banco-MySQL-4479A1?logo=mysql&logoColor=white)
![Relatórios](https://img.shields.io/badge/Relat%C3%B3rios-Google%20Sheets-34A853?logo=googlesheets&logoColor=white)

**Cadastre produtos, acompanhe movimentações, visualize indicadores e envie relatórios para o Google Sheets.**

</div>

---

## ✨ Visão geral

Este projeto é uma cópia adaptada do sistema original, que permanece inalterado.
A aplicação usa **Tkinter** na interface, **MySQL** para armazenar os dados e
**Google Sheets** para compartilhar os relatórios.

> [!NOTE]
> A exportação é feita para uma planilha do Google Sheets. O programa não gera
> arquivos Excel `.xlsx`.

| Área | O que oferece |
|---|---|
| **Produtos** | Cadastro, consulta, edição e exclusão |
| **Movimentações** | Registro de entradas e saídas vinculado a produtos |
| **Análise visual** | Gráficos por categoria e evolução mensal |
| **Relatórios** | Abas de dados consolidados e resumo gerencial no Sheets |

## 📊 Relatórios e análises

A planilha contém duas abas:

1. **Dados Consolidados** — movimentações relacionadas aos produtos por meio
   de uma consulta `JOIN`.
2. **Resumo Gerencial** — indicadores de estoque por categoria e movimentações
   por tipo.

As abas são formatadas automaticamente com cabeçalho destacado, linhas
alternadas, larguras ajustadas, primeira linha congelada e formatos adequados
para datas e valores monetários. Antes de exportar, o sistema solicita
confirmação, pois os dados das duas abas serão atualizados.

A aplicação também exibe dois gráficos:

- **Movimentações por categoria:** quantidade de unidades movimentadas.
- **Saldo mensal:** entradas menos saídas agrupadas por mês.

## 🚀 Como configurar

### 1. Preparar o banco de dados

Instale o MySQL e execute [`schema.sql`](schema.sql) pelo MySQL Workbench ou
cliente MySQL. O script cria o banco `estoque_gestao` com as tabelas
relacionadas, **5 produtos** e **15 movimentações iniciais**.

### 2. Instalar as dependências

No PowerShell, dentro da pasta do projeto:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

### 3. Configurar o MySQL

Defina as variáveis de ambiente conforme a configuração do seu banco:

| Variável | Valor padrão |
|---|---|
| `DB_HOST` | `localhost` |
| `DB_USER` | `root` |
| `DB_PASSWORD` | vazio |
| `DB_NAME` | `estoque_gestao` |

### 4. Configurar o Google Sheets

1. Configure uma conta de serviço do Google e habilite as APIs **Google Sheets**
   e **Google Drive**.
2. Compartilhe a planilha com o e-mail da conta de serviço.
3. Defina `GOOGLE_CREDENTIALS_FILE` com o caminho do arquivo JSON da chave.
4. Opcionalmente, defina `GOOGLE_SHEETS_ID` para usar uma planilha existente.

> [!WARNING]
> Não publique nem inclua o arquivo JSON de credenciais no repositório. Mantenha
> esse arquivo em local seguro e configure o caminho por variável de ambiente.

Se `GOOGLE_SHEETS_ID` não for definido, o sistema procura uma planilha chamada
`relatorio_estoque` e tenta criá-la caso ainda não exista.

### 5. Iniciar a aplicação

```powershell
python exel.py
```

## 🗂️ Arquivos do projeto

| Arquivo | Descrição |
|---|---|
| [`exel.py`](exel.py) | Aplicação Tkinter, operações no banco, gráficos e exportação |
| [`schema.sql`](schema.sql) | Estrutura relacional e dados iniciais |
| [`requirements.txt`](requirements.txt) | Dependências Python |

---

<div align="center">

Feito com **Python**, **MySQL** e **Google Sheets**.

</div>
