# 📊 SQL Server Schema to Excel Exporter

![Python](https://img.shields.io/badge/python-3670A0?style=for-the-badge&logo=python&logoColor=ffdd54)
![Pandas](https://img.shields.io/badge/pandas-%23150458.svg?style=for-the-badge&logo=pandas&logoColor=white)
![SQL Server](https://img.shields.io/badge/SQLServer-CC2927?style=for-the-badge&logo=microsoft-sql-server&logoColor=white)

Uma automação em Python desenvolvida para extrair a estrutura completa de um banco de dados SQL Server e consolidar as informações em um único arquivo Excel (`.xlsx`). 

O script varre todas as tabelas do banco de dados, extraindo detalhes de colunas, tipos de dados, chaves primárias (PK), chaves estrangeiras (FK) e índices, separando cada tabela em uma aba (sheet) dedicada no arquivo Excel. Ideal para criar dicionários de dados ou auditorias de estrutura rapidamente.

## ✨ Funcionalidades

* **Mapeamento Automático:** Identifica todas as tabelas base (`BASE TABLE`) do banco de dados especificado.
* **Detalhamento de Colunas:** Extrai nome, tipo de dado (com precisão e tamanho), se permite nulo, e fórmulas/valores padrão.
* **Identificação de Chaves:** Mapeia automaticamente quais colunas são Primary Keys (PK) ou Foreign Keys (FK).
* **Mapeamento de Índices:** Lista os índices atrelados a cada tabela, incluindo o tipo e as colunas correspondentes.
* **Exportação Organizada:** Gera um arquivo `.xlsx` onde cada tabela do banco de dados ganha sua própria aba (sheet), dividindo visualmente a seção de colunas e a seção de índices.

## 🚀 Pré-requisitos

Antes de começar, você precisará ter o Python instalado na sua máquina e o driver ODBC do SQL Server.

1. **Python 3.x**
2. **ODBC Driver for SQL Server:** O script utiliza o `{ODBC Driver 17 for SQL Server}`. Certifique-se de tê-lo instalado no seu sistema operacional.
3. **Bibliotecas Python:** Instale as dependências executando o comando abaixo:

```bash
pip install pyodbc pandas openpyxl
