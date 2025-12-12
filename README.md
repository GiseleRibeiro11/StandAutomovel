🚗 Stand Aguiar – Gestão de Viaturas

O Stand Aguiar – Gestão de Viaturas é uma aplicação Windows Forms desenvolvida em Visual Basic .NET, com ligação a SQL Server (LocalDB), destinada ao registo, gestão e impressão de informação sobre viaturas.

Projeto clássico, orientado a formulários, focado em operações CRUD, validação de dados e integração com base de dados. Funciona, é claro e cumpre o objetivo.

🎯 Objetivo do Projeto

Permitir a gestão simples e eficaz de viaturas, incluindo:

Inserção de dados

Alteração e eliminação de registos

Listagem e impressão

Exportação dos dados para base de dados SQL Server

🛠️ Funcionalidades

Adicionar viaturas (marca, modelo, matrícula e quilómetros)

Alterar registos existentes

Eliminar viaturas selecionadas

Sincronização entre listas (marca, modelo, matrícula e kms)

Validação de campos obrigatórios

Impressão da listagem de viaturas

Pré-visualização de impressão

Exportação dos dados para base de dados SQL Server

Ligação à base de dados validada no arranque da aplicação

🧰 Tecnologias Utilizadas

Visual Basic .NET

Windows Forms

SQL Server LocalDB

ADO.NET (SqlConnection, SqlCommand, SqlDataReader)

System.Drawing.Printing

Tecnologia clássica. Base sólida. Ainda muito usada em contextos empresariais legados — e quem domina isto, domina o resto.

🗄️ Base de Dados

Base de dados: LocalDB

Tabela: viaturas

Campos:

marca

modelo

matricula

kms

Os dados são exportados diretamente a partir da interface para a base de dados via comandos parametrizados.

⚙️ Como Executar o Projeto

Abrir o projeto no Visual Studio

Confirmar o caminho da base de dados no SqlConnection

Garantir que a tabela viaturas existe na base de dados

Executar a aplicação

Inserir dados e utilizar as funcionalidades disponíveis

🖨️ Impressão

A aplicação permite:

Pré-visualização da impressão

Impressão da listagem completa de viaturas

Layout estruturado com cabeçalho e colunas definidas

📌 Estado do Projeto

✔ Funcional
✔ Estrutura simples e objetiva
✔ Ideal para fins académicos e demonstração de lógica CRUD
✔ Base sólida para evolução futura

📄 Autora

Gisele Ribeiro
Programadora
Projeto desenvolvido em Visual Basic .NET com foco em lógica, validação e integração com base de dados.
