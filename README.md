# Sistema de Gestão de Manufatura em Python

![Python](https://img.shields.io/badge/python-3.x-blue.svg) ![Pandas](https://img.shields.io/badge/pandas-1.x-brightgreen.svg) ![ReportLab](https://img.shields.io/badge/reportlab-3.x-red.svg)

## 📖 Sobre o Projeto

Este é um sistema de linha de comando (CLI) para gestão de um pequeno processo de manufatura. Ele foi desenvolvido como uma solução prática para controlar o inventário de insumos (matéria-prima) e produtos acabados, automatizando a verificação de estoque antes de iniciar uma ordem de produção.

A principal característica do sistema é a sua capacidade de calcular a viabilidade de fabricar uma determinada quantidade de produtos, informando ao usuário se os insumos disponíveis são suficientes e alertando sobre o que está em falta. Os dados são lidos e salvos em uma planilha Excel, facilitando a visualização e edição externa.

## ✨ Funcionalidades

-   **📦 Módulo de Cadastro e Eliminação:**
    -   Cadastro de novos insumos com nome, descrição, quantidade e valor.
    -   Cadastro de produtos, com a definição da "receita" (quais insumos e quantidades são necessários).
    -   Eliminação segura de insumos e produtos (com validações para não excluir itens com estoque).

-   **🔄 Módulo de Estoque:**
    -   Atualização rápida da quantidade de qualquer insumo em estoque.

-   **🏭 Módulo de Fabricação:**
    -   Seleção de um produto e da quantidade a ser fabricada.
    -   Análise automática do estoque: o sistema compara os insumos necessários com os disponíveis.
    -   Feedback detalhado: informa se a produção é viável, listando os insumos suficientes e os que estão em falta.
    -   Cálculo do custo total de produção.

-   **📄 Módulo de Relatórios:**
    -   Geração de um relatório de insumos (`Relatorio_de_Insumos.pdf`), listando todos os itens em estoque, com seus detalhes.
    -   Geração de um relatório de produtos (`Relatorio_Produtos.pdf`), detalhando a composição de cada produto cadastrado.

## 🛠️ Tecnologias Utilizadas

-   **Python:** Linguagem principal do projeto.
-   **Pandas:** Utilizado para a manipulação de dados e interação com a planilha Excel (`BD/PLAN_ITENS.xlsx`), que funciona como banco de dados.
-   **ReportLab:** Biblioteca usada para gerar os relatórios em formato PDF de forma programática.

## 🚀 Como Executar

1.  **Clone o repositório:**
    ```bash
    git clone [https://URL-DO-SEU-REPOSITORIO.git](https://URL-DO-SEU-REPOSITORIO.git)
    cd NOME-DA-PASTA
    ```

2.  **Instale as dependências:**
    ```bash
    pip install pandas openpyxl reportlab
    ```
    *Nota: `openpyxl` é necessário para que o Pandas possa ler e escrever em arquivos `.xlsx`.*

3.  **Estrutura de arquivos:**
    Certifique-se de que a estrutura de pastas está correta. O script espera encontrar:
    -   Um diretório `BD/` contendo o arquivo `PLAN_ITENS.xlsx`.
    -   Um diretório `Relatorios/` onde os PDFs gerados serão salvos.

4.  **Execute o script principal:**
    ```bash
    python nome_do_seu_arquivo.py
    ```