from reportlab.lib import colors
from reportlab.platypus import Table
from reportlab.lib.pagesizes import letter
import datetime
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.pagesizes import letter, landscape
from datetime import datetime


''' CONVENÇÃO
Variável _ primeira letra minúscula (numero_De_Entrada)
Constante - tudo maisculo (VALOR_DE_ENTRADA)
Método - primeira letra maiúscula de cada palavra (Valor_De_Entrada)'''


planItens_df = pd.read_excel("BD\PLAN_ITENS.xlsx")

Banco_De_Dados = planItens_df[['item ', 'quantidade_De_Entrada']].set_index(
    'item ').to_dict()['quantidade_De_Entrada']
Valor_Unitario_Item = planItens_df[['item ', 'valor_unitario']].set_index(
    'item ').to_dict()['valor_unitario']

# dicionarios criados para gerar o banco de dicionarios para poder guardar os dicionarios salvos.
Banco_De_Dicionarios = {
    "cadeira001": {'encosto_plastico': 1,
                   'acento_plastico': 1,
                   'armacao': 1,
                   'parafuso_Philips': 12},
    "cadeira002": {'encosto_Plus': 1,
                   'acento_Plus': 1,
                   'armacao': 1,
                   'parafuso_Philips': 16}
}

# Banco de dicionarios para ser usado no relatorio.
Descrições_De_Produtos_E_Insumos = {"Produtos": {}, "Insumos": {}}


# FUNÇÕES:

def Menu_Principal():  # nao foi alterada com atuazalização do excel
    '''-> Função: Exibe um menu com 4 opções:
    1 - Cadastro e eliminação de items
    2 - Atualização de estoque
    3 - Fabricação de produto
    4 - Imprimir Relatorios
    5 - Sair'''

    print()
    print('-='*16)
    print(" MENU PRINCIPAL ".center(30, ' '))
    print('-='*16)
    print("ESCOLHA UMA OPÇÃO:")
    print(" 1 - Cadastro e eliminação de items \n 2 - Atualização de estoque \n 3 - Fabricação de produto \n 4 - Imprimir Relatorios \n 5 - Sair")
    escolha = input("Opção desejada: ").strip()
    return escolha


def Menu_Cadastro_E_Eliminação():  # nao foi alterada com atualização do excel
    """
    Pede para o usuário escolher entre as opções:
    1 - Cadastro de produtos 
    2 - Cadastrar insumos necessários para produção de determinado produto 
    3 - Cadastro de insumos
    4 - Eliminação de produtos 
    5 - Eliminação de insumos 
    6 - Voltar ao Menu Principal"
    """

    continuar = True
    while continuar:
        print()
        print('-'*40)
        print(" CADASTRO E ELIMINAÇÃO DE PRODUTOS E INSUMOS ".center(40, ' '))
        print('-'*40)
        print("ESCOLHA UMA OPÇÃO:")
        print(" 1 - Cadastro de produtos \n 2 - Cadastro de insumos \n 3 - Eliminação de produtos \n 4 - Eliminação de insumos \n 5 - Voltar ao Menu Principal")

        escolha = input("Opção desejada: ").strip()
        if escolha == "1":
            Cadastro_De_Produto()

        elif escolha == "2":

            Cadastro_De_Insumo()

        elif escolha == "3":
            Eliminar_Produto()

        elif escolha == "4":
            Eliminar_Insumo()

        elif escolha == "5":
            print()
            print("Voltando ao Menu Principal ...")
            print()
            continuar = False
        else:
            print()
            print(f'\033[31mOpção inválida.\033[m Escolha um valor de 1 a 5.')
            print()
    return


def Cadastro_De_Insumo():  # DEF ATUALIZADA COM PANDAS UTILIZANDO A PLAN_ITENS
    """
    Cadastro de novos insumos: Pergunta o nome, a descrição, a quantidade do insumo, a data de criação e o valor.

    Armazenamento de dados:
    Nome, quantidade, descrição, data de criação e valor -> DataFrame
    """
    planItens_df = pd.read_excel("BD\PLAN_ITENS.xlsx")

    print()
    print('-'*30)
    print("CADASTRO DE INSUMO OU EMBALAGEM".center(30))
    print('-'*30)

    # Inicializa planItens_df como um DataFrame vazio ou lê o arquivo Excel se existir
    try:
        planItens_df = pd.read_excel("BD\PLAN_ITENS.xlsx")
    except FileNotFoundError:
        print("Arquivo 'PLAN_ITENS.xlsx' não encontrado. Criando um DataFrame vazio.")
        planItens_df = pd.DataFrame(
            columns=['item ', 'descricao', 'valor', 'data_De_Criacao'])

    while True:
        nome_do_insumo = input("Nome do Insumo:").strip().capitalize()

        if nome_do_insumo:
            descricao_do_insumo = input("Descrição:").strip().capitalize()

            if descricao_do_insumo:
                quantidade = input("Quantidade: ").strip()

                if quantidade.isdigit():  # Verifica se a quantidade é um número
                    quantidade = int(quantidade)

                    # Adiciona um novo registro ao DataFrame existente

                    planItens_df.loc[len(planItens_df)] = {
                        'item ': nome_do_insumo,
                        'descricao': descricao_do_insumo,
                        'valor': float(input("Valor: ").strip()),
                        'data_De_Criacao': pd.to_datetime(datetime.now().strftime('%d/%m/%Y %H:%M:%S'), format='%d/%m/%Y %H:%M:%S'),
                        'quantidade_De_Entrada': quantidade
                    }

                    continuar = input(
                        "Deseja cadastrar mais insumos? (S/N): ").strip().lower()

                    if continuar != 's':
                        break
                else:
                    print("\033[31mErro: Valor digitado não é um número.\033[m")
            else:
                print("\033[31mErro: Descrição vazia.\033[m")
        else:
            print("\033[31mErro: Nome vazio.\033[m")

    planItens_df.to_excel("BD\PLAN_ITENS.xlsx", index=False)
    # calcular_valor_unitario()


def Cadastro_De_Produto():
    """
    Cadastro de novos produtos: Pergunta o nome, a descrição e a quantidade do produto. Além disso, chama a função Cadastrar_Insumos_Necessarios() para cadastrar os insumos necessários para a fabricação do produto cadastrado.

    Armazenamento de dados: 
    Nome e quantidade -> Banco_De_Dados["Produtos"]
    Nome e descrição -> Descrições_De_Produtos_E_Insumos["Produtos"]
    """

    continuar = True
    while continuar:
        print()
        print('-'*30)
        print("CADASTRO DE PRODUTO".center(30))
        print('-'*30)
        nome_Do_Produto = input("Nome do Produto:").strip().capitalize()

        if nome_Do_Produto != "":  # tratamento de dados para entrada vazia
            while continuar:
                descrição_Do_produto = input("Descrição:").strip()

                if descrição_Do_produto != "":

                    Cadastrar_Insumos_Necessarios(nome_Do_Produto)

                    # Adiciona o nome do insumo e sua descrição no dicionário
                    Descrições_De_Produtos_E_Insumos["Produtos"][nome_Do_Produto] = descrição_Do_produto

                    continuar = False

                else:
                    print("\033[31mErro: Entrada vazia.\033[m")
        else:
            print("\033[31mErro: Entrada vazia.\033[m")


def Cadastrar_Insumos_Necessarios(NOME_PRODUTO):
    """
    Cadastra os insumos necessários para a produção de determinado produto.
    Obs.: Pode ser chamada pela função Cadastro_De_Produto(), ou pelo Menu_Cadastro_E_Eliminação().

    Param NOME_PRODUTO: nome do produto cadastrado.
    """

    print()
    print('-'*40)
    print(f'CADASTRAR INSUMOS NECESSÁRIOS PARA FABRICAÇÃO DE "{NOME_PRODUTO}"')
    print('-'*40)
    print()

    Parcial_Insumos_Necessarios = {}
    cont1 = 1
    continuar1 = True
    while continuar1:

        print()
        print(f"Selecione o {cont1}° insumo necessário: ")
        cont2 = 0
        for keys in Banco_De_Dados.keys():
            cont2 += 1
            print(cont2, "-", keys)
        print()

        numero_Do_Insumo = input('Insumo necessário: ').strip()

        if (numero_Do_Insumo != "") and (numero_Do_Insumo.isnumeric() == True):
            numero_Do_Insumo = int(float(numero_Do_Insumo)//1)

            if 0 < numero_Do_Insumo <= len(Banco_De_Dados):
                numero_Do_Insumo = int(float(numero_Do_Insumo)//1)

                # Cria uma lista com os nomes das chaves dos dicionarios
                lista_keys_Banco_De_Dados = list(Banco_De_Dados.keys())

                # Pega o nome do dicionario com base no indice
                nome_Do_Insumo = lista_keys_Banco_De_Dados[numero_Do_Insumo-1]

                continuar2 = True
                while continuar2:
                    quantidade = input(
                        f'Quantidade necessária do insumo "{nome_Do_Insumo}": ').strip()

                    if (quantidade != "") and (quantidade.isnumeric() == True):

                        Parcial_Insumos_Necessarios[nome_Do_Insumo] = float(
                            quantidade)

                        # Mostrar a seleção parcial dos insumos e pergunta se o usuário deseja continuar
                        continuar1 = COMPLEMENTAR_Cadastrar_Insumos_Necessários(
                            Parcial_Insumos_Necessarios, NOME_PRODUTO)
                        continuar2 = False
                        cont1 += 1

                    else:
                        print(
                            "\033[31mErro: Entrada vazia ou valor digitado não númerico.\033[m")
            else:
                print(
                    f'\033[31mValor Inválido. Favor digitar um valor de 1 a {len(Banco_De_Dados)+1} \033[m')
        else:
            print(
                "\033[31mErro: Entrada vazia ou valor digitado não númerico.\033[m")


def COMPLEMENTAR_Cadastrar_Insumos_Necessários(parcial, NOME_PRODUTO):
    """
    def complementar a def Cadastrar_Insumos_Necessarios()
    -> Função: Mostra a seleção parcial dos insumos necessarios para a produção de um produto e pergunta se o usuario deseja continuar adicionando novos insumos, caso não, finaliza a def cadastrar_Insumos_Necessarios()

    Saída: Armazena os insumos necessários no dicionario insumos_necessarios().
    """

    print()
    print('-'*40)
    print("Insumo: ".center(20), "Quantidade".ljust(20))
    for insumo, quantidade in parcial.items():
        print(f"{insumo.ljust(20)} {quantidade:^20.2f}")
    print('-'*40)
    print()

    print("Deseja adicionar mais insumos? \n 1 - Sim \n 2 - Não")
    while True:
        resposta = input('Sua resposta: ').strip()
        if (resposta != "") and (resposta.isnumeric() == True):
            resposta = int(float(resposta)//1)

            if resposta == 1:
                return True

            elif resposta == 2:
                print()
                print("Deseja salvar as alterações? \n 1 - Sim \n 2 - Não")

                while True:
                    escolha = input('Sua decisão: ').strip()

                    if (escolha != "") and (escolha.isnumeric() == True):
                        escolha = int(float(escolha)//1)

                        if escolha == 1:
                            # nome_Do_Produto = str(NOME_PRODUTO)
                            dicionario_parcial = parcial.copy()
                            Banco_De_Dicionarios[NOME_PRODUTO] = dicionario_parcial

                            print()
                            print("ALTERAÇÕES SALVAS COM SUCESSO!✅")
                            print()
                            return False

                        else:
                            "ALTERAÇÕES NÃO SALVAS!❌"
                            return False

                    else:
                        print(
                            "\033[31mErro: Entrada vazia ou valor digitado não númerico \033[m")
            else:
                print(
                    f'\033[31mValor Inválido. Favor digitar um valor de 1 a 2 \033[m')
        else:
            print(
                "\033[31mErro: Entrada vazia ou valor digitado não númerico \033[m")


def Menu_Eliminar_Item():  # def não atualizada com excel
    """
    Menu de eliminação de produtos e insumos.
    Pede para o usuário escolher se que excluir itens de produção ou insumos.
    """

    print()
    print('-'*30)
    print("ELIMINAÇÃO DE ITENS DE PRODUÇÃO ou INSUMOS:".center(30))
    print('-'*30)
    print('Escolha uma opção: ')
    print(' 1 - Excluir itens de produção \n 2 - Excluir insumos \n 3 - Voltar o Menu Principal')
    continuar = True
    while continuar:
        escolha = input('Opção desejada: ').strip()
        if 0 < escolha <= 3:
            if escolha == 3:
                continuar = False

            elif escolha == 1:
                Eliminar_Insumo()

            else:  # escolha ==2
                Eliminar_Produto()
        else:
            print(
                f'\033[31mValor Inválido. Favor digitar um valor de 1 a 3 \033[m')


def Eliminar_Insumo():  # DEF ATUALIZADA COM EXCEL

    global planItens_df  # Adiciona esta linha para acessar a variável global
    planItens_df = pd.read_excel("BD\PLAN_ITENS.xlsx")

    print()
    print("-"*40)
    print("ELIMINAÇÃO DE INSUMO".center(40))
    print("-"*40)
    print('Escolha o insumo que deseja eliminar: ')

    lista_keys_banco_de_dicionarios = list(Banco_De_Dados.keys())

    cont = 0
    for keys in lista_keys_banco_de_dicionarios:
        cont += 1
        print(cont, "-", keys)
    print("ou".center(10))
    print(f"{cont+1} - VOLTAR PARA O MENU DE CADASTRO E ELIMINAÇÃO")

    continuar = True
    while continuar:
        print()
        numero_Do_Insumo = input('Opção desejada: ')
        print()

        if (numero_Do_Insumo != "") and (numero_Do_Insumo.isnumeric() == True):
            numero_Do_Insumo = int(float(numero_Do_Insumo)//1)

            if numero_Do_Insumo == (len(lista_keys_banco_de_dicionarios))+1:
                print("Voltando ao Menu de Cadastro e Eliminação ...")
                continuar = False

            elif 0 < numero_Do_Insumo <= len(lista_keys_banco_de_dicionarios):
                item_a_eliminar = lista_keys_banco_de_dicionarios[numero_Do_Insumo-1]

                quantidade_entrada = planItens_df.loc[planItens_df['item ']
                                                      == item_a_eliminar, 'quantidade_De_Entrada'].values[0]

                if quantidade_entrada > 0:
                    print(
                        f" Ops... O insumo '{item_a_eliminar}' possui quantidade de entrada maior que zero e não pode ser eliminado.")
                else:
                    print(
                        f'ATENÇÃO: Deseja excluir "{item_a_eliminar}" do estoque? Após executada, a ação não poderá ser desfeita.')
                    print(" 1 - Sim, Estou ciente das implicações \n 2 - Não.")

                    while continuar:
                        escolha = input('Decisão final: ').strip()

                        if (escolha != "") and (escolha.isalpha() == False):
                            escolha = int(float(escolha)//1)

                            # Verifica se há quantidade na coluna 'quantidade_De_Entrada'
                            if escolha == 1:
                                # Remove a linha correspondente ao insumo
                                planItens_df = planItens_df[planItens_df['item ']
                                                            != item_a_eliminar]

                                # Salva o DataFrame atualizado na planilha
                                planItens_df.to_excel(
                                    "BD\PLAN_ITENS.xlsx", index=False)

                                print(
                                    f"O insumo '{item_a_eliminar}' foi eliminado com sucesso.✅")
                                continuar = False

                            elif escolha == 2:
                                print(
                                    "Voltando ao Menu de Cadastro e eliminação ...")
                                continuar = False
                        else:
                            print()
                            print(
                                f'\033[31mValor Inválido. Favor digitar 1 ou 2 \033[m')
                            print()
            else:
                print()
                print(
                    f'\033[31mValor Inválido. Favor digitar um valor de 1 a {len(Banco_De_Dados.keys())+1}\033[m')
                print()
        else:
            print("\033[31mErro: Entrada vazia ou valor não númerico.\033[m")


def Eliminar_Produto():  # DEF ATUALIZADA COM EXCEL
    """
    Exclui determinado produto do Banco_De_Dados.
    Obs.: Só exclui se o estoque do produto estiver vazio.
    """

    print()
    print("-"*40)
    print("ELIMINAÇÃO DE PRODUTOS".center(40))
    print("-"*40)
    print('Escolha o produto que deseja eliminar: ')

    lista_keys_banco_de_dicionarios = list(Banco_De_Dicionarios.keys())
    cont = 0
    for keys in lista_keys_banco_de_dicionarios:
        cont += 1
        print(cont, "-", keys)
    print("ou".center(10))
    print(f"{cont+1} - VOLTAR PARA O MENU DE CADASTRO E ELIMINAÇÃO")

    continuar = True
    while continuar:
        numero_Do_Produto = input('Opção desejada: ').strip()

        if (numero_Do_Produto != "") and (numero_Do_Produto.isnumeric() == True):
            numero_Do_Produto = int(float(numero_Do_Produto)//1)

            if numero_Do_Produto == (len(lista_keys_banco_de_dicionarios))+1:
                print("Voltando ao Menu de Cadastro e Eliminação ...")
                continuar = False

            elif 0 < numero_Do_Produto <= len(lista_keys_banco_de_dicionarios):
                nome_Produto_Desejado = lista_keys_banco_de_dicionarios[numero_Do_Produto-1]

                print(
                    f'ATENÇÃO: Deseja excluir "{nome_Produto_Desejado}" do estoque? Após executada, a ação não poderá ser desfeita.')
                print(" 1 - Sim, Estou ciente das implicações \n 2 - Não.")

                while continuar:
                    escolha = input('Decisão final: ').strip()

                    if (escolha != "") and (escolha.isalpha() == False):
                        escolha = int(float(escolha)//1)

                        if escolha == 1:
                            del Banco_De_Dicionarios[nome_Produto_Desejado]
                            print()
                            print(
                                f'"{nome_Produto_Desejado}" eliminado com sucessso.✅')

                            continuar = False

                        elif escolha == 2:
                            print(
                                "Voltando ao Menu de Cadastro e eliminação ...")
                            continuar = False
                        else:
                            print()
                            print(
                                f'\033[31mValor Inválido. Favor digitar 1 ou 2 \033[m')
                            print()
                    else:
                        print(
                            "\033[31mErro: Entrada vazia ou não numerica\033[m")
            else:
                print()
                print(
                    f'\033[31mValor Inválido. Favor digitar um valor de 1 a {len(Banco_De_Dados)+1}\033[m')
                print()
        else:
            print("\033[31mErro: Entrada vazia ou valor não númerico.\033[m")


def Atualizacao_De_Estoque():  # DEF ATUALIZADA COM EXCEL
    global planItens_df
    planItens_df = pd.read_excel("BD\PLAN_ITENS.xlsx")

    print()
    print("-"*40)
    print("ATUALIZAÇÃO DE ESTOQUE".center(40))
    print("-"*40)
    print('Escolha o item que deseja atualizar: ')

    lista_keys_banco_de_dicionarios = list(Banco_De_Dados.keys())

    cont = 0
    for keys in lista_keys_banco_de_dicionarios:
        cont += 1
        print(cont, "-", keys)
    print("ou".center(10))
    print(f"{cont+1} - VOLTAR PARA O MENU DE PRINCIPAL")

    continuar = True
    while continuar:
        print()
        numero_Do_Insumo = input('Opção desejada: ')
        if (numero_Do_Insumo != "") and (numero_Do_Insumo.isnumeric() == True):
            numero_Do_Insumo = int(float(numero_Do_Insumo)//1)

            if numero_Do_Insumo == (len(lista_keys_banco_de_dicionarios))+1:
                print("Voltando ao Menu de Principal ...")
                continuar = False

            elif 0 < numero_Do_Insumo <= len(lista_keys_banco_de_dicionarios):
                item_a_atualizar = lista_keys_banco_de_dicionarios[numero_Do_Insumo-1]

                # Pede ao usuário a nova quantidade
                nova_quantidade = input("Digite a nova quantidade: ").strip()

                # Verifica se a entrada é válida (não vazia e numérica)
                if nova_quantidade != "" and nova_quantidade.isnumeric():
                    nova_quantidade = int(nova_quantidade)

                    # Atualiza a quantidade no DataFrame
                    planItens_df.loc[planItens_df['item '] == item_a_atualizar,
                                     'quantidade_De_Entrada'] = nova_quantidade

                    # Atualiza a data_De_Criacao com a data e hora atual
                    planItens_df.loc[planItens_df['item '] == item_a_atualizar, 'data_De_Criacao'] = pd.to_datetime(
                        datetime.now().strftime('%d/%m/%Y %H:%M:%S'), format='%d/%m/%Y %H:%M:%S')

                    # Salva o DataFrame atualizado na planilha
                    planItens_df.to_excel("BD\PLAN_ITENS.xlsx", index=False)

                    print(
                        f"A quantidade do item '{item_a_atualizar}' foi atualizada para {nova_quantidade} com sucesso.✅")
                    continuar = False
                else:
                    print("Entrada inválida. A quantidade deve ser um valor numérico.")

            else:
                print()
                print(
                    f'\033[31mValor Inválido. Favor digitar um valor de 1 a {len(Banco_De_Dados.keys())+1}\033[m')
                print()
        else:
            print("\033[31mErro: Entrada vazia ou valor não númerico.\033[m")
    print()
    print("Voltando ao Menu Principal...")


def Seleção_De_Produtos():
    '''-> Mostra a listagem dos produtos cadastrados no dicionario banco_De_Dados.'''

    print()
    print("-"*30)
    print(" FABRICAÇÃO DE PRODUTO ".center(30, ' '))
    print("-"*30)
    print("ESCOLHA O PRODUTO QUE DESEJA FABRICAR:")
    cont = 0

    for keys in Banco_De_Dicionarios.keys():
        cont += 1
        print(cont, "-", keys)

    print("ou".center(10))
    print(f"{cont+1} - VOLTAR PARA MENU PRINCIPAL")
    print()


def Saudacao():
    """-> Função: Retorna uma saudação de despedida de acordo com o horário.
    Return: 
    "Tenha um bom dia!😊" ou 
    "Tenha uma boa tarde!😊" ou
    "Tenha uma boa noite!😊".
    """

    # Obtém a hora atual
    hora_atual = datetime.now().hour

    # Define as mensagens de saudação
    if 6 <= hora_atual < 12:
        return "Tenha um bom dia!😊"
    elif 12 <= hora_atual < 18:
        return "Tenha uma boa tarde!😊"
    else:
        return "Tenha uma boa noite!😊"


# DEF PARA GERAR RELATORIO DE PRODUTO, ELE MOSTRA A ESTRUTURA DE CADA ITEM.
def Relatorio_Produto_Venda_PDF(dicionario_produtos, caminho_salvar):
    pdf_filename = f"{caminho_salvar}\\Relatorio_Produtos.pdf"

    # Criando o PDF
    c = canvas.Canvas(pdf_filename)

    # Adicionando a data no topo
    data_atual = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    c.drawString(100, 800, f"Data de criação: {data_atual}")
    c.drawString(100, 780, "-"*70)

    c.drawString(100, 760, "Relatório de Produtos:")
    c.drawString(100, 740, "-"*70)

    # Adicionando os itens ao PDF
    y_position = 730
    for codigo_produto, dados_produto in dicionario_produtos.items():
        c.drawString(100, y_position, f"Produto: {codigo_produto}")
        y_position -= 30

        for item, quantidade in dados_produto.items():
            c.drawString(120, y_position, f"{item.ljust(30)} {quantidade}")
            y_position -= 30

        c.drawString(100, y_position, "-"*70)
        y_position -= 30

    c.save()


# DEF PARA GERAR O RELATÓRIO DE ITENS QUE TEMOS COMO INSUMOS.
def Relatorio_de_Insumos_PDF():
    # atualização de dataframe
    planItens_df = pd.read_excel("BD\PLAN_ITENS.xlsx")

    # Inicializar o PDF
    pdf_path = "Relatorios\Relatorio_de_Insumos.pdf"
    c = canvas.Canvas(pdf_path, pagesize=landscape(A4))

    # Adicionar cabeçalho
    c.setFont("Helvetica", 16)
    c.drawCentredString(110, 520, "Relatório de Itens")

    # Adicionar tabela
    width, height = landscape(A4)
    data_atual = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
    tabela = [['Item', 'Descrição', 'Valor', 'Quantidade de Entrada',
               'Data de Criação']]

    for index, row in planItens_df.iterrows():
        tabela.append([row['item '], row['descricao'], row['valor'], row['quantidade_De_Entrada'],
                       row['data_De_Criacao']])

    table_style = [('GRID', (0, 0), (-1, -1), 1, colors.black)]
    table = Table(tabela, style=table_style, colWidths=150,
                  rowHeights=25, hAlign='CENTER')

    # Posicionar a tabela no PDF
    table.wrapOn(c, width, height)
    table.drawOn(c, 50, height - 260)

    # Adicionar rodapé
    c.setFont("Helvetica", 10)
    c.drawCentredString(width / 2, 30, f"Gerado em: {data_atual}")

    # Salvar o PDF
    c.save()


def Fabricacao_Produto():
    """
    -> Produz uma certa quantidade de um determinado produto.

    Entrada: nome do produto e quantidade.
    Saída: mostra se a quantidade de insumo, presente no estoque, é suficiente.
    """

    Seleção_De_Produtos()
    continuar = True
    while continuar:
        lista_keys_produtos = list(Banco_De_Dicionarios.keys())
        # Add tratamento p/ entrada vazia, simbolos e letras."É o produto desejado? [S/N]"
        numero_Do_Produto = int(float(input('Opção desejada: ').strip())//1)
        if numero_Do_Produto == len(lista_keys_produtos)+1:
            print("Voltando ao Menu Principal ...")
            continuar = False

        elif 0 < numero_Do_Produto <= len(lista_keys_produtos):
            # Extrai o nome do produto, de dentro do dicionario, por meio do indice.
            produto_a_fabricar = lista_keys_produtos[numero_Do_Produto-1]

            # Obtém o dicionário do produto
            dicionario_produto = Banco_De_Dicionarios[produto_a_fabricar]

            # Pede ao usuário a quantidade desejada
            quantidade_desejada = float(
                input("Digite a quantidade desejada: "))

            # Mostra se tem a qtdd suficiente de insumos.
            produzir_OU_Não = Comparacao_Insumo(
                produto_a_fabricar, quantidade_desejada)

            if produzir_OU_Não == "os insumos são suficientes":
                # Inicializa o valor total do produto
                valor_Total_Do_Produto = 0

                # Multiplica os valores correspondentes nos dicionários
                for chave, valor_cadeira in dicionario_produto.items():
                    if chave in Valor_Unitario_Item:
                        valor_Unitario = Valor_Unitario_Item[chave]
                        valor_Total_Do_Produto += valor_cadeira * valor_Unitario

                # Calcula o valor total com base na quantidade desejada
                valor_total = valor_Total_Do_Produto * quantidade_desejada

                # Exibe os resultados
                print()
                print(
                    f"O valor total para uma unidade do produto {produto_a_fabricar} é: R${valor_Total_Do_Produto:.2f}")
                print(
                    f"O valor total para {quantidade_desejada} unidades é: R${valor_total:.2f}")

                continuar = False
            else:
                print()
                print("Os insumos não são suficientes. Favor, pedir mais insumos.")
                print()
        else:
            print()
            print(
                f'\033[31mValor Inválido. Favor digitar um valor de 1 a {len(lista_keys_produtos)+1}\033[m')
            print()


def Comparacao_Insumo(produto, quantidade):
    """
    Calcula se a quantidade de insumo que existe no estoque é suficiente para executar a produção de determinada quantidade de um produto.

    Param produto: Nome do produto.
    Param quantidade: Quantidade do produto que se dejesa produzir.
    """
    verificador = 0

    nome_Insumo_E_Quantidade_Total = dict()
    insumos_E_Quantidade_Q_Falta = dict()
    print()
    print('-'*70)
    print('-> STATUS DO ESTOQUE:')
    for nome_Do_Insumo, quantidade_Do_Insumo in Banco_De_Dicionarios[produto].items():
        quantidade_Total_Do_Insumo = quantidade_Do_Insumo * quantidade

        if Banco_De_Dados[nome_Do_Insumo] >= quantidade_Total_Do_Insumo:
            print(f"{nome_Do_Insumo.ljust(26, ' ')} {(f'Qtdd necessária: {quantidade_Total_Do_Insumo},').ljust(30, ' ')} status: Suficiente✅")

            nome_Insumo_E_Quantidade_Total[nome_Do_Insumo] = quantidade_Do_Insumo
            verificador += 1

        else:
            quantidade_Q_Falta = quantidade_Total_Do_Insumo - \
                Banco_De_Dados[nome_Do_Insumo]
            insumos_E_Quantidade_Q_Falta[nome_Do_Insumo] = quantidade_Q_Falta

            print(
                f"{nome_Do_Insumo.ljust(26, ' ')} {(f'Qtdd necessária: {quantidade_Total_Do_Insumo},').ljust(30, ' ')} status: Insuficiente❌", end=' ')
            print(
                f'// Falta(m) {quantidade_Q_Falta} unidade(s).')
    print('-'*70)
    if verificador == len(Banco_De_Dicionarios[produto].keys()):
        return "os insumos são suficientes"
        # Mandar_produzir(produto, quantidade)
    else:
        return "insumos insuficientes"

        # Pedir_Mais_Insumo(insumos_E_Quantidade_Q_Falta)


# PROGRAMA PRINCIPAL
def Principal():
    """
    Menu principal que comanda todo o programa.
    Chama direta ou, indiretamente, todas as def's.
    """

    continuar = True
    while continuar:
        escolha = Menu_Principal()

        if escolha == "1":
            Menu_Cadastro_E_Eliminação()

        elif escolha == "2":
            Atualizacao_De_Estoque()

        elif escolha == "3":
            Fabricacao_Produto()

        elif escolha == "4":
            Menu_De_Relatorio()

        elif escolha == "5":
            print()
            print(f"PROGRAMA FINALIZADO.\n{Saudacao()}")
            print()
            continuar = False

        else:
            print()
            print(f'\033[31mOpção inválida.\033[m Escolha um valor de 1 a 5.')
            print()


def Menu_De_Relatorio():
    print()
    print("-"*30)
    print(" MENU DE RELATORIO ". center(30, ' '))
    print("-"*30)

    # Adicionar opção para voltar ao menu principal
    print('1 - Relatorio - Insumo \n2 - Relatorio - Produto Venda \n3 - VOLTAR PARA MENU PRINCIPAL')

    continuar = True
    while continuar:
        escolha = int(input('Opção desejada: '))
        if escolha == 3:
            print("Voltando ao Menu Principal ...")
            continuar = False

        elif escolha == 1:
            Relatorio_de_Insumos_PDF()
            continuar = False

        elif escolha == 2:
            Relatorio_Produto_Venda_PDF(Banco_De_Dicionarios, "Relatorios")
            continuar = False

        else:
            print(f'\033[31mOpção inválida.\033[m Escolha um valor de 1 a 3.')


# PROGRAMA PRICIPAL
Principal()
