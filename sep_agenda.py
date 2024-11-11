import pandas as pd
import os

# Caminhos de entrada e saída
arquivo_entrada = r"C:\Users\Carlos Morais\Documents\Projetos\sampacast\agenda.xlsx"
pasta_saida = r'C:\Users\Carlos Morais\Documents\Projetos\sampacast\out'

def agrupar_usuário():
    # Verifica se o arquivo de entrada existe
    if not os.path.exists(arquivo_entrada):
        print(f"Arquivo de entrada '{arquivo_entrada}' não encontrado.")
        return

    # Lê o arquivo Excel
    df = pd.read_excel(arquivo_entrada)

    # Verifica se a coluna 'Unidade' está presente
    colunas_obrigatorias = ['Unidade', 'Usuário - Documento', 'Usuário - Nome', 'Data do Agendamento']
    for coluna in colunas_obrigatorias:
        if coluna not in df.columns:
            print(f"A coluna '{coluna}' não está presente no DataFrame original.")
            return

    # Adiciona uma nova coluna para contar as datas
    df['Quantidade de Datas'] = 1

    # Agrupa os dados pelas colunas especificadas
    grouped = df.groupby(['Unidade', 'Usuário - Documento']).agg({
        'Usuário - Nome': 'first',  # Mantém o primeiro nome encontrado
        'Data do Agendamento': lambda x: ', '.join(x.dropna().astype(str)),  # Combina as datas em uma string
        'Quantidade de Datas': 'sum',  # Soma a quantidade de agendamentos
        'Usuário - Email': 'first',
        'Usuário - Data Nascimento': 'first',
        'Usuário - CEP': 'first',
        'Usuário - Rua': 'first',
        'Usuário - Gênero': 'first',
        'Usuário - Nacionalidade': 'first',
        'Usuário - Deficiência': 'first',
        'Usuário - Renda Mensal': 'first',
        'Usuário - Formação Escolar': 'first',
        'Usuário - Cor da Pele': 'first'
    }).reset_index()

    # Reorganiza as colunas na ordem desejada
    columns_order = [
        'Unidade', 'Usuário - Nome', 'Data do Agendamento', 'Quantidade de Datas', 
        'Usuário - Email', 'Usuário - Documento', 'Usuário - Data Nascimento', 
        'Usuário - CEP', 'Usuário - Rua', 'Usuário - Gênero', 'Usuário - Nacionalidade', 
        'Usuário - Deficiência', 'Usuário - Renda Mensal', 'Usuário - Formação Escolar', 
        'Usuário - Cor da Pele'
    ]
    grouped = grouped[columns_order]

    # Cria a pasta de saída se não existir
    if not os.path.exists(pasta_saida):
        os.makedirs(pasta_saida)

    # Salva cada grupo por unidade em arquivos Excel separados
    for unidade, dados in grouped.groupby('Unidade'):
        nome_arquivo = os.path.join(pasta_saida, f'dados_{unidade}.xlsx')
        dados.to_excel(nome_arquivo, index=False)
        print(f"Arquivo '{nome_arquivo}' salvo com sucesso.")

# Chama a função para agrupar e salvar os dados
agrupar_usuário()
