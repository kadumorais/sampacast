import pandas as pd
import os

arquivo_entrada = r"C:\Users\Carlos Morais\Documents\Projetos\confirmação_sampacast\agenda.xlsx"
pasta_saida = r'C:\Users\Carlos Morais\Documents\Projetos\confirmação_sampacast\out'

def agrupar_usuário():
    if os.path.exists(arquivo_entrada):
        df = pd.read_excel(arquivo_entrada)

        if 'Unidade' not in df.columns:
            print("A coluna 'Unidade' não está presente no DataFrame original.")
            return

        df['Quantidade de Datas'] = 1

        grouped = df.groupby(['Unidade', 'Usuário - Documento']).agg({
            'Usuário - Nome': 'first',
            'Data do Agendamento': lambda x: ', '.join(x.astype(str)),
            'Quantidade de Datas': 'sum',
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

        columns_order = ['Unidade', 'Usuário - Nome', 'Data do Agendamento', 'Quantidade de Datas', 'Usuário - Email', 'Usuário - Documento', 'Usuário - Data Nascimento', 'Usuário - CEP', 'Usuário - Rua', 'Usuário - Gênero', 'Usuário - Nacionalidade', 'Usuário - Deficiência', 'Usuário - Renda Mensal', 'Usuário - Formação Escolar', 'Usuário - Cor da Pele']
        grouped = grouped[columns_order]

        if not os.path.exists(pasta_saida):
            os.makedirs(pasta_saida)

        for unidade, dados in grouped.groupby('Unidade'):
            nome_arquivo = os.path.join(pasta_saida, f'dados_{unidade}.xlsx')
            dados.to_excel(nome_arquivo, index=False)

agrupar_usuário()