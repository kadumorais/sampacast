import time
import pandas as pd
from datetime import date, datetime
from urllib.parse import quote


def generate_whatsapp_link(phone_number, message):
    phone_number = ''.join(filter(str.isdigit, phone_number))

    encoded_message = quote(message)

    return f"{phone_number} : https://wa.me/55{phone_number}/?text={encoded_message} \n"

def send_whatsapp_message(phone_number, message):
    whatsapp_link = generate_whatsapp_link(phone_number, message)
    print(f"{whatsapp_link}")

def send_messages_from_excel(file_path):
    try:
        df = pd.read_excel(file_path)
        for index, row in df.iterrows():
            name = str(row['Usuário - Nome'])
            phone_number = str(row['Usuário - Telefone'])
            unid = str(row['Unidade'])
            hora = datetime.strptime(row['Horário Início'], '%H:%M:%S')

            horario = hora.strftime('%H:%M')

            message = f"Bom dia! Tudo bem, {name}? Nós da ADE SAMPA estamos entrando em contato para cancelar seu agendamento na {unid} no dia 18/10. Poderia me informar uma data, horário e seu email na plataforma para reagendar sua gravação para Novembro? Para mais informações sobre o projeto Sampa Cast: https://adesampa.com.br/sampacast/"

            send_whatsapp_message(phone_number, message)

    except Exception as e:
        print(f"Ocorreu um erro ao processar o arquivo Excel: {e}")

if __name__ == "__main__":
    file_path = r'C:\Users\Carlos Morais\Documents\Projetos\confirmação_sampacast/agenda.xlsx'

    send_messages_from_excel(file_path)