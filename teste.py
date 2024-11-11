import pandas as pd
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
            name = str(row['Usuário - Nome']).split()[0].title()
            phone_number = str(row['Usuário - Telefone'])

            message = f"Bom dia! Tudo bem, {name}? Nós da ADE SAMPA estamos entrando em contato para cancelar sua agenda de Segunda feira. Devido ao feriado, foi estipulado ponto facultativo para as unidades onde se encontram as unidades Sampa Cast. Sendo assim, não conseguimos estipular uma data de reagendamento. Lamentamos e esperamos você em uma outra oportunidade!\nPara mais informações sobre o projeto Sampa Cast: https://adesampa.com.br/sampacast/"

            send_whatsapp_message(phone_number, message)

    except Exception as e:
        print(f"Ocorreu um erro ao processar o arquivo Excel: {e}")

if __name__ == "__main__":
    file_path = r'C:\Users\Carlos Morais\Documents\Projetos\sampacast/agenda.xlsx'

    send_messages_from_excel(file_path)