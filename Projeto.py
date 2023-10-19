import smtplib
import pandas as pd
import os

#caminho da base que será lida 
caminho_arquivo = "D:\\Espaço_de_projetos\\ETL_DIO_Projeto\\dados.xlsx"

# Lê o arquivo Excel usando o pandas
try:
    # Assume que a primeira folha do arquivo Excel será lida
    df = pd.read_excel(caminho_arquivo, engine='openpyxl')
    
    # Extrai o nome completo do jogador da coluna full_name
    full_names = df['full_name'].tolist()

    # Limpa os nomes repetidos deixando apenas um nome
    full_names = list(dict.fromkeys(full_names))

    ### Transformação de dados ###

    # Criando um dataframe para colocar os nomes
    df_nomes = pd.DataFrame(columns=['full_name'])

    # Armazenando o dataframe
    df_nomes['full_name'] = full_names

    # Salvando o dataframe em um novo arquivo Excel com os nomes limpos
    df_nomes.to_excel('nomes_limpos.xlsx', index=False)

    # Configurações do e-mail
    gmail_user = 'ficticio@gmail.com'
    gmail_password = 'senha'
    sent_from = gmail_user

    #variaveis de envio de e-mail
    to = ['destinatario@gmail.com']  #Para que será enviado o e-mail com os nomes do arquivo de exemplo
    subject = 'Nomes Limpos' #Assunto do e-mail
    body = 'Segue o arquivo com os nomes limpos.' #Corpo do e-mail

    # Monta o e-mail
    email_text = """\
    From: %s
    To: %s
    Subject: %s

    %s
    """ % (sent_from, ", ".join(to), subject, body) #Formata o e-mail

    # Envia o e-mail e verifica se foi enviado com sucesso
    try:
        # Conectando ao servidor de e-mail do Gmail
        server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        server.ehlo()
        server.login(gmail_user, gmail_password)
        server.sendmail(sent_from, to, email_text)
        server.close()

        print('Email enviado com sucesso!')

        #Apaga o arquivo com os nomes limpos
        os.remove("nomes_limpos.xlsx")

        #encerra
        exit()

    except smtplib.SMTPException as e:
        print('Erro ao enviar o e-mail:', e)


except FileNotFoundError:
    print(f"Arquivo '{caminho_arquivo}' não encontrado.")
except Exception as e:
    print(f"Erro ao ler o arquivo Excel: {e}")
