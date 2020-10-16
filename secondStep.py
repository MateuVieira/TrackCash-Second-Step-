# Este é o código de como ler emails do gmail
# discutido no calango, ele foi escrito para rodar em
# python 3
import os
import email
import imaplib

EMAIL = 'automacao@trackcash.com.br'
PASSWORD = 'mudar!@#'
SERVER = 'webmail.trackcash.com.br'
SUBJECT_INFO = 'Planilha de Repasse'


def read_email():
    # Abre uma conexão com o servidor do gmail
    # realiza o login e navega para a inbox
    mail = imaplib.IMAP4_SSL(SERVER)
    mail.login(EMAIL, PASSWORD)
    mail.select('inbox')

    # Realiza uma busca na inbox com o critério de busca ALL
    # para pegar todos os emails da imbox
    # a busca retorna o status da operação e uma lista com
    # os ids dos emails
    status, data = mail.search(None, f'SUBJECT "{SUBJECT_INFO}"')
    # data é uma lista com ids em blocos de bytes separados
    # por espaço neste formato: [b'1 2 3', b'4 5 6']
    # então para separar os ids nós primeiramente criamos
    # uma lista vazia
    mail_ids = []
    # e em seguida iteramos pelo data separando os blocos
    # de bytes e concatenando a lista resultante com nossa
    # lista inicial
    for block in data:
        # a função split chamada sem nenhum parâmetro
        # transforma texto ou bytes em listas usando como
        # ponto de divisão o espaço em branco:
        # b'1 2 3'.split() => [b'1', b'2', b'3']
        mail_ids += block.split()

    # agora para cada id de email baixaremos ele do gmail
    # e extrairemos o conteúdo
    for i in mail_ids:
        # a função fetch baixa o email do gmail passando id
        # e o padrão RFC que você deseja que a mensagem venha
        # formatada
        status, data = mail.fetch(i, '(RFC822)')

        # data por algum motivo que eu desconheç vem naquele
        # formato que um item é uma tupla e o outro é só b')'
        # por isso vamos iterar pelo data até encontrar a tupla
        for response_part in data:
            # se for a tupla a gente extrai o conteúdo
            if isinstance(response_part, tuple):
                # usando a função e extrair os dados do email
                # a gente pasa o segundo item da tupla que tem o
                # conteúdo porque o primeiro é só a informação
                # do formato do conteúdo
                message = email.message_from_bytes(response_part[1])

                # daí com o resultado conseguimos tirar as
                # informações de quem enviou o email e o assunto
                mail_from = message['from']
                mail_subject = message['subject']

                # agora para o conteúdo precisa de um pouco mais de
                # trabalho porque ele pode vir em texto puro
                # ou multipart, se for texto puro é só ir para o
                # else e extrair o conteúdo, senao tem que extrair
                # somente o que precisa
                if message.is_multipart():
                    mail_content = ''

                    # no caso do multipart vem junto com o email
                    # anexos e outras versões do mesmo email em
                    # diferentes formatos tipo texto imagem e html
                    # para isso vamos andar pelo payload do email
                    for part in message.get_payload():
                      fileName = part.get_filename()

                      if fileName != None:
                        filePath = os.path.join('./Downloads/', fileName)
                        
                        if not os.path.isfile(filePath):
                          fp = open(filePath, 'wb')
                          fp.write(part.get_payload(decode=True))
                          fp.close()

                        mail_content = 'Attached file download to the Downloads forlder'
                        mail_content += f'\n-> Filename: {fileName}'
                else:
                    mail_content = message.get_payload()

                # pro fim a gente printa os dados extraídos na tela
                print('From: {}'.format(mail_from))
                print('Subject: {}'.format(mail_subject))
                print('Content: {}'.format(mail_content))


# Iniciando script
read_email()