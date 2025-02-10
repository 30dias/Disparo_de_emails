# IMPORTANDO A BIBLIOTECA PANDAS PARA ANALISE E MANIPULAÇÃO DE PLANILHAS
import os
import smtplib
import pandas as pd
from time import sleep
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# LENDO AS PLANILHAS COM PANDAS
ws_resultados = pd.read_excel("Base Resultados de Exame.xlsx")
ws_prestadores = pd.read_csv("Base Prestadores.csv", encoding='unicode_escape', delimiter=';')

# FILTRANDO A "BASE RESULTADOS" UTILIZANDO SOMENTE AQUELAS SEM 'DATA RESULTADO EXAME'
sem_data = ws_resultados[ws_resultados['Data Res. Exame'].isna()]
codigos_prest = sem_data['Código Prestador'].fillna(0).astype(int) # PREENCHENDO VAZIAS COM 0

# CRIANDO COLUNAS PARA TABELA DE PRESTADORES
prestadores = {
    "Codigo":[],
    "Prestadora":[],
    "Email":[]
}
# NAVEGANDO TABELA DE CÓDIGOS DOS REGISTROS SEM DATA LINHA POR LINHA
for _, row in ws_prestadores.iterrows(): 
    if row['codigoPrestador'] in codigos_prest.values:
        cod = row['codigoPrestador'] 
        prestadora = row['nomePrestador']
        email = row['email']
    
    # ADICIONANDO AS INFORMAÇÕES CADASTRADAS DENTRO DA TABELA DE PRESTADORES
        prestadores['Codigo'].append(cod)
        prestadores['Prestadora'].append(prestadora)
        prestadores['Email'].append(email)

# CRIANDO COLUNAS PARA TABELA DE PESSOAS
pessoas = {
    "Codigo": [],
    "Funcionario": [],
    "CPF": [],
    "Tipo de Exame": [],
    "Exame": [],
    "Empresa": [],
    "Prestador": [],
    "Email": []
}
# NAVEGANDO NA PLANILHA FILTRADA COM OS REGISTROS SEM DATA LINHA POR LINHA
for _, resultado in sem_data.iterrows():
    # CASO O VALOR DA CELULA ESTEJA EM BRANCO CONVERTE PARA 'not_found' PARA TENTAR ACHAR ALGUM REGISTRO NA BASE
    if pd.isnull(resultado['Código Prestador']):
        resultado['Código Prestador'] = 'not_found'

    # SE O CÓDIGO DO REGISTRO ESTIVER SALVO NA TABELA DE PRESTADORES OU FOR 'not_found' PROSSEGUE COLETANDO DADOS
    if resultado['Código Prestador'] in prestadores['Codigo'] or 'not_found':
        nome = resultado['Nome Funcionário']
        cpf = resultado['CPF']
        tipo = resultado['Tipo de exame']
        exame = resultado['Nome do Exame']
        empresa = resultado['Nome da Empresa']
        cod = resultado['Código Prestador']
        nome_prestador = resultado['Nome do Prestador']

        cpf = str(cpf).format("00.000.000/0000-00")
    # TENTA PEGAR O CADASTRO DA PRESTADORA COM BASE NO CÓDIGO DO PRESTADOR
        try:
            index = prestadores['Codigo'].index(cod)
    # SE NÃO CONSEGUIR PELO CÓDIGO, TENTA PEGAR O CADASTRO DIRETAMENTE PELO NOME DO PRESTADOR
        except:
            try:
                index = prestadores['Prestadora'].index(nome_prestador)
            except:
            # CASO NÃO ENCONTRE POR CÓDIGO E NEM EMAIL ENTÃO O REGISTRO ESTÁ INCORRETO NA PLANILHA BASE E O CÓDIGO DEVE IGNORAR
                print(f'Registro inválido na base: {nome}, {cpf}, {empresa}')
                continue

    # UTILIZA O CADASTRO PEGO NA TENTATIVA ACIMA PARA PEGAR O NOME E EMAIL DA PRESTADORA
        try:
            nome_prestador = prestadores['Prestadora'][index]
            email_prestador = prestadores['Email'][index]
        except:
            print('ERRO!')

    # ADICIONANDO AS INFORMAÇÕES CADASTRADAS DENTRO DA TABELA DE PESSOAS
        pessoas['Codigo'].append(cod)
        pessoas['Funcionario'].append(nome)
        pessoas['CPF'].append(cpf)
        pessoas['Tipo de Exame'].append(tipo)
        pessoas['Exame'].append(exame)
        pessoas['Empresa'].append(empresa)
        pessoas['Prestador'].append(nome_prestador)
        pessoas['Email'].append(email_prestador)

# TRANSFORMANDO A TABELA DE PESSOAS EM UM DATAFRAME
base_df = pd.DataFrame(pessoas)
# SALVANDO AS INFORMAÇÕES EM UMA NOVA PLANILHA DO EXCEL CHAMADA 'Base'
base_df.to_excel("Base.xlsx",index=False)
print('Arquivo excel "Base.xlsx" gerado com sucesso!')
print('\nEnviando emails...')

# SOMENTE PARA TESTE, ARQUIVO EXCEL JÁ ESTÁ MODIFICADO
base_df = pd.read_excel('Base.xlsx')


# SALVANDO OS EMAILS DIFERENTES EM LISTA
emails = []
for email in base_df['Email']:
    if email not in emails:
        emails.append(email)

# SALVANDO AS PRESTADORAS DIFERENTES EM LISTA
prestadoras = []
for prest in base_df['Prestador']:
    if prest not in prestadoras:
        prestadoras.append(prest)

# APRESENTANDO INFORMAÇÕES SOBRE O LOGIN PARA DISPARO DOS EMAILS
if not os.path.exists('Login.txt'):
    print('\nPara que o envio dos emails seja bem sucedido é necessário usar um email da google! Essa versão demonstrativa do código atua com o servidor SMTP pré configurado nos servidores da Google, para enviar de um email de domínio privado é necessário possuir os dados do servidor e porta específica do email.\n')
    remetente = input('Digite um gmail para fazer o disparo: ')

    print('\nPara enviar os emails é necessário fazer login na conta. Por razões de privacidade e segurança a Google não permite acesso direto com os dados da conta, por isso, será necessário gerar uma senha de aplicativo para sua conta gmail. Um bloco de notas com o passo a passo está no diretório.\n')
    senha = input('Digite a senha de app da conta para fazer o disparo: ')

    # SALVANDO OS DADOS DE LOGIN EM UM BLOCO DE NOTAS NO MESMO DIRETÓRIO PARA OS PRÓXIMOS USOS
    with open('Login.txt','w',encoding='utf-8') as login:
        login.write(remetente + '\n')
        login.write(senha + '\n')
else:
    # CASO O LOGIN JÁ ESTIVER SALVO NO BLOCO DE NOTAS PERGUNTA SE DESEJA TROCAR DE CONTA
    resposta = input('Alterar conta?[S/N]: ')
    while True:
        if resposta in 'Nn': # SE NÃO QUISER TROCAR PEGA OS DADOS SALVOS E CONTINUA
            with open('Login.txt','r',encoding='utf-8') as login:
                remetente = login.readline().strip()
                senha = login.readline().strip()
            break
        elif resposta in 'Ss': # SE QUISER TROCAR SOLICITA NOVOS DADOS E SALVA
            remetente = input('Digite um gmail para fazer o disparo: ')
            senha = input('Digite a senha de app da conta para fazer o disparo: ')

            with open('Login.txt','w',encoding='utf-8') as login:
                login.write(remetente + '\n')
                login.write(senha + '\n')
            break
        else: # SE A RESPOSTA NÃO FOR S OU N PEDE NOVAMENTE
            print('Não entendi. Digite "S" para sim ou "N" para não!')


# CREDENCIAIS DE ACESSO AO SEU GMAIL
print(f'Logado em {remetente}\n')

# SERVIDOR E PORTA PADRÃO UTILIZADAS PARA O DISPARO DE EMAILS PELO GMAIL
smtpServer = 'smtp.gmail.com'  
porta = 587  

# PARA CADA PRESTADOR NA PLANILHA
for prest in prestadoras:
    nome_prestadora = prest # NOME DO PRESTADOR

    # PARA CADA EMAIL NA PLANILHA
    for email in emails:
        table_rows = "" # ZERA O HTML QUE RECEBE A TABELA DOS FUNCIONÁRIOS
        destinatario = email # COLOCA O EMAIL COMO DESTINATÁRIO DA MENSAGEM

        # PARA CADA LINHA DENTRO DA PLANILHA COMEÇA A BUSCA
        for _, row in base_df.iterrows():
            # SE O EMAIL E O NOME DO PRESTADOR FOREM IGUAIS AGRUPA OS FUNCIONÁRIOS
            if row['Email'] == email and row['Prestador'] == prest: 
                    # CRIA UMA NOVA LINHA NA TABELA HTML PARA RECEBER OS REGISTROS
                    table_rows += f"""
                    <tr>
                        <td style="border: 1px solid #ccc; padding: 8px;">{row['Funcionario']}</td>
                        <td style="border: 1px solid #ccc; padding: 8px;">{row['CPF']}</td>
                        <td style="border: 1px solid #ccc; padding: 8px;">{row['Tipo de Exame']}</td>
                        <td style="border: 1px solid #ccc; padding: 8px;">{row['Exame']}</td>
                        <td style="border: 1px solid #ccc; padding: 8px;">{row['Empresa']}</td>
                    </tr>
                    """

        # CORPO DA MENSAGEM (HTML BÁSICO FEITO SOMENTE PARA DEMONSTRAÇÃO)
        corpo = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>Documentos Pendentes</title>
        </head>
        <body style="font-family: Arial, sans-serif;">
            <div style="max-width: 600px; margin: auto; border: 1px solid #ccc; padding: 20px;">
                <div style="text-align: center;">
                    <img src="https://www.inmestra.com.br/logo.png" alt="Inmestra" style="max-width: 100%;">
                </div>
                <h2 style="text-align: center;">DOCUMENTOS PENDENTES</h2>
                
                <p>Olá <strong>{nome_prestadora}</strong>, tudo bem?</p>
                
                <p>Solicitamos o envio dos ASOS dos colaboradores listados abaixo:</p>
                
                <table style="width: 100%; border-collapse: collapse;">
                    <tr style="background-color: #f2f2f2;">
                        <th style="border: 1px solid #ccc; padding: 8px; text-align: left;">Funcionário</th>
                        <th style="border: 1px solid #ccc; padding: 8px; text-align: left;">CPF</th>
                        <th style="border: 1px solid #ccc; padding: 8px; text-align: left;">Tipo de exame</th>
                        <th style="border: 1px solid #ccc; padding: 8px; text-align: left;">Exame</th>
                        <th style="border: 1px solid #ccc; padding: 8px; text-align: left;">Empresa</th>
                    </tr>
                    {table_rows}
                </table>
                
                <p>Com a obrigatoriedade do <strong>eSocial</strong>, os documentos precisam ser enviados em tempo hábil para processamento e envio das informações para o Ministério do Trabalho, o não envio pode ocasionar <strong>multa</strong> para a empresa.</p>
                
                <p>Os documentos devem ser enviados para o e-mail: <a href="mailto:asos@inmestra.com.br">asos@inmestra.com.br</a></p>
                
                <p><strong>Caso o colaborador não tenha comparecido, favor informar para lançarmos falta nos exames não realizados.</strong></p>
                
                <p>Atenciosamente,<br>
                Equipe Inmestra.</p>
                
                <div style="text-align: center; margin-top: 20px; background-color: #003366; padding: 15px; color: white;">
                    <p><strong>(19) 3447-4700</strong></p>
                    <p>R. Gomes Carneiro, 1289 - Centro, Piracicaba - SP, 13400-530</p>
                    <p>
                        <a href="#" style="margin: 0 10px;"><img src="https://image.shutterstock.com/image-vector/linkedin-icon-260nw-1569257942.jpg" alt="LinkedIn" style="width: 24px; height: 24px;"></a>
                        <a href="#" style="margin: 0 10px;"><img src="https://image.shutterstock.com/image-vector/youtube-icon-260nw-1569257943.jpg" alt="YouTube" style="width: 24px; height: 24px;"></a>
                    </p>
                </div>
            </div>
        </body>
        </html>

        """
            
        # CRIANDO A MENSAGEM
        msg = MIMEMultipart()
        msg['From'] = remetente
        msg['To'] = destinatario
        msg['Subject'] = 'INMESTRA - Documentos Pendentes'

        # DEFININDO O CORPO HTML DA MENSAGEM
        msg.attach(MIMEText(corpo, "html"))

        # TENTANDO ENVIAR O EMAIL
        try:
            with smtplib.SMTP(smtpServer, porta) as servidor:
                servidor.starttls()  # Iniciar conexão segura
                servidor.login(remetente, senha)
                servidor.send_message(msg)
                
                print(f"Email enviado!")
                print(nome_prestadora)
                print(email+'\n')

                sleep(10)
        except Exception as e:
            print(f"Ocorreu um erro: {e}")