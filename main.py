import smtplib
import pandas
import pyautogui as pa
from time import sleep


link_arquivo = 'https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga'

# CONFIGURAÇÕES
pa.PAUSE = 0.5

# 1 - Entra no sistema da empresa
pa.press('win')
pa.write('chrome')
pa.press('enter')
pa.write(link_arquivo)
pa.press('enter')
# Click para clicar ou posso usar o moveTo, somente para mover o mouse
sleep(3)
pa.click(x=333, y=286, clicks=2)

# Fazendo download do arquivo
sleep(2)
pa.click(button='right')
sleep(2)
pa.click(x=462, y=643)

# Encerrando arquivo.
sleep(10)
pa.hotkey('alt', 'f4')

# Lê arquivo e pegar os indicadores
# OBS Coloquei os dados em arquivo somente para fica mais fácil, mas de correto o arquivo vai lê em downloads.

ARQUIVO = r''
tabela = pandas.read_excel(ARQUIVO)

FATURAMENTO = tabela['Valor Final'].sum()
QUANTIDADE = tabela['Quantidade'].sum()
DATA = tabela['Data'].copy()[0]


# Enviado o e-mail com os dados

EMAIL_QUE_ENCAMINHA = ''
SENHA_DO_EMAIL = ''

with smtplib.SMTP('smtp.mail.yahoo.com', port=587) as conexao:
    conexao.starttls()
    conexao.login(
        user=EMAIL_QUE_ENCAMINHA,
        password=SENHA_DO_EMAIL
    )
    conexao.sendmail(
        from_addr=EMAIL_QUE_ENCAMINHA,
        # to_addrs é o email para quem você deseja enviar.
        to_addrs='',
        msg=f'Subject:Relatorio {DATA}!\n\nForam vendidos {QUANTIDADE} faturamento do dia R$ {FATURAMENTO}'
    )
