from __future__ import print_function
import os.path
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from reportlab.lib.pagesizes import landscape, letter
from reportlab.pdfgen import canvas
from reportlab.platypus import Table, TableStyle
import pandas as pd
import datetime
from googleapiclient.http import MediaFileUpload

# If modifying these scopes, delete the file token.json.
scopes = ['https://www.googleapis.com/auth/spreadsheets',
          'https://www.googleapis.com/auth/drive']

creds = None
if os.path.exists('token.json'):
    creds = Credentials.from_authorized_user_file('token.json', scopes)
    # If there are no (valid) credentials available, let the user log in.
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file(
                'client_secret.json', scopes)
        creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
    with open('token.json', 'w') as token:
        token.write(creds.to_json())


def google_sheets():
    sheet_id = ''
    sheet_aba = 'dados'
    try:
        service = build('sheets', 'v4', credentials=creds)

        # Call the Sheets API
        sheet = service.spreadsheets()
        result = sheet.values().get(spreadsheetId=sheet_id,
                                    range=sheet_aba).execute()
        values = result.get('values', [])
        print(values)
        with open('extrato.csv', 'w', encoding='utf-8') as f:
            for row in values:
                f.write(';'.join(row) + '\n')

    except HttpError as err:
        print(err)


def google_drive(caminho_arquivo):
    folder_id = ''
    try:
        # create drive api client
        service = build('drive', 'v3', credentials=creds)

        file_metadata = {
            'name': os.path.basename(caminho_arquivo),
            'parents': [folder_id]
        }
        media = MediaFileUpload(caminho_arquivo,
                                mimetype='application/pdf', resumable=True)
        # pylint: disable=maybe-no-member
        file = service.files().create(body=file_metadata, media_body=media, fields='id').execute()
        print(F'File ID: "{file.get("id")}".')

    except HttpError as error:
        print(F'An error occurred: {error}')


def table_style():
    return TableStyle([('BACKGROUND', (0, 0), (-1, 0), '#CCCCCC'),
                       ('TEXTCOLOR', (0, 0), (-1, 0), '#000000'),
                       ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                       ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                       ('FONTSIZE', (0, 0), (-1, 0), 10),
                       ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                       ('BACKGROUND', (0, 1), (-1, -1), '#FFFFFF'),
                       ('GRID', (0, 0), (-1, -1), 1, '#000000')])


def dados_base():
    # Define a imagem a ser inserida
    imagem = 'capa.png'
    c.drawImage(imagem, 15, 530, width=760, height=70)
    # Define formatação dos campos texto
    c.setFont("Helvetica", 10)  # Define a fonte e o tamanho da fonte
    c.setFillColorRGB(0, 0, 0)
    # Define os campos de texto
    c.drawString(15, 500, "Nome: ")
    c.drawString(15, 480, "Cargo: ")
    c.drawString(15, 460, "Área: ")
    c.drawString(15, 440, "Grupo: ")
    c.drawString(50, 500, nome)
    c.drawString(50, 480, cargo)
    c.drawString(50, 460, area)
    c.drawString(50, 440, grupo)
    c.drawString(480, 500, "Matrícula: ")
    c.drawString(480, 480, "Admissão: ")
    c.drawString(480, 460, "CPF: ")
    c.drawString(480, 440, "Ciclo: ")
    c.drawString(535, 500, matricula)
    c.drawString(535, 480, admissao)
    c.drawString(535, 460, cpf)
    c.drawString(535, 440, ciclo)


def tabela_metas():
    # Estiliza a tabela
    estilo_tabela = table_style()
    # Define a largura das colunas
    col_widths = [130, 60, 90, 90, 90, 90, 90, 90]
    # Define a tabela das metas
    dados = [
        ["Descrição do Objetivo", "Peso", "70%\nMeta Mínima", "100%\nMeta Alvo", "120%\nMeta Superação", "Resultado",
         "Atingimento", "Atingimento\nPonderado"],
        ["Aumentar Nº de Cliente", "40%", "9.312.370", "13.303.385", "15.964.062", "12.334.836", "92,72%", "37,09%"],
        ["Aumentar Nº de Negócios", "60%", "14.438.562", "20.626.517", "24.751.820", "18.628.958", "90,32%", "54,19%"]]
    table1 = Table(dados, colWidths=col_widths)
    table1.setStyle(estilo_tabela)
    # Define a posição da tabela das metas na página
    table1.wrapOn(c, 400, 400)
    table1.drawOn(c, 15, 350)


def linha_apuracao():
    # Estiliza a tabela
    estilo_tabela = table_style()
    # Define a linha de apuração
    apuracao = [["Atingimento Ponderado para pagamento", "91.28% (1)"]]
    col_widths = [270, 90]  # Largura de cada coluna em pontos
    table2 = Table(apuracao, colWidths=col_widths)
    table2.setStyle(estilo_tabela)
    # Define a posição da linha de apuração na página
    table2.wrapOn(c, 400, 400)
    table2.drawOn(c, 385, 323)


def legenda():
    # Define formatação dos campos texto
    c.setFont("Helvetica", 9)  # Define a fonte e o tamanho da fonte
    c.setFillColorRGB(0, 0, 0)
    # Define a legenda
    c.drawString(15, 300, "(1) Atingimento ponderado entre os dois objetivos que será utilizado para cálculo;")
    c.drawString(15, 290, "(2) Salário nominal do mês de dezembro;")
    c.drawString(15, 280, "(3) Média de comissões recebidas durante o ano;")
    c.drawString(15, 270, "(4) Valor definido em Convenção Coletiva de acordo com sindicato de cada categoria;")
    c.drawString(15, 260, "(5) É a soma do salário nominal(2) e adicionais: média de comissões(3)"
                          " + anuênio/triênio(4);")
    c.drawString(15, 250, "(6) Múltiplo definido de acordo com cada grupo de cargo;")
    c.drawString(15, 240, "(7) Meses trabalhados durante o ano;")
    c.drawString(15, 230, "(8) Meses em que esteve afastado durante o ano, desconsiderados para cálculo;")
    c.drawString(15, 220, "(9) É o atingimento ponderado(1) x salário para cálculo(5) x múltiplo salarial(6)"
                          " x meses trabalhados")
    c.drawString(27, 210, "(7). Este valor está sujeito a tributação do Imposto de Renda.")


def base_calculo():
    # Define as coordenadas do retângulo
    x = 425
    y = 160
    width = 50
    height = 145
    # Define a cor do preenchimento do retângulo
    fill_color = "#CCCCCC"
    # Define a cor do texto
    text_color = "#000000"
    # Define a fonte e o tamanho da fonte
    c.setFont("Helvetica-Bold", 10)
    # Define a cor de preenchimento do retângulo
    c.setFillColor(fill_color)
    # Desenha o retângulo com a cor de preenchimento
    c.rect(x, y, width, height, fill=True, stroke=True)
    # Define a cor do texto
    c.setFillColor(text_color)
    # Define os textos
    texto1 = 'BASE'
    texto2 = 'DE'
    texto3 = 'CÁLCULO'
    # Calcula a posição vertical para centralizar cada texto
    text_height = c._leading
    text_y1 = y + (height - text_height) / 2 + 2 * text_height
    text_y2 = y + (height - text_height) / 2
    text_y3 = y + (height - text_height) / 2 - 2 * text_height
    # Desenha os textos centralizados dentro do retângulo
    c.drawCentredString(x + width / 2, text_y1, texto1)
    c.drawCentredString(x + width / 2, text_y2, texto2)
    c.drawCentredString(x + width / 2, text_y3, texto3)


def tabela_dados():
    # Estiliza a tabela
    estilo_tabela3 = TableStyle([('GRID', (0, 0), (-1, -1), 1, '#000000'),
                                 ('FONTNAME', (0, 0), (-1, 0), 'Helvetica'),
                                 ('FONTSIZE', (0, 0), (-1, 0), 10),
                                 ])
    col_widths = [180, 90]  # Largura de cada coluna em pontos
    dados = [
        ["(2) Salário Nominal", salario_nominal],
        ["(3) Média de Comissões", media_comissoes],
        ["(4) Anuênio / Triênio", anuenio_trienio],
        ["(5) Salário para Cálculo", salario_calculo],
        ["(6) Múltiplo Salarial", multiplo_salarial],
        ["(7) Meses Trabalhados", meses_trabalhados],
        ["(8) Meses de Afastamento", meses_afastamento],
        ["(9) Total Bruto", total_bruto],
    ]
    table3 = Table(dados, colWidths=col_widths)
    table3.setStyle(estilo_tabela3)
    # Define a posição da linha de apuração na página
    table3.wrapOn(c, 400, 400)
    table3.drawOn(c, 475, 160)


google_sheets()
path_entrada = 'extrato.csv'
df1 = pd.read_csv(path_entrada, encoding='utf-8', delimiter=';')
for i, row in df1.iterrows():
    if str(row['Múltiplo']) == '0,5':
        grupo = 'Equipe'
    else:
        grupo = 'Liderança / Especialistas / Consultores'
    today = datetime.date.today()
    ciclo = str(today.year - 1)
    nome = str(row['Nome'])
    cargo = str(row['Cargo'])
    area = str(row['Área'])
    matricula = str(row['Matrícula'])
    admissao = str(row['Admissão'])
    cpf = str(row['CPF'])
    salario_nominal = str(row['Salário'])
    media_comissoes = str(row['Média de Comissão ou Remuneração Variável'])
    anuenio_trienio = str(row['Anuênio/Triênio'])
    salario_calculo = str(row['Salário Cálculo'])
    multiplo_salarial = str(row['Múltiplo'])
    meses_trabalhados = str(row['Meses trabalhado'])
    meses_afastamento = str(row['Meses afastamento'])
    total_bruto = str(row['Valor bruto apurado'])
    # Cria um novo arquivo PDF
    c = canvas.Canvas("extrato_" + matricula + ".pdf", pagesize=landscape(letter))
    # Chama as funções
    dados_base()
    tabela_metas()
    linha_apuracao()
    legenda()
    base_calculo()
    tabela_dados()
    # Fecha o arquivo PDF
    c.save()
    caminho_arquivo = os.path.abspath("extrato_" + matricula + ".pdf")
    google_drive(caminho_arquivo)
