import pandas as pd
import openpyxl as xl
from PIL import Image, ImageDraw, ImageColor, ImageFont
from time import sleep
import os


def limpar_tela():
    if os.name == 'nt':
        os.system('cls')
    else:
        os.system('clear')
    


# Analise de dados
print('Carregando Dados...')
sleep(3)
df = pd.read_csv('res/dados.csv')

# Mensagem
print('Verificando qualidade dos dados...')
sleep(3)
print('Carregando amostra: 30 REGISTROS')
limpar_tela()
print('AMOSTRA -> 30 REGISTOS')
print(df.head(30))
sleep(7)


# Mensagem
limpar_tela()
print('Contagem de dados inconsistentes: \n')
sleep(3)
print(df.isna().sum())

# correção dos dados.
print('Iniciando correção dos dados...')
sleep(3)
data_emissao = df['data_emissao'].combine_first(df['data_fim'])
moda_carga_horia = df['carga_horaria'].mode()[0]
media_aproveitamento = df['aproveitamento'].mean()

# insere no dataframe -> Recomendações pandas 3
print('Inserindo informações corrigidas...\n')
sleep(5)
df['data_emissao'] = df['data_emissao'].fillna(data_emissao) # os dados mostravamos que a data_fim e a data_emissao eram as mesmas
df['carga_horaria'] = df['carga_horaria'].fillna(moda_carga_horia).astype(int).astype(str)
df['aproveitamento'] = df['aproveitamento'].fillna(media_aproveitamento).astype(int).astype(str)

# Salva o xlsx corrigido
print('Salvando arquivo corrigido (dados.xlsx)! ')
sleep(3)
df.to_excel('res/dados.xlsx', index=False)

limpar_tela()
print('Não há mais dados inconsistentes...')
print(df.isna().sum())
sleep(5)

# Automação
# Abrindo o arquivo
limpar_tela()
print('Iniciando módulo de Automação do Certificado.')
sleep(5)
arquivo = xl.load_workbook('res/dados.xlsx')
pagina = arquivo['Sheet1']

# loop para robo
for i, linha in enumerate(pagina.iter_rows(min_row=2)): # ignora o cabeçalho.
    nome_curso = linha[0].value
    nome_aluno = linha[1].value
    modalidade = linha[2].value
    data_inicio = linha[3].value
    data_fim = linha[4].value
    data_emissao = linha[5].value
    carga_horaria = linha[6].value
    aproveitamento = linha[7].value

# configura fontes
    fonte_do_nome = ImageFont.truetype('res/PinyonScript-Regular.ttf', 110)
    fonte_geral = ImageFont.truetype('res/tahoma.ttf', 35)

# carrega certificado e crie objeto
    certificado = Image.open('res/certificadomodelo.png')
    cor_do_nome = ImageColor.getrgb('#E4BD5A')
    insere_info = ImageDraw.Draw(certificado)

# insere informações no certificado.
    insere_info.text((780, 630), nome_aluno, fill=cor_do_nome, font=fonte_do_nome)
    insere_info.text((964, 758), nome_curso, fill='White', font=fonte_geral)
    insere_info.text((750, 813), modalidade, fill='White', font=fonte_geral)
    insere_info.text((900,868), aproveitamento, fill='White', font=fonte_geral)
    insere_info.text((519,950), carga_horaria, fill='White', font=fonte_geral)
    insere_info.text((730,950), data_inicio, fill='White', font=fonte_geral)
    insere_info.text((1032,950), data_fim, fill='White', font=fonte_geral)
    insere_info.text((1320,950), data_emissao, fill='White', font=fonte_geral)
    
# Salva certificado
    certificado.save(f'certificados/{i+1}_{nome_aluno}_certificado.png')

# bandeiras
    print(f'Certificado {i}, criado!')
