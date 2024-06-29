import openpyxl
from PIL import Image, ImageDraw, ImageFont

#abrir a dados da planilha 
alunos = openpyxl.load_workbook('AlunosDados.xlsx')
página_alunos = alunos['Página1']

for indice, linha in enumerate(página_alunos.iter_rows(min_row=2)):
    # Pegas os Nome, Tipo do curso, Carga horaria Inicio E Fim
    curso = linha[0].value
    nome = linha[1].value
    inicio = linha[2].value
    termino = linha[3].value
    carga_horaria = linha[4].value

    #escrever esses dados no certificado por escritos 
    fonte_nome = ImageFont.truetype('./Aston Script.ttf', 90)
    fonte_padrao= ImageFont.truetype('./GildaDisplay-Regular.ttf', 80)
    fonte_data= ImageFont.truetype('./GildaDisplay-Regular.ttf', 50)

    imagem = Image.open('./Certificado Padrão.png')
    desenhar = ImageDraw.Draw(imagem)
    
    desenhar.text((430,570), nome, fill='#333436', font=fonte_nome)
    desenhar.text((530,395), curso, fill='#333436', font=fonte_padrao)
    desenhar.text((208,925), str(inicio), fill='#333436', font=fonte_data)
    desenhar.text((208,1055), str(termino), fill='#333436', font=fonte_data)
    desenhar.text((1680,1130), str(carga_horaria), fill='black', font=fonte_data)

    imagem.save(f'./{indice}{nome}certificado.png')