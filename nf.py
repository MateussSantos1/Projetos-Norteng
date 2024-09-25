import os
import pandas as pd
import unicodedata
from PyPDF2 import PdfWriter, PdfReader

# Caminho para o PDF e a planilha
arq_pdf = r"C:\Users\mateus.santos.AD\Pictures\NOTA FISCAL AUTOM\nfexemp.pdf"


planilha = r"C:\Users\mateus.santos.AD\Pictures\NOTA FISCAL AUTOM\Sodexo.xlsx"

# Carregando a planilha
planilha_df = pd.read_excel(planilha)

codigos_nao_encontrados = []
cont_nao_encontrados = 0

codigos_doc = planilha_df['Número do Documento'].tolist()

# Procurando cada documento da tabela no PDF
for codigo in codigos_doc:
    # Removendo os 3 primeiros dígitos do documento


    # Selecionando a linha correspondente ao código
    print(f"Código é{codigo}")
    linha_selecionada = planilha_df.loc[planilha_df['Número do Documento'] == codigo]
    print(f"linha selecionada é {linha_selecionada}")

    codigo = codigo[3:]  # Atualizado para remover antes de procurar no PDF
    codigo_com_nf = f"nf={codigo}"


    
    if not linha_selecionada.empty:

        print('olá')
        # Obtendo o índice da primeira linha encontrada
        indice = linha_selecionada.index[0]
        
        referencia = unicodedata.normalize('NFKD', linha_selecionada['Histórico'].values[0]).encode('ASCII', 'ignore').decode('ASCII')
        obra = unicodedata.normalize('NFKD', linha_selecionada['Centro de Custo'].values[0]).encode('ASCII', 'ignore').decode('ASCII')
        valor = linha_selecionada['Valor líquido'].values[0]

        paginas = PdfReader(arq_pdf).pages
        encontrei_codigo = False

        for pagina in paginas:
            texto_pagina = pagina.extract_text()

            if codigo_com_nf in texto_pagina:
                conteudo_pdf_novo = PdfWriter()
                conteudo_pdf_novo.add_page(pagina)
                nome_novo_pdf = f"{referencia} -- {obra} -- {valor}.pdf"

                with open(nome_novo_pdf, "wb") as novo_pdf:
                    conteudo_pdf_novo.write(novo_pdf)

                print(f"Arquivo {codigo} gerado com sucesso!")
                encontrei_codigo = True
                break

        if not encontrei_codigo:
            codigos_nao_encontrados.append(codigo)
            cont_nao_encontrados += 1
            print(f"{codigo} não foi encontrado no PDF!!!")




# Relatório de não encontrados
print(f"Os seguintes documentos não foram encontrados: {codigos_nao_encontrados}")
print(f"No total, não foram encontrados os comprovantes de {cont_nao_encontrados} notas")

# Criação de DataFrame com os códigos não encontrados
df_codigos_nao_encontrados = pd.DataFrame({'CODIGOS NÃO ENCONTRADOS': codigos_nao_encontrados})

diretorio = r'C:\Users\mateus.santos.AD\Pictures\NOTA FISCAL AUTOM'
arquivo_saida = os.path.join(diretorio, 'Codigos_nao_encontrados.xlsx')
df_codigos_nao_encontrados.to_excel(arquivo_saida, index=False)
print(f'Foi gerado um arquivo Excel com os códigos não encontrados salvo em: {arquivo_saida}')

