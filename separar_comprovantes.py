import os
import pandas as pd
import re
import unicodedata
from PyPDF2 import PdfWriter, PdfReader

# Caminho para o PDF e a planilha
arq_pdf = r"C:\Users\mateus.santos.AD\Documents\CONFERENCIACOMLIQUIDO\compjun24.pdf"
planilha = r"C:\Users\mateus.santos.AD\Documents\CONFERENCIACOMLIQUIDO\liquido.xlsx"

# Usuário escolhe a aba da planilha que irá trabalhar
user_abas = input("Qual aba você quer abrir? (Exemplo: '199'): ")
aba_escolhida = pd.read_excel(planilha, sheet_name=user_abas)
nomes = aba_escolhida["NOME"].tolist()

nomes_nao_encontrados = []

cont_nao_encontrados = 0

# Verificando cada nome da coluna de nomes
for nome in nomes:
    nome = unicodedata.normalize('NFKD', nome).encode('ASCII', 'ignore').decode('ASCII')
    partes_nome = nome.split()
    contador_nome_pessoa = sum(1 for parte in partes_nome if len(parte) > 3)
    encontrei_nome = False
    paginas = PdfReader(arq_pdf)

    for pagina in paginas.pages:
        texto_pagina = pagina.extract_text()
        texto_pagina_formatada = texto_pagina.replace(".", "").replace(",", "")

        # Verifica se o nome completo está na página
        if nome in texto_pagina_formatada:
            conteudo_pdf_novo = PdfWriter()
            conteudo_pdf_novo.add_page(pagina)
            nome_novo_pdf = f"{nome}.pdf"
            encontrei_nome = True

            with open(nome_novo_pdf, "wb") as novo_pdf:
                conteudo_pdf_novo.write(novo_pdf)

            print(f"Arquivo {nome_novo_pdf} gerado com sucesso!")
            
            break

        # Para nomes com quantidade de palavras maior que 4 // 2
        elif contador_nome_pessoa >= 2:
            linha_do_nome = aba_escolhida[aba_escolhida['NOME'] == nome]


            # COM ISSO, ESSE ERRO DE INDEX OF SERIA RESOLVIDO???

            if not linha_do_nome.empty:
                numero_linha = linha_do_nome.index[0]
                liquido = aba_escolhida.loc[numero_linha, 'LIQUIDO']

                e_sobrenome = ""
                palavras_encontradas = []
                for palavra in texto_pagina.split():
                    if palavra in partes_nome[-1] and len(palavra) > 3:
                        e_sobrenome = palavra

                    if len(palavra) > 3 and palavra in nome.split():
                        palavras_encontradas.append(palavra)

                        if len(palavras_encontradas) >= 2 and e_sobrenome in partes_nome[-1] and str(liquido) in texto_pagina_formatada:
                            conteudo_pdf_novo = PdfWriter()
                            conteudo_pdf_novo.add_page(pagina)
                            nome_novo_pdf = f"{nome}.pdf"
                            encontrei_nome = True

                            with open(nome_novo_pdf, "wb") as novo_pdf:
                                conteudo_pdf_novo.write(novo_pdf)

                            print(f"Arquivo {nome_novo_pdf} gerado com sucesso!")
                            
                            break

    if not encontrei_nome:
        nomes_nao_encontrados.append(nome)
        cont_nao_encontrados += 1
        print(f"{nome} não foi encontrado no PDF!!!")

# Contagem total de páginas no PDF
count_pdf_geral = len(paginas.pages)

print("--------------------------------------------------------------------------------------")
print(f"No total, existem {count_pdf_geral} comprovantes.")
print("--------------------------------------------------------------------------------------")
print("--------------------------------------------------------------------------------------")
print(f"Os seguintes nomes não foram encontrados: {nomes_nao_encontrados}")
print(f"No total, não foram encontrados os comprovantes de {cont_nao_encontrados} pessoas!")

# Criação de DataFrame com os nomes não encontrados
df_nomes_nao_encontrados = pd.DataFrame({'NOMES NÃO ENCONTRADOS': nomes_nao_encontrados})
diretorio = r'C:\Users\mateus.santos\Documents\CONFERENCIACOMLIQUIDO'
arquivo_saida = os.path.join(diretorio, '_Erro_Nomes_nao_encontrados_no_diretorio_.xlsx')
df_nomes_nao_encontrados.to_excel(arquivo_saida, index=False)
print(f'Foi gerado um arquivo Excel com os nomes não encontrados salvo em: {arquivo_saida}')
