import os
import pandas as pd
import unicodedata
from PyPDF2 import PdfWriter, PdfReader
import re
import Levenshtein

# Caminho para o PDF e a planilha
arq_pdf = r"E:\teste_sodexo\salsichao.pdf"
planilha = r"E:\teste_sodexo\sodexo.xlsx"

# Carregando a planilha
planilha_df = pd.read_excel(planilha)
cont_encontrados = 0
codigos_nao_encontrados = []
cont_nao_encontrados = 0

codigos_doc = planilha_df['Número do Documento'].tolist()

# Procurando cada documento da tabela no PDF
for codigo in codigos_doc:
    # Selecionando a linha correspondente ao código
    linha_selecionada = planilha_df.loc[planilha_df['Número do Documento'] == codigo]
    
    # Convertendo o código para string antes de remover os 3 primeiros dígitos
    codigo = str(codigo)[3:]
    
    if not linha_selecionada.empty:
        # Obtendo o índice da primeira linha encontrada
        indice = linha_selecionada.index[0]
        
        referencia = unicodedata.normalize('NFKD', linha_selecionada['Histórico'].values[0]).encode('ASCII', 'ignore').decode('ASCII')
        referencia = referencia.replace('*', ' ')
        referencia = referencia.replace('/', ' ')

        obra = unicodedata.normalize('NFKD', linha_selecionada['Centro de Custo'].values[0]).encode('ASCII', 'ignore').decode('ASCII')
        valor = linha_selecionada['Valor líquido'].values[0]
        
        paginas = PdfReader(arq_pdf).pages
        encontrei_codigo = False
        
        # Agora apenas usando a verificação da expressão regular
        for pagina in paginas:
            texto_pagina = pagina.extract_text()
            padrao = r"Nro\sFedido\.\:\s+([A-Za-z0-9]{8})"
            
            resultado_mascara = re.search(padrao, texto_pagina)
            margem_de_erro = 1
            
            if resultado_mascara:
                numero_encontrado = resultado_mascara.group(1)  # Corrigido para group(1)
                
                # Ver quantos caracteres diferem entre si
                comparacao = Levenshtein.distance(numero_encontrado, codigo)
                print(str(valor))
                valor_formatado = "{:,.2f}".format(valor).replace(",", "X").replace(".", ",").replace("X", ".") 
                if comparacao <= margem_de_erro and codigo[:1] in numero_encontrado[:1] and str(valor_formatado) in texto_pagina:
                    conteudo_pdf_novo = PdfWriter()
                    conteudo_pdf_novo.add_page(pagina)
                    nome_novo_pdf = f"{referencia} -- {obra} -- {valor}.pdf"

                    with open(nome_novo_pdf, "wb") as novo_pdf:
                        conteudo_pdf_novo.write(novo_pdf)

                    print(f"Arquivo {codigo} gerado com sucesso! ----- {referencia} -- {obra} -- {valor}")
                    encontrei_codigo = True
                    cont_encontrados += 1
                    break
        
        if not encontrei_codigo:
            codigos_nao_encontrados.append(codigo)
            cont_nao_encontrados += 1
            print(f"{codigo} não foi encontrado no PDF!!!")

print(f"Foram encontrados com sucesso os comprovantes de {cont_encontrados} notas!")

# Relatório de não encontrados
print(f"Os seguintes documentos não foram encontrados: {codigos_nao_encontrados}")
print(f"No total, não foram encontrados os comprovantes de {cont_nao_encontrados} notas")

# Criação de DataFrame com os códigos não encontrados
df_codigos_nao_encontrados = pd.DataFrame({'CODIGOS NÃO ENCONTRADOS': codigos_nao_encontrados})

diretorio = r'E:\teste_sodexo'
arquivo_saida = os.path.join(diretorio, 'Codigos_nao_encontrados.xlsx')
df_codigos_nao_encontrados.to_excel(arquivo_saida, index=False)
print(f'Foi gerado um arquivo Excel com os códigos não encontrados salvo em: {arquivo_saida}')
