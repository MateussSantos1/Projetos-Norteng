import os
import pandas as pd
from PyPDF2 import PdfWriter, PdfReader
import re
import unicodedata
from PyPDF2 import PdfWriter, PdfReader

arq_pdf = r"C:\Users\mateus.santos.AD\Pictures\CONTRACHEQUE E PONTO SGA\CQS.pdf"
pontos_pdf = r"C:\Users\mateus.santos.AD\Pictures\CONTRACHEQUE E PONTO SGA\PONTOS.pdf"
    
paginas = PdfReader(arq_pdf).pages
paginas2 = PdfReader(pontos_pdf).pages     


        # Agora apenas usando a verificação da expressão regular
for pagina in paginas:
            texto_pagina = pagina.extract_text()

            # Padrão regex para cinco dígitos, um espaço e um nome com até 7 palavras
            padraoo = r'[0-9]{6}\s+([A-ZÀ-ÿ]+(?:\s+[A-ZÀ-ÿ]+){0,3})'

            
            resultado_mascara = re.search(padraoo, texto_pagina)


            if resultado_mascara:
                nome_completo = resultado_mascara.group(1)  # Acessa o grupo capturado

                nome_completo_normalizado = unicodedata.normalize('NFKD', nome_completo).encode('ASCII', 'ignore').decode('ASCII')
                print(nome_completo_normalizado)
                print(nome_completo_normalizado[-1:])
                sla = nome_completo_normalizado.split(" ")

                print(sla)
                
                #REMOVER OS 3 ULTIMOS DIIGITOS
                if "/n" in sla[-1]:
                       nome_completo_normalizado = nome_completo_normalizado



                for pagina2 in paginas2:
                      texto_pagina2 = pagina2.extract_text()
                      texto_pagina2 = unicodedata.normalize('NFKD', texto_pagina2).encode('ASCII', 'ignore').decode('ASCII')
                      
                      if nome_completo_normalizado in texto_pagina2:
                            conteudo_pdf_novo = PdfWriter()
                            conteudo_pdf_novo.add_page(pagina)
                            conteudo_pdf_novo.add_page(pagina2)
                            nome_novo_pdf = f"{nome_completo_normalizado}.pdf"
                            with open(nome_novo_pdf, "wb") as novo_pdf:
                                conteudo_pdf_novo.write(novo_pdf)

                            print(f"Arquivo {nome_novo_pdf} gerado com sucesso!")
                            
                            break
                      
                if nome_completo_normalizado not in texto_pagina2:
                            print(f"{nome_completo_normalizado} não encontrado !!!")
                            


            else:
                print("Nenhum nome encontrado.")
        
            # print(texto_pagina)
