import os
import re
import sys
import pandas as pd
import pdfplumber
from datetime import datetime
from urllib.parse import unquote

class Logger(object):
    def __init__(self, filename="log_extracao.txt"):
        self.terminal = sys.stdout
        self.log = open(filename, "w", encoding='utf-8')

    def write(self, message):
        self.terminal.write(message)
        self.log.write(message)

    def flush(self):
        pass

def extrair_informacoes_pdf(caminho_arquivo):
    try:
        texto_completo = ""
        with pdfplumber.open(caminho_arquivo) as pdf:
            for pagina in pdf.pages:
                texto_pagina = pagina.extract_text(x_tolerance=2, use_text_flow=True)
                if texto_pagina:
                    texto_completo += texto_pagina + "\n"
        
        nome_arquivo_upper = os.path.basename(caminho_arquivo).upper()

        pss_match = re.search(r'(?:PSS|PROCESSO SELETIVO SIMPLIFICADO)\s*N?º?\s*(\d{2,3}/\d{4})', texto_completo, re.IGNORECASE)
        pss = pss_match.group(1) if pss_match else "Não encontrado"
        ano = pss.split('/')[1] if pss != "Não encontrado" else "Não encontrado"

        edital = os.path.basename(caminho_arquivo)

        municipio_match = re.search(r'MUNICÍPIO:\s*([^\s\n]+)', texto_completo, re.IGNORECASE)
        municipio = municipio_match.group(1).strip() if municipio_match else "Não encontrado"

        disciplina = "Não encontrado"
        padroes_disciplina = [r'CARGO:\s*(.*?)(?:\n|N°)', r'FUNÇÃO:\s*(.*?)(?:\n|N°)']
        for padrao in padroes_disciplina:
            disciplina_match = re.search(padrao, texto_completo, re.IGNORECASE)
            if disciplina_match:
                disciplina = ' '.join(disciplina_match.group(1).strip().split())
                break

        if "QUADRO DE VAGAS" in texto_completo.upper():
            try:
                partes = re.split(r'QUADRO DE VAGAS', texto_completo, flags=re.IGNORECASE)
                if len(partes) > 1:
                    texto_da_tabela = partes[1]
                    linhas_da_tabela = texto_da_tabela.strip().split('\n')
                    
                    header_keywords = ['COLUNA', 'MUNICÍPIO', 'DISCIPLINA', 'LOCALIDADE', 'CLASSIFICAÇÃO', 'VAGAS']
                    junk_values = ['NÚMERO', 'NÚMERO DE', 'Nº DE', 'NOME', 'N DE', 'DE']

                    for linha in linhas_da_tabela:
                        linha_strip = linha.strip()
                        if not linha_strip or any(keyword in linha_strip.upper() for keyword in header_keywords):
                            continue
                        
                        primeira_linha_dados = linha_strip
                        colunas = re.split(r'\s{2,}', primeira_linha_dados)
                        
                        if municipio == "Não encontrado" and len(colunas) > 0:
                            valor_municipio = colunas[0].strip()
                            if valor_municipio.upper() not in header_keywords and valor_municipio.upper() not in junk_values:
                                municipio = valor_municipio
                        
                        if disciplina == "Não encontrado" and len(colunas) > 1:
                            valor_disciplina = colunas[1].strip()
                            if valor_disciplina.upper() not in header_keywords and valor_disciplina.upper() not in junk_values:
                                disciplina = valor_disciplina
                        
                        break 
            except Exception as e:
                print(f"    -> Aviso: Tentativa de ler tabela falhou. Erro: {e}")

        if ano == "Não encontrado":
            ano_match_fn = re.search(r'(\b20\d{2}\b)', nome_arquivo_upper)
            if ano_match_fn: ano = ano_match_fn.group(1)

        if pss == "Não encontrado":
            pss_match_fn = re.search(r'PSS\s*(\d+[-_]\d{4})', nome_arquivo_upper)
            if pss_match_fn: pss = pss_match_fn.group(1).replace('-', '/').replace('_', '/')
        
        if disciplina == "Não encontrado":
            cargos = ['PROFESSOR', 'MERENDEIRA', 'ANALISTA', 'ASSISTENTE', 'PSICOLOGO', 'ENGENHEIRO', 'ARQUITETO', 'CONTADOR', 'TRADUTOR', 'EDUCADOR', 'SECTET', 'ADMINISTRADOR', 'ACOMPANHANTE ESPECIALIZADO', 'INTERPRETE DE LIBRAS', 'CEMEP']
            for cargo in cargos:
                if cargo.replace(' ', '') in nome_arquivo_upper.replace(' ', ''):
                    disciplina = cargo.title()
                    break
        
        if pss == "Não encontrado":
            print(f"AVISO: PSS não encontrado.")
            return None

        return {"Ano": ano, "PSS": pss, "Edital": edital, "Município": municipio, "Disciplina": disciplina, "Arquivo_Origem": os.path.basename(caminho_arquivo), "Data_Extração": datetime.now().strftime('%Y-%m-%d')}

    except Exception as e:
        print(f"ERRO CRÍTICO ao processar o arquivo: {e}")
        return None

def processar_pasta_e_salvar_planilha(pasta_principal, nome_arquivo_saida):
    registros = []
    arquivos_encontrados = 0
    arquivos_ignorados = 0
    palavras_a_ignorar = ['RETIFICACAO', 'RESULTADO', 'REVOGACAO', 'ADENDO', 'NOTA DE ESCLARECIMENTO', 'ERRATA']

    print("--- INICIANDO EXTRAÇÃO DE CONVOCAÇÕES ESPECIAIS ---")
    
    if not os.path.isdir(pasta_principal):
        print(f"ERRO: A pasta principal não foi encontrada: '{pasta_principal}'")
        return

    for raiz, _, arquivos in os.walk(pasta_principal):
        if os.path.basename(raiz).upper() == 'CONVOCACOES_ESPECIAIS':
            print(f"\nAnalisando pasta: {raiz}")
            for nome_arquivo in arquivos:
                if nome_arquivo.lower().endswith('.pdf'):
                    arquivos_encontrados += 1
                    nome_limpo = unquote(nome_arquivo)
                    
                    if any(palavra in nome_limpo.upper() for palavra in palavras_a_ignorar):
                        print(f" -> Ignorando (documento auxiliar): {nome_limpo}")
                        arquivos_ignorados += 1
                        continue
                    
                    if 'CONVOCACAOESPECIAL' not in nome_limpo.upper().replace(' ', ''):
                        print(f" -> Ignorando (não é Convocação Especial): {nome_limpo}")
                        arquivos_ignorados += 1
                        continue

                    caminho_completo = os.path.join(raiz, nome_arquivo)
                    print(f" -> Processando Convocação Especial: {nome_limpo}...")
                    dados = extrair_informacoes_pdf(caminho_completo)
                    if dados:
                        registros.append(dados)
                        print("    -> Sucesso!")
                    else:
                        print("    -> Falha na extração.")
    
    df = pd.DataFrame(registros)
    df.to_csv(nome_arquivo_saida, index=False, sep=';', encoding='utf-8-sig')
    
    print("\n--- EXTRAÇÃO FINALIZADA ---")
    print(f"Total de arquivos PDF encontrados na pasta: {arquivos_encontrados}")
    print(f"Arquivos ignorados (convocações regulares, resultados, etc.): {arquivos_ignorados}")
    print(f"Convocações ESPECIAIS processadas com sucesso: {len(registros)}")
    print(f"Dados salvos em '{nome_arquivo_saida}'!")

if __name__ == "__main__":
    sys.stdout = Logger("log_extracao.txt")
    
    print(f"--- LOG DE EXECUÇÃO - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ---")
    
    pasta_raiz_pss = "C:/Users/SEDUC/Desktop/Downloads PSS"
    arquivo_csv_saida = "convocacoes_especiais.csv"
    
    try:
        processar_pasta_e_salvar_planilha(pasta_raiz_pss, arquivo_csv_saida)
    except Exception as e:
        print(f"\n!!! OCORREU UM ERRO INESPERADO QUE INTERROMPEU O SCRIPT !!!")
        print(f"ERRO: {e}")
    
    print(f"\n--- FIM DO LOG ---")