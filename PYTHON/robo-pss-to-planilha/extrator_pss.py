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
                # Usar layout=True pode ajudar a preservar a estrutura da tabela
                texto_pagina = pagina.extract_text(x_tolerance=2, layout=True)
                if texto_pagina:
                    texto_completo += texto_pagina + "\n"

        nome_arquivo = os.path.basename(caminho_arquivo)

        pss_match = re.search(r'(?:PSS|PROCESSO SELETIVO SIMPLIFICADO)\s*N?º?\s*(\d{2,3}/\d{4})', texto_completo, re.IGNORECASE)
        pss = pss_match.group(1) if pss_match else "Não encontrado"
        ano = pss.split('/')[1] if pss != "Não encontrado" else "Não encontrado"
        if ano == "Não encontrado":
            ano_match_fn = re.search(r'(\b20\d{2}\b)', nome_arquivo)
            if ano_match_fn: ano = ano_match_fn.group(1)
        
        registros_do_arquivo = []
        if "QUADRO DE VAGAS" not in texto_completo.upper():
            print(f"    -> AVISO: 'QUADRO DE VAGAS' não encontrado em {nome_arquivo}.")
            return []

        partes = re.split(r'QUADRO DE VAGAS', texto_completo, flags=re.IGNORECASE)
        texto_da_tabela = partes[-1]
        linhas_da_tabela = texto_da_tabela.strip().split('\n')

        idx_municipio, idx_disciplina = -1, -1
        data_start_index = 0
        header_found = False

        for i, linha in enumerate(linhas_da_tabela[:6]):
            linha_upper = linha.strip().upper()
            colunas_header = [col.strip() for col in re.split(r'\s{2,}', linha_upper)]
            
            temp_idx_mun, temp_idx_disc = -1, -1
            for j, coluna in enumerate(colunas_header):
                if any(kw in coluna for kw in ["MUNICÍPIO", "LOCALIDADE"]):
                    temp_idx_mun = j
                if any(kw in coluna for kw in ["DISCIPLINA", "FUNÇÃO", "CARGO"]):
                    temp_idx_disc = j

            if temp_idx_mun != -1 and temp_idx_disc != -1:
                idx_municipio, idx_disciplina = temp_idx_mun, temp_idx_disc
                data_start_index = i + 1
                header_found = True
                break
        
        if not header_found:
            print(f"    -> AVISO: Cabeçalho não identificado em {nome_arquivo}. Tentando extração por posição padrão.")

            idx_municipio, idx_disciplina = 0, 1 

        for linha in linhas_da_tabela[data_start_index:]:
            linha_strip = linha.strip()

            if not linha_strip or "TOTAL DE VAGAS" in linha_strip.upper():
                continue

            colunas = re.split(r'\s{2,}', linha_strip)
            
            if len(colunas) > max(idx_municipio, idx_disciplina):
                municipio_raw = colunas[idx_municipio].strip()
                disciplina_raw = colunas[idx_disciplina].strip()

                if len(municipio_raw) < 3 or len(disciplina_raw) < 3 or \
                   municipio_raw.upper() in ["MUNICÍPIO", "LOCALIDADE"] or \
                   disciplina_raw.upper() in ["DISCIPLINA", "FUNÇÃO"]:
                    continue

                municipios_individuais = [m.strip() for m in re.split(r',| E ', municipio_raw) if len(m.strip()) > 2]
                
                for mun in municipios_individuais:
                    registros_do_arquivo.append({
                        "Ano": ano, "PSS": pss, "Edital": nome_arquivo,
                        "Município": mun, "Disciplina": disciplina_raw,
                        "Arquivo_Origem": nome_arquivo, "Data_Extração": datetime.now().strftime('%Y-%m-%d')
                    })
        
        if not registros_do_arquivo:
            print(f"    -> FALHA: Nenhum registro de vaga válido foi extraído de {nome_arquivo}.")
        
        return registros_do_arquivo

    except Exception as e:
        print(f"    -> ERRO GERAL ao processar {os.path.basename(caminho_arquivo)}: {e}")
        return []

def processar_pasta_e_salvar_planilha(caminho_da_pasta, nome_arquivo_saida):
    registros = []
    
    print("\n--- INICIANDO EXTRAÇÃO ---")
    if not os.path.isdir(caminho_da_pasta):
        print(f"ERRO: A pasta '{caminho_da_pasta}' não foi encontrada!")
        return

    arquivos_na_pasta = [f for f in os.listdir(caminho_da_pasta) if f.lower().endswith('.pdf')]
    print(f"Encontrados {len(arquivos_na_pasta)} arquivos PDF em '{caminho_da_pasta}'.")

    for nome_arquivo in sorted(arquivos_na_pasta):
        caminho_completo = os.path.join(caminho_da_pasta, nome_arquivo)
        print(f" -> Processando: {nome_arquivo}...")
        
        dados_extraidos = extrair_informacoes_pdf(caminho_completo)
        
        if dados_extraidos:
            registros.extend(dados_extraidos)
            print(f"    -> Sucesso! ({len(dados_extraidos)} registros extraídos)")

    if not registros:
        print("\n--- EXTRAÇÃO FINALIZADA ---")
        print("Nenhum registro foi extraído. A planilha não foi gerada.")
        return

    df = pd.DataFrame(registros)
    df = df[["Ano", "PSS", "Edital", "Município", "Disciplina", "Arquivo_Origem", "Data_Extração"]]
    df.to_csv(nome_arquivo_saida, index=False, sep=';', encoding='utf-8-sig')

    print("\n--- EXTRAÇÃO FINALIZADA ---")
    print(f"Total de arquivos PDF processados: {len(arquivos_na_pasta)}")
    print(f"Total de registros de vagas salvos na planilha: {len(registros)}")
    print(f"Dados salvos em '{nome_arquivo_saida}'!")

if __name__ == "__main__":
    sys.stdout = Logger("log_extracao_final_sem_api.txt")
    print(f"--- LOG DE EXECUÇÃO - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ---")
    
    pasta_convocacoes_especiais = "C:/Users/SEDUC/Desktop/conv especiais"
    arquivo_csv_saida = "convocacoes_especiais.csv"
    
    try:
        processar_pasta_e_salvar_planilha(pasta_convocacoes_especiais, arquivo_csv_saida)
    except Exception as e:
        print(f"\n OCORREU UM ERRO INESPERADO QUE INTERROMPEU O SCRIPT ")
        print(f"ERRO: {e}")
    finally:
        print(f"\n--- FIM DO LOG ---")