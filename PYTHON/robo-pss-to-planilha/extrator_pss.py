import os
import re
import sys
import pandas as pd
import pdfplumber
from datetime import datetime

class Logger(object):
    def __init__(self, filename="log_extracao.txt"):
        self.terminal = sys.stdout
        self.log = open(filename, "w", encoding="utf-8")

    def write(self, message):
        self.terminal.write(message)
        self.log.write(message)

    def flush(self):
        pass

def clean_text(text):
    text = re.sub(r'\s+\d+$', '', text).strip()
    text = re.sub(r'\s+', ' ', text).replace('\n', ' ').strip()
    return text

def extract_info_from_line(line, keywords):
    for kw in keywords:
        if kw in line.upper():
            return clean_text(line.split(kw)[-1].strip())
    return None

def padrao_lista_disciplina_unica(texto_completo, **kwargs):
    print("      -> Padrão detectado: Lista de Disciplina Única")
    registros = []
    
    disciplina = "Não Identificada"
    disciplina_match = re.search(r'(?:FUNÇÃO|CARGO|DISCIPLINA)\s*:\s*([\s\S]+?)(?=\n\n|MUNICÍPIO|QUADRO DE VAGAS|RELAÇÃO DE E-MAILS|RELAÇÃO DE E-MAIL|TOTAL DE VAGAS|ASSINATURA)', texto_completo, re.IGNORECASE)
    if disciplina_match:
        disciplina = clean_text(disciplina_match.group(1))
    
    if disciplina == "Não Identificada" or not disciplina:
        if "ATENDIMENTO EDUCACIONAL ESPECIALIZADO" in texto_completo.upper() or "AEE" in texto_completo.upper():
            disciplina = "Atendimento Educacional Especializado (AEE)"
        elif "SALA DE RECURSO MULTIFUNCIONAL" in texto_completo.upper():
            disciplina = "Professor/Sala de Recurso Multifuncional"
        elif "PROJETO MUNDI" in texto_completo.upper():
            disciplina = "Professor para o Projeto Mundi"
        elif "INTÉRPRETE DE LIBRAS" in texto_completo.upper():
            disciplina = "Intérprete de Libras"
        elif "ACOMPANHAMENTO ESPECIALIZADO" in texto_completo.upper():
            disciplina = "Professor de Acompanhamento Especializado"
        elif "ENSINO REGULAR" in texto_completo.upper():
            disciplina = "Ensino Regular"
        elif "SOME" in texto_completo.upper():
            disciplina = "Sistema de Organização Modular de Ensino (SOME)"
        elif "CONTADOR" in texto_completo.upper():
            disciplina = "Contador"
        elif "DOCENTE" in texto_completo.upper():
            disciplina = "Docente"

    partes = re.split(r'QUADRO DE VAGAS|MUNICÍPIOS DE LOTAÇÃO|MUNICÍPIO DE LOTAÇÃO', texto_completo, flags=re.IGNORECASE)
    if len(partes) < 2: return []
    
    texto_lista = partes[-1]
    
    municipios = []
    for linha in texto_lista.strip().split('\n'):
        linha_upper = linha.upper()
        if any(kw in linha_upper for kw in ["TOTAL", "VAGAS", "PÁGINA", "EDITAL", "CARGO", "FUNÇÃO", "ASSINATURA", "DISCIPLINA", "COLUNA"]):
            continue
        
        mun_clean = clean_text(linha)
        if len(mun_clean) > 3:
            municipios.append(mun_clean)
            
    for mun in municipios:
        if disciplina != "Não Identificada":
            registros.append({"Município": mun, "Disciplina": disciplina})
            
    return registros

def padrao_reforco_escolar(texto_completo, **kwargs):
    print("      -> Padrão detectado: Projeto Reforço Escolar")
    registros = []
    
    texto_inicial = texto_completo[:2000].lower()
    disciplinas = []
    
    if re.search(r'língua\s*portuguesa\s*(e|ou)?\s*matemática', texto_inicial):
        disciplinas = ["Língua Portuguesa", "Matemática"]
    elif re.search(r'pedagogia\s*(e|ou)?\s*educação\s*física', texto_inicial):
        disciplinas = ["Pedagogia", "Educação Física"]
    elif "língua portuguesa" in texto_inicial:
        disciplinas.append("Língua Portuguesa")
    elif "matemática" in texto_inicial:
        disciplinas.append("Matemática")
    elif "pedagogia" in texto_inicial:
        disciplinas.append("Pedagogia")
    elif "educação física" in texto_inicial:
        disciplinas.append("Educação Física")

    if not disciplinas:
        print("      -> AVISO: Não foi possível extrair as disciplinas (LP/MAT ou PED/EF) do texto inicial.")
        return []
    
    partes = re.split(r'QUADRO DE VAGAS|MUNICÍPIO DE LOTAÇÃO', texto_completo, flags=re.IGNORECASE)
    if len(partes) < 2: return []
        
    texto_lista = partes[-1]
    municipios = []
    
    linhas_municipios = texto_lista.strip().split('\n')
    for linha in linhas_municipios:
        linha_upper = linha.upper()
        if any(kw in linha_upper for kw in ["TOTAL", "VAGAS", "PÁGINA", "EDITAL", "CARGO", "FUNÇÃO", "ASSINATURA", "DISCIPLINA", "COLUNA"]):
            continue
        
        mun_clean = clean_text(linha)
        if len(mun_clean) > 3:
            municipios.append(mun_clean)

    for mun in municipios:
        if len(mun.strip()) > 2:
            for disc in disciplinas:
                registros.append({"Município": mun.strip(), "Disciplina": disc})
    return registros

def padrao_projeto_sei_refinado(texto_completo, **kwargs):
    print("      -> Padrão detectado: Projeto SEI (Refinado)")
    registros = []
    
    if 'QUADRO DE VAGAS' not in texto_completo.upper():
        print("      -> AVISO: Documento do Projeto SEI sem quadro de vagas.")
        kwargs['status_flags']['sem_vagas'] = True 
        return []

    partes_fim = re.split(r'RELAÇÃO DE E-MAILS|RELAÇÃO DE E-MAIL|ANEXO I', texto_completo, flags=re.IGNORECASE)
    texto_relevante = partes_fim[0]

    partes_inicio = re.split(r'QUADRO DE VAGAS', texto_relevante, flags=re.IGNORECASE)
    texto_tabela = partes_inicio[-1] if len(partes_inicio) > 1 else ""

    disciplina = "Docente Projeto SEI"
    ure_atual = "Não Identificada"
    
    for linha in texto_tabela.strip().split('\n'):
        linha_upper = linha.upper()
        if "TOTAL DE VAGAS" in linha_upper or "TOTAL GERAL" in linha_upper or "ASSINATURA" in linha_upper: break
        
        ure_match = re.search(r'(\d+ª?\s*URE\s*[-\s]*[A-ZÇÃ-ÕÁ-Ú]+)', linha, re.IGNORECASE)
        if ure_match:
            ure_atual = ure_match.group(1).strip()
        else:
            localidade = clean_text(re.sub(r'\s+\d+(?:\s*Manhã/Tarde/Noite)?$', '' , linha))
            if len(localidade) > 3 and "VAGAS" not in localidade.upper() and "MUNICÍPIO" not in localidade.upper() and "DISCIPLINA" not in localidade.upper():
                if ure_atual != "Não Identificada":
                    registros.append({"Município": f"{localidade} ({ure_atual})", "Disciplina": disciplina})
                else:
                    registros.append({"Município": localidade, "Disciplina": disciplina})
    return registros

def padrao_tabela_generica(texto_completo, **kwargs):
    print("      -> Padrão detectado: Tabela Genérica Flexível")
    partes = re.split(r'QUADRO DE VAGAS', texto_completo, flags=re.IGNORECASE)
    if len(partes) < 2: return []
    
    texto_tabela = partes[-1]
    linhas = texto_tabela.strip().split('\n')
    registros = []

    header_idx = -1
    keywords_municipio = ['MUNICÍPIO', 'LOCALIDADE', 'MUNICÍPIOS/LOCALIDADES']
    keywords_disciplina = ['DISCIPLINA', 'FUNÇÃO', 'CARGO', 'COMPONENTE CURRICULAR', 'ÁREA']
    
    for i, linha in enumerate(linhas):
        linha_upper = linha.upper()
        tem_municipio = any(kw in linha_upper for kw in keywords_municipio)
        tem_disciplina = any(kw in linha_upper for kw in keywords_disciplina)
        if tem_municipio and tem_disciplina:
            header_idx = i
            break
            
    if header_idx != -1:
        header_line = linhas[header_idx]
        col_positions = []
        
        for key, kws in [("municipio", keywords_municipio), ("disciplina", keywords_disciplina)]:
            best_match_start = -1
            for kw in kws:
                match = re.search(r'\b' + kw + r'\b', header_line, re.IGNORECASE)
                if match:
                    if best_match_start == -1 or match.start() < best_match_start:
                        best_match_start = match.start()
            if best_match_start != -1:
                col_positions.append((best_match_start, key))
        
        col_positions.sort()
        
        if len(col_positions) >= 2:
            col_slices = {}
            for i, (start, key) in enumerate(col_positions):
                end = col_positions[i+1][0] if i + 1 < len(col_positions) else None
                col_slices[key] = slice(start, end)
            
            municipio_propagate = ""
            for linha in linhas[header_idx + 1:]:
                linha_upper = linha.upper()
                if "TOTAL" in linha_upper or "VAGAS" in linha_upper or "ASSINATURA" in linha_upper or "RELAÇÃO DE E-MAILS" in linha_upper: break
                
                municipio_raw = linha[col_slices["municipio"]].strip() if "municipio" in col_slices else ""
                disciplina_raw = linha[col_slices["disciplina"]].strip() if "disciplina" in col_slices else ""
                
                municipio_clean = clean_text(municipio_raw)
                disciplina_clean = clean_text(disciplina_raw)

                if len(municipio_clean) > 2 and not any(kw in municipio_clean.upper() for kw in ["COLUNA", "Nº", "CLASSIFICAÇÃO"]):
                    municipio_propagate = municipio_clean
                
                if len(disciplina_clean) > 2 and not any(kw in disciplina_clean.upper() for kw in ["COLUNA", "Nº", "CLASSIFICAÇÃO"]) and municipio_propagate:
                    registros.append({"Município": municipio_propagate, "Disciplina": disciplina_clean})

    if not registros and header_idx != -1:
        print("      -> AVISO: Análise por posição falhou ou não encontrou dados. Tentando método de divisão por espaços.")
        municipio_propagate = ""
        for linha in linhas[header_idx + 1:]:
            linha_upper = linha.upper()
            if "TOTAL" in linha_upper or "VAGAS" in linha_upper or "ASSINATURA" in linha_upper or "RELAÇÃO DE E-MAILS" in linha_upper: break
            
            colunas = [c.strip() for c in re.split(r'\s{2,}' , linha.strip()) if c.strip()]
            
            if len(colunas) >= 2:
                mun_candidato = ""
                disc_candidato = ""

                for c in colunas:
                    if any(m in c.upper() for m in keywords_municipio) or re.search(r'^[A-ZÀ-Ú]+\s*([A-ZÀ-Ú]+)?\s*$', c):
                        mun_candidato = c
                        break
                
                for c in colunas:
                    if any(d in c.upper() for d in keywords_disciplina) or re.search(r'^[A-ZÀ-Ú]+\s*([A-ZÀ-Ú]+)?\s*$', c):
                        if c != mun_candidato:
                            disc_candidato = c
                            break

                if len(mun_candidato) > 2 and not any(kw in mun_candidato.upper() for kw in ["COLUNA", "Nº", "CLASSIFICAÇÃO"]):
                    municipio_propagate = clean_text(mun_candidato)
                
                if len(disc_candidato) > 2 and not any(kw in disc_candidato.upper() for kw in ["COLUNA", "Nº", "CLASSIFICAÇÃO"]) and municipio_propagate:
                    registros.append({"Município": municipio_propagate, "Disciplina": clean_text(disc_candidato)})

    return registros

def extrair_informacoes_pdf(caminho_arquivo):
    status_flags = {'sem_vagas': False, 'imagem': False}
    try:
        with pdfplumber.open(caminho_arquivo) as pdf:
            texto_completo = "\n".join([pagina.extract_text(x_tolerance=2, layout=True) or "" for pagina in pdf.pages])

        nome_arquivo = os.path.basename(caminho_arquivo)
        print(f" -> Processando: {nome_arquivo}...")
        
        if len(texto_completo) < 250:
            print(f"      -> AVISO: PDF pode ser uma imagem ou está vazio. Extração ignorada.")
            status_flags['imagem'] = True
            return []

        pss_match = re.search(r'(?:PSS|PROCESSO SELETIVO SIMPLIFICADO)\s*N?º?\s*(\d+/\d{4})', texto_completo, re.IGNORECASE)
        pss = pss_match.group(1) if pss_match else "Não encontrado"
        
        ano_match = re.search(r'\b(20\d{2})\b', nome_arquivo + " " + texto_completo)
        ano = ano_match.group(1) if ano_match else (pss.split('/')[1] if '/' in pss else "Não encontrado")

        texto_upper = texto_completo.upper()
        dados_base = []

        keywords_disciplina_unica = ["AEE", "ATENDIMENTO EDUCACIONAL ESPECIALIZADO", "SALA DE RECURSO", "PROJETO MUNDI", "INTÉRPRETE DE LIBRAS", "ACOMPANHAMENTO ESPECIALIZADO", "ENSINO REGULAR", "SOME", "CONTADOR", "DOCENTE"]
        
        if "REFORÇO ESCOLAR" in texto_upper:
            dados_base = padrao_reforco_escolar(texto_completo=texto_completo, status_flags=status_flags)
        elif "PROJETO SEI" in texto_upper:
            dados_base = padrao_projeto_sei_refinado(texto_completo=texto_completo, status_flags=status_flags)
        elif "QUADRO DE VAGAS" in texto_upper:
            dados_base = padrao_tabela_generica(texto_completo=texto_completo, status_flags=status_flags)
            if not dados_base:
                dados_base = padrao_lista_disciplina_unica(texto_completo=texto_completo, status_flags=status_flags)
        elif any(kw in texto_upper for kw in keywords_disciplina_unica):
            dados_base = padrao_lista_disciplina_unica(texto_completo=texto_completo, status_flags=status_flags)
        
        if not dados_base:
            if not status_flags['sem_vagas'] and not status_flags['imagem']:
                print(f"      -> FALHA: Nenhum registro válido extraído de {nome_arquivo} com os padrões testados.")
            return []

        registros_finais = []
        for item in dados_base:
            municipio_clean = item.get("Município", "").replace('\n', ' ').strip()
            disciplina_clean = item.get("Disciplina", "").replace('\n', ' ').strip()

            if municipio_clean and disciplina_clean:
                item.update({ "Ano": ano, "PSS": pss, "Edital": nome_arquivo, "Município": municipio_clean, "Disciplina": disciplina_clean, "Arquivo_Origem": nome_arquivo, "Data_Extração": datetime.now().strftime('%Y-%m-%d') })
                registros_finais.append(item)
        
        print(f"      -> Sucesso! ({len(registros_finais)} registros extraídos)")
        return registros_finais

    except Exception as e:
        print(f"      -> ERRO GERAL ao processar {os.path.basename(caminho_arquivo)}: {e}")
        return []

def processar_pasta_e_salvar_planilha(caminho_da_pasta, nome_arquivo_saida):
    registros = []
    print("\n- - INÍCIO DA EXTRAÇÃO DE DADOS DOS ARQUIVOS PDF --- ")
    if not os.path.isdir(caminho_da_pasta):
        print(f"ERRO: A pasta '{caminho_da_pasta}' não foi encontrada!")
        return
        
    arquivos_na_pasta = [f for f in os.listdir(caminho_da_pasta) if f.lower().endswith('.pdf')]
    print(f"Encontrados {len(arquivos_na_pasta)} arquivos PDF em '{caminho_da_pasta}'.")
    
    for nome_arquivo in sorted(arquivos_na_pasta):
        caminho_completo = os.path.join(caminho_da_pasta, nome_arquivo)
        dados_extraidos = extrair_informacoes_pdf(caminho_completo)
        if dados_extraidos:
            registros.extend(dados_extraidos)

    if not registros:
        print("\n--- EXTRAÇÃO FINALIZADA ---")
        print("Nenhum registro foi extraído. A planilha não foi gerada.")
        return

    df = pd.DataFrame(registros)
    colunas_finais = ["Ano", "PSS", "Edital", "Município", "Disciplina", "Arquivo_Origem", "Data_Extração"]
    df = df.reindex(columns=colunas_finais)

    pasta_saida = os.path.dirname(nome_arquivo_saida)
    if pasta_saida and not os.path.exists(pasta_saida):
        os.makedirs(pasta_saida)
        
    df.to_csv(nome_arquivo_saida, index=False, sep=';', encoding='utf-8-sig')
    print("\n--- EXTRAÇÃO FINALIZADA ---")
    print(f"Total de arquivos PDF processados: {len(arquivos_na_pasta)}")
    print(f"Total de registros de vagas salvos na planilha: {len(registros)}")
    print(f"Dados salvos em '{nome_arquivo_saida}'!")

if __name__ == "__main__":
    pasta_saida_nome = "planilha conv especial"
    os.makedirs(pasta_saida_nome, exist_ok=True)

    caminho_log = os.path.join(pasta_saida_nome, "log_extracao.txt")
    arquivo_csv_saida = os.path.join(pasta_saida_nome, "convocacoes_especiais.csv")

    sys.stdout = Logger(caminho_log)
    print(f"--- LOG DE EXECUÇÃO - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ---")

    pasta_convocacoes_especiais = r"C:\Users\SEDUC\Desktop\conv especiais"
    
    try:
        processar_pasta_e_salvar_planilha(pasta_convocacoes_especiais, arquivo_csv_saida)
    except Exception as e:
        print(f"\n OCORREU UM ERRO INESPERADO QUE INTERROMPEU O SCRIPT ")
        print(f"ERRO: {e}")
    finally:
        print(f"\n--- FIM DO LOG ---")
