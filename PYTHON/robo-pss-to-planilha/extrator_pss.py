import os
import re
import sys
import pandas as pd
import pdfplumber
from datetime import datetime

class Logger(object):
    def __init__(self, filename="log_extracao.txt"):
        self.terminal = sys.stdout
        self.log = open(filename, "w", encoding='utf-8')

    def write(self, message):
        self.terminal.write(message)
        self.log.write(message)

    def flush(self):
        pass

def padrao_reforco_escolar(texto_completo, **kwargs):

    print("      -> Padrão detectado: Projeto Reforço Escolar")
    registros = []
    
    disciplinas_match = re.search(r'licenciatura em (letras com habilitação em língua portuguesa e licenciatura em matemática|pedagogia.*?e.*?educação física)', texto_completo, re.IGNORECASE)
    disciplinas = []
    if disciplinas_match:
        texto_disciplinas = disciplinas_match.group(1)
        if "portuguesa" in texto_disciplinas and "matemática" in texto_disciplinas:
            disciplinas = ["Língua Portuguesa", "Matemática"]
        elif "pedagogia" in texto_disciplinas and "física" in texto_disciplinas:
            disciplinas = ["Pedagogia", "Educação Física"]

    if not disciplinas:
        print("      -> AVISO: Não foi possível extrair as disciplinas do texto inicial.")
        return []
    
    partes = re.split(r'QUADRO DE VAGAS|MUNICÍPIO DE LOTAÇÃO', texto_completo, flags=re.IGNORECASE)
    texto_lista = partes[-1] if len(partes) > 1 else ""
    
    municipios = re.findall(r'^\s*([A-ZÇÃ-ÕÁ-Ú][a-zA-Zçã-õá-ú\s-]+?)\s+\d+\s*$', texto_lista, re.MULTILINE)
    
    for mun in municipios:
        mun_clean = mun.strip()
        if len(mun_clean) > 2:
            for disc in disciplinas:
                registros.append({"Município": mun_clean, "Disciplina": disc})
    return registros

def padrao_projeto_sei_refinado(texto_completo, **kwargs):

    print("      -> Padrão detectado: Projeto SEI (Refinado)")
    registros = []
    
    partes_fim = re.split(r'RELAÇÃO DE E-MAILS|RELAÇÃO DE E-MAIL', texto_completo, flags=re.IGNORECASE)
    texto_relevante = partes_fim[0]

    partes_inicio = re.split(r'QUADRO DE VAGAS', texto_relevante, flags=re.IGNORECASE)
    texto_tabela = partes_inicio[-1] if len(partes_inicio) > 1 else ""

    disciplina = "Docente Projeto SEI"
    ure_atual = "Não Identificada"
    
    for linha in texto_tabela.strip().split('\n'):
        if "TOTAL DE VAGAS" in linha.upper(): break
        ure_match = re.search(r'(\d+ª?\s*URE\s*[-\s]*[A-ZÇÃ-ÕÁ-Ú]+)', linha, re.IGNORECASE)
        if ure_match:
            ure_atual = ure_match.group(1).strip()
        else:
            localidade = re.sub(r'\s+\d+\s*Manhã/Tarde/Noite.*$', '', linha).strip()
            if len(localidade) > 3 and "VAGAS" not in localidade.upper():
                registros.append({"Município": f"{localidade} ({ure_atual})", "Disciplina": disciplina})
    return registros

def padrao_tabela_generica(texto_completo, **kwargs):

    print("      -> Padrão detectado: Tabela Genérica Flexível")
    partes = re.split(r'QUADRO DE VAGAS', texto_completo, flags=re.IGNORECASE)
    texto_tabela = partes[-1] if len(partes) > 1 else ""
    linhas = texto_tabela.strip().split('\n')
    registros = []

    idx_municipio, idx_disciplina = -1, -1
    data_start_index = 0
    
    texto_cabecalho = " ".join(linhas[:5]).upper().replace('\n', ' ')
    
    municipio_kw = ["LOCALIDADE PARA PROVIMENTO", "MUNICÍPIO DE LOTAÇÃO", "MUNICÍPIO"]
    disciplina_kw = ["DISCIPLINA", "FUNÇÃO", "CARGO", "COMPONENTE CURRICULAR"]

    pos_mun, pos_disc = -1, -1
    for kw in municipio_kw:
        if kw in texto_cabecalho:
            pos_mun = texto_cabecalho.find(kw)
            break
    for kw in disciplina_kw:
        if kw in texto_cabecalho:
            pos_disc = texto_cabecalho.find(kw)
            break
            
    if pos_mun != -1 and pos_disc != -1:
        colunas_header = [col.strip() for col in re.split(r'\s{3,}', linhas[2])]
        if len(colunas_header) > 1:
            if pos_mun < pos_disc:
                idx_municipio = 1
                idx_disciplina = 2
            else:
                idx_municipio = 2
                idx_disciplina = 1
        data_start_index = 3 
    else:
         print("      -> AVISO: Não foi possível determinar a ordem das colunas. Usando posições padrão.")
         idx_municipio = 0
         idx_disciplina = 1
         data_start_index = 1
    
    municipio_propagate = ""
    for linha in linhas[data_start_index:]:
        if "TOTAL DE VAGAS" in linha.upper(): break
        colunas = re.split(r'\s{2,}', linha.strip())

        if len(colunas) <= max(idx_municipio, idx_disciplina): continue

        municipio_raw = colunas[idx_municipio].strip()
        disciplina_raw = colunas[idx_disciplina].strip()
        
        if len(municipio_raw) > 3:
            municipio_propagate = municipio_raw
        elif not municipio_raw and municipio_propagate:
             municipio_raw = municipio_propagate
        
        if len(municipio_raw) > 3 and len(disciplina_raw) > 3:
            municipios_col_a = colunas[0].strip()
            if len(municipios_col_a) > len(municipio_raw):
                municipios_para_add = [m.strip() for m in re.split(r',|\n', municipios_col_a) if len(m.strip()) > 2]
                for mun in municipios_para_add:
                    registros.append({"Município": municipio_raw, "Disciplina": disciplina_raw})
            else:
                registros.append({"Município": municipio_raw, "Disciplina": disciplina_raw})
    return registros

def extrair_informacoes_pdf(caminho_arquivo):
    try:
        with pdfplumber.open(caminho_arquivo) as pdf:
            texto_completo = "\n".join([pagina.extract_text(x_tolerance=2, layout=True) or "" for pagina in pdf.pages])

        nome_arquivo = os.path.basename(caminho_arquivo)
        print(f" -> Processando: {nome_arquivo}...")

        pss_match = re.search(r'(?:PSS|PROCESSO SELETIVO SIMPLIFICADO)\s*N?º?\s*(\d{2,3}/\d{4})', texto_completo, re.IGNORECASE)
        pss = pss_match.group(1) if pss_match else "Não encontrado"
        ano_match = re.search(r'(\b20\d{2}\b)', nome_arquivo)
        ano = ano_match.group(1) if ano_match else (pss.split('/')[1] if '/' in pss else "Não encontrado")

        texto_upper = texto_completo.upper()
        extrator_selecionado = None
        
        if "REFORÇO ESCOLAR" in texto_upper:
            extrator_selecionado = padrao_reforco_escolar
        elif "PROJETO SEI" in texto_upper:
            extrator_selecionado = padrao_projeto_sei_refinado
        elif any(kw in texto_upper for kw in ["SOME", "CEMEP", "EJA CAMPO", "EDUCAÇÃO ESPECIAL"]) or "ENSINO REGULAR" in texto_upper:
            extrator_selecionado = padrao_tabela_generica
        
        if not extrator_selecionado:
            print(f"      -> AVISO: Nenhum padrão de extração conhecido foi encontrado para {nome_arquivo}.")
            return []

        dados_base = extrator_selecionado(texto_completo=texto_completo)

        registros_finais = []
        for item in dados_base:
            item.update({
                "Ano": ano, "PSS": pss, "Edital": nome_arquivo,
                "Arquivo_Origem": nome_arquivo, "Data_Extração": datetime.now().strftime('%Y-%m-%d')
            })
            registros_finais.append(item)
        
        if not registros_finais:
            print(f"      -> FALHA: Nenhum registro válido extraído de {nome_arquivo} com o padrão selecionado.")
        else:
            print(f"      -> Sucesso! ({len(registros_finais)} registros extraídos)")
            
        return registros_finais

    except Exception as e:
        print(f"      -> ERRO GERAL ao processar {nome_arquivo}: {e}")
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
    df = df[["Ano", "PSS", "Edital", "Município", "Disciplina", "Arquivo_Origem", "Data_Extração"]]
    df.to_csv(nome_arquivo_saida, index=False, sep=';', encoding='utf-8-sig')
    print("\n--- EXTRAÇÃO FINALIZADA ---")
    print(f"Total de arquivos PDF processados: {len(arquivos_na_pasta)}")
    print(f"Total de registros de vagas salvos na planilha: {len(registros)}")
    print(f"Dados salvos em '{nome_arquivo_saida}'!")

if __name__ == "__main__":
    sys.stdout = Logger()
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