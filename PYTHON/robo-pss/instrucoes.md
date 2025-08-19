# Web Scraping

## Objetivo

Desenvolver um script automatizado para:

- Acessar o portal de PSS da SEDUC PA.
- Extrair todos os arquivos PDF referentes a Processos Seletivos Simplificados.
- Organizar os arquivos em pastas seguindo uma estrutura específica.
- Extrair informações específicas dos PDFs de convocações especiais e consolidá-las em uma planilha.

## Detalhes importantes

- Ao extrair os arquivos, todos devem ser armazenados em uma pasta com o nome do respectivo processo seletivo
- Em cada pasta, criar uma subpasta chamada 'CONVOCACOES ESPECIAIS'
- Incluir nestas pastas os arquivos que contenham a 'CONVOCACAO ESPECIAL'

## Extrair Dados das Convocações Especiais

Para cada PDF na pasta CONVOCACOES ESPECIAIS:

1. texto do PDF usando biblioteca especializada (PyPDF2 ou pdfplumber).

2. Buscar e capturar as seguintes informações:
 - Ano: Extrair do texto ou do nome do arquivo/pasta (padrão: 4 dígitos consecutivos)
 - PSS: Identificar código/número do processo seletivo
 - Nome do Edital: Extrair do texto
 - Município: Localizar no texto (buscar por "Município:", "Local:", ou similar)
 - Disciplina: Identificar área de atuação (buscar por "Disciplina:", "Área:", ou similar)

## Consolidar Dados em Planilha

1. Criar/atualizar arquivo convocações_especiais.xlsx (formato Excel) ou convocações_especiais.csv.
2. Estrutura da planilha:

Ano, PSS, Edital,	Município,	Disciplina,	Arquivo_Origem,	Data_Extração

3. Adicionar os dados extraídos de todas as convocações especiais.
4. Incluir timestamp de quando a extração foi realizada.

## Link Seduc

https://www.seduc.pa.gov.br/pagina/12173-acompanhe-todos-os-pss-da-seduc
