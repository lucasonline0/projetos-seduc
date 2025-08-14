import os
import re
import time
import requests
import logging
import sys
from pathlib import Path
from urllib.parse import urljoin, urlparse
from bs4 import BeautifulSoup
from typing import List, Dict

class SeducPSScraper:
    def __init__(self):
        self.base_url = "https://www.seduc.pa.gov.br/pagina/12173-acompanhe-todos-os-pss-da-seduc"
        self.downloads_folder = Path.home() / "Desktop" / "Downloads PSS"
        if not self.downloads_folder.parent.exists():
            self.downloads_folder = Path.cwd() / "Downloads PSS"
        
        self.downloads_folder.mkdir(exist_ok=True)

        self.session = requests.Session()
        self.session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        })
        
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')

        # cria log com timestamp por execução

        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        log_file = f'seduc_scraper_{timestamp}.log'

        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler('seduc_scraper_otimizado.log'),
                logging.StreamHandler()
            ]
        )
        self.logger = logging.getLogger(__name__)
        self.stats = {'encontrados': 0, 'baixados': 0, 'erros': 0}
        
        print(f"🚀 Seduc PSS Scraper Otimizado")
        print(f"📂 Pasta: {self.downloads_folder}")
        print(f"📝 Log: {log_file}")

    def run(self):
        try:
            self.logger.info("🚀 Iniciando scraper otimizado")
            
            pss_pages = self.find_pss_pages()
            if not pss_pages:
                self.logger.error("❌ Nenhuma página PSS encontrada")
                return
            
            all_pdfs = []
            for i, page in enumerate(pss_pages, 1):
                self.logger.info(f"[{i}/{len(pss_pages)}] 🔍 {page['name']}")
                pdfs = self.extract_pdfs(page['url'], page['name'])
                all_pdfs.extend(pdfs)
                time.sleep(1)  
            
            self.logger.info(f"📊 Total encontrados: {len(all_pdfs)}")
            self.download_pdfs(all_pdfs)
            
            self.print_report()
            
        except Exception as e:
            self.logger.error(f"❌ Erro geral: {e}")
            self.stats['erros'] += 1

    def find_pss_pages(self) -> List[Dict]:
        try:
            response = self.session.get(self.base_url, timeout=30)
            response.raise_for_status()
            soup = BeautifulSoup(response.content, 'html.parser')
            
            pages = []
            for link in soup.find_all('a', href=True):
                href = link.get('href')
                text = link.get_text(strip=True)
                
                if any(term in f"{href} {text}".upper() for term in ['PSS', 'PROCESSO SELETIVO', 'SELETIVO SIMPLIFICADO']):
                    name = self.extract_pss_name(text, href)
                    pages.append({
                        'name': name,
                        'url': urljoin(self.base_url, href)
                    })
            
            unique_pages = []
            seen_urls = set()
            for page in pages:
                if page['url'] not in seen_urls:
                    seen_urls.add(page['url'])
                    unique_pages.append(page)
            
            self.logger.info(f"✅ Encontradas {len(unique_pages)} páginas PSS")
            return unique_pages
            
        except Exception as e:
            self.logger.error(f"❌ Erro ao buscar páginas PSS: {e}")
            return []

    def extract_pss_name(self, text: str, href: str) -> str:
        combined = f"{text} {href}".upper()
        
        patterns = [
            r'PSS[_\s-]*(\d+)[_\s-]*(\d{4})[_\s-]*([A-Z]+)',
            r'PSS[_\s-]*(\d+)[_\s-]*(\d{4})',
        ]
        
        for pattern in patterns:
            match = re.search(pattern, combined)
            if match:
                groups = match.groups()
                if len(groups) == 3:
                    return f"PSS_{groups[0]}_{groups[1]}_{groups[2]}"
                elif len(groups) == 2:
                    return f"PSS_{groups[0]}_{groups[1]}"
        
        clean = re.sub(r'[^\w\s]', '_', text)[:30]
        return clean if clean else "PSS_GERAL"

    def extract_pdfs(self, url: str, pss_name: str) -> List[Dict]:
        try:
            response = self.session.get(url, timeout=30)
            response.raise_for_status()
            soup = BeautifulSoup(response.content, 'html.parser')
            
            pdfs = []
            download_section = self.find_download_section(soup)
            links = download_section.find_all('a') if download_section else soup.find_all('a')
            
            for link in links:
                arquivo_download = link.get('arquivo_download', '')
                href = link.get('href', '')
                url_arquivo = arquivo_download or href
                
                if url_arquivo and self.is_valid_file(url_arquivo):
                    full_url = urljoin(url, url_arquivo)
                    filename = self.extract_filename(url_arquivo)
                    is_especial = 'CONVOCACAO' in f"{link.get_text()} {url_arquivo}".upper()
                    
                    pdfs.append({
                        'url': full_url,
                        'filename': filename,
                        'pss_name': pss_name,
                        'is_especial': is_especial
                    })
            
            self.logger.info(f"📄 {len(pdfs)} arquivos encontrados em {pss_name}")
            self.stats['encontrados'] += len(pdfs)
            return pdfs
            
        except Exception as e:
            self.logger.error(f"❌ Erro em {pss_name}: {e}")
            self.stats['erros'] += 1
            return []

    def find_download_section(self, soup):
        for element in soup.find_all(text=lambda t: t and 'ARQUIVOS PARA DOWNLOAD' in t.upper()):
            parent = element.parent
            while parent and parent.name != 'html':
                arquivo_links = parent.find_all('a', attrs={'arquivo_download': True})
                if len(arquivo_links) > 1:
                    return parent
                parent = parent.parent
        
        for container in soup.find_all(['div', 'section']):
            arquivo_links = container.find_all('a', attrs={'arquivo_download': True})
            if len(arquivo_links) >= 2:
                return container
        
        return None

    def is_valid_file(self, url: str) -> bool:
        valid_extensions = ['.pdf', '.doc', '.docx', '.xls', '.xlsx', '.txt', '.zip']
        url_lower = url.lower()
        
        return (
            any(url_lower.endswith(ext) for ext in valid_extensions) or
            any(ext in url_lower for ext in valid_extensions) or
            'arquivo' in url_lower or 'download' in url_lower
        )

    def extract_filename(self, url: str) -> str:
        filename = os.path.basename(urlparse(url).path)
        if not filename or '.' not in filename:
            filename = f"documento_{hash(url) % 10000}.pdf"
        return re.sub(r'[<>:"/\\|?*]', '_', filename)

    def download_pdfs(self, pdfs: List[Dict]):
        if not pdfs:
            return
        
        self.logger.info(f"🔄 Iniciando download de {len(pdfs)} arquivos")
        pss_groups = {}
        for pdf in pdfs:
            pss_name = pdf['pss_name']
            if pss_name not in pss_groups:
                pss_groups[pss_name] = []
            pss_groups[pss_name].append(pdf)

        for pss_name, pss_pdfs in pss_groups.items():
            self.download_pss_files(pss_name, pss_pdfs)

    def download_pss_files(self, pss_name: str, pdfs: List[Dict]):
        pss_folder = self.downloads_folder / self.sanitize_name(pss_name)
        especiais_folder = pss_folder / "CONVOCACOES_ESPECIAIS"
        
        pss_folder.mkdir(exist_ok=True)
        especiais_folder.mkdir(exist_ok=True)
        
        self.logger.info(f"📂 Baixando {pss_name} ({len(pdfs)} arquivos)")
        
        for i, pdf in enumerate(pdfs, 1):
            try:
                target_folder = especiais_folder if pdf['is_especial'] else pss_folder
                filepath = target_folder / pdf['filename']
                
                if filepath.exists():
                    continue

                response = self.session.get(pdf['url'], timeout=60)
                response.raise_for_status()
                
                filepath.write_bytes(response.content)
                self.logger.info(f"✅ [{i}/{len(pdfs)}] {pdf['filename']}")
                self.stats['baixados'] += 1
                time.sleep(0.5) 
                
            except Exception as e:
                self.logger.error(f"⚠️ Erro ao baixar {pdf['filename']}: {e}")
                self.stats['erros'] += 1

    def sanitize_name(self, name: str) -> str:
        return re.sub(r'[<>:"/\\|?*]', '_', name)

    def print_report(self):
        print("\n" + "="*60)
        print("📊 RELATÓRIO FINAL")
        print("="*60)
        print(f"📄 Arquivos encontrados: {self.stats['encontrados']}")
        print(f"✅ Arquivos baixados: {self.stats['baixados']}")
        print(f"❌ Erros: {self.stats['erros']}")
        print(f"📁 Pasta: {self.downloads_folder}")
        print("="*60)

def main():
    try:
        scraper = SeducPSScraper()
        scraper.run()
        return 0
    except KeyboardInterrupt:
        print("\n Interrompido pelo usuário")
        return 1
    except Exception as e:
        print(f"❌ Erro fatal: {e}")
        return 1

if __name__ == "__main__":
    exit(main())