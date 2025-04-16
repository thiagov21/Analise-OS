import requests
from bs4 import BeautifulSoup
import re
import logging
from datetime import datetime, timedelta
import time
import os
import pandas as pd
import json

class ServiceOrderScraper:
    BASE_URL = "https://exemplo.exemplo.com.br"
    LOGIN_URL = f"{BASE_URL}/Account/Login"
    CONSULTA_URL = f"{BASE_URL}/OrdemServico/Consulta"
    COMPROVANTE_URL = f"{BASE_URL}/OrdemServico/AlteraComprovante"
    OUTPUT_FOLDER = "imagens_comprovantes"
    DATA_FOLDER = "Data"
    
    # dicionario de status
    # atualizar forma de mapear, processo paleativo ainda
    
    STATUS_MAPPING = {
        "1": "ABERTO",
        "2": "AGUARDANDO CONFIRMACAO DA PECA",
        "3": "PEÇA CONFIRMADA",
        "4": "AGUARDANDO APROVAÇÃO DO CLIENTE",
        "5": "AGUARDANDO PAGAMENTO",
        "6": "ENCAMINHADO PARA ESTORNO",
        "7": "ESTORNADO",
        "8": "AGUARDANDO COMPRA PECA",
        "9": "PECA EM TRANSITO",
        "10": "EM SERVICO",
        "11": "PRONTO",
        "12": "CONCLUIDO",
        "13": "DIRECIONAR PARA DESCARTE",
        "14": "AGUARDANDO RELOGIO PARA TROCA",
        "15": "AGUARDANDO ANALISE FORNITURA",
        "16": "AGUARDANDO ABONO COMERCIAL",
        "17": "ENVIADO PARA ANALISE SMARTWATCH",
        "18": "RECEBIDO TRANSBORDO",
        "19": "TENTAR ADAPTAÇÃO LOCAL",
        "20": "AGUARDANDO APROVAÇÃO DE ADAPTAÇÃO",
        "21": "AGUARDANDO APROVAÇÃO TROCA",
        "22": "AGUARDANDO DECISÃO DE TROCA/CLIENTE",
        "23": "TROCA REPROVADA",
        "24": "ANÁLISE EXCEPCIONAL",
        "25": "ADAPTAÇÃO APROVADA",
        "26": "ADAPTAÇÃO REPROVADA",
        "27": "SOLUÇÃO DO POSTO",
        "28": "SLA DE TROCA EXCEDIDO",
        "29": "DEVOLVER RELOGIO AO CLIENTE",
        "30": "DIRECIONAR PARA DESCARTE",
        "31": "ENVIADO PARA DESCARTE",
        "32": "RECEBIDO DESCARTE",
        "33": "RELOGIO DEVOLVIDO AO CLIENTE",
        "34": "RELOGIO DE TROCA ENVIADO AO POSTO",
        "35": "RELOGIO AGUARDANDO RETIRADA NO POSTO",
        "36": "RELOGIO DE TROCA ENVIADO AO CLIENTE"
    }
    
    def __init__(self):
        self.session = requests.Session()
        self.is_authenticated = False
        os.makedirs(self.OUTPUT_FOLDER, exist_ok=True)
        os.makedirs(self.DATA_FOLDER, exist_ok=True)
        
        # log config
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler('exemplo.log'),
                logging.StreamHandler()
            ]
        )
        
    def get_status_description(self, status_id: str) -> str:
        """Convert status ID to its corresponding description."""
        return self.STATUS_MAPPING.get(status_id, f"Status Desconhecido (ID: {status_id})")
        
    def login(self, username: str = "", password: str = "") -> bool:
        try:
            response = self.session.get(self.LOGIN_URL)
            soup = BeautifulSoup(response.text, "html.parser")
            token = soup.find("input", {"name": "__RequestVerificationToken"})
            
            if not token or 'value' not in token.attrs:
                logging.error("Could not find verification token")
                return False
                
            login_data = {
                "Login": username,
                "Password": password,
                "__RequestVerificationToken": token["value"]
            }
            
            response = self.session.post(self.LOGIN_URL, data=login_data)
            
            if "Autorizadas exemplo" in response.text:
                self.is_authenticated = True
                logging.info("Login successful")
                return True
            else:
                logging.error("Login failed")
                return False
                
        except Exception as e:
            logging.error(f"Login error: {str(e)}")
            return False

    def analyze_single_os(self, os_number: str) -> dict:
        if not self.is_authenticated:
            if not self.login():
                raise Exception("Authentication required")

        try:
            img_path = os.path.join(self.OUTPUT_FOLDER, f"comprovante_{os_number}.jpg")
            image_exists = os.path.exists(img_path)

            if not image_exists:
                url_os = f"{self.COMPROVANTE_URL}/{os_number}"
                response_os = self.session.get(url_os)
                
                if response_os.status_code == 200:
                    soup_os = BeautifulSoup(response_os.text, "html.parser")
                    media_images = soup_os.find_all("img")
                    
                    for media_image in media_images:
                        img_url = media_image.get("src")
                        
                        if img_url and img_url.startswith("/"):
                            img_url = f"{self.BASE_URL}{img_url}"
                        
                        if img_url and "exemplo.exemplo.com.br" in img_url and "/Imagens/Comprovantes/" in img_url:
                            try:
                                img_response = self.session.get(img_url)
                                img_response.raise_for_status()
                                
                                with open(img_path, "wb") as img_file:
                                    img_file.write(img_response.content)
                                image_exists = True
                                logging.info(f"Image downloaded for OS {os_number}")
                                break
                            except Exception as e:
                                logging.error(f"Error downloading image for OS {os_number}: {str(e)}")

            url = f"{self.CONSULTA_URL}/{os_number}"
            response = self.session.get(url)
            soup = BeautifulSoup(response.text, 'html.parser')
            
            # pegar o status da OS
            status_input = soup.find('input', {'id': 'StatusId'})
            if not status_input:
                status_input = soup.select_one('#demo-form2 > div:nth-child(8) > div:nth-child(1) > div > input.form-control')
            
            status_id = status_input['value'] if status_input else None
            status = self.get_status_description(status_id) if status_id else "Status não encontrado"

            cliente = soup.find('input', {'id': 'Cliente'})['value'] if soup.find('input', {'id': 'Cliente'}) else "Cliente não encontrado"
            garantia = soup.find('input', {'id': 'Garantia'})['value'] if soup.find('input', {'id': 'Garantia'}) else "Garantia não encontrada"
            referencia = soup.find('input', {'id': 'Referencia'})['value'] if soup.find('input', {'id': 'Referencia'}) else "Referência não encontrada"
            data_compra = soup.find('input', {'id': 'DataCompra'})['value'] if soup.find('input', {'id': 'DataCompra'}) else "Data de compra não encontrada"
            
            historico_status = []
            historico_table = soup.find('table', {'id': 'HistoricoStatus'})
            if historico_table:
                rows = historico_table.find_all('tr')
                for row in rows[1:]:  
                    cols = row.find_all('td')
                    if len(cols) >= 3:
                        historico_status.append({
                            "status": cols[0].text.strip(),
                            "dataHora": cols[1].text.strip(),
                            "alteradoPor": cols[2].text.strip()
                        })
            
            status_atual = historico_status[-1]['status'] if historico_status else "Status atual não encontrado"

            return {
                "url": url,
                "status": status,
                "status_id": status_id,
                "image_valid": image_exists,
                "image_path": img_path if image_exists else None,
                "cliente": cliente,
                "garantia": garantia,
                "statusAtual": status_atual,
                "referencia": referencia,
                "dataCompra": data_compra,
                "historicoStatus": historico_status,
                "numero_os": os_number
            }
            
        except Exception as e:
            logging.error(f"Error analyzing OS {os_number}: {str(e)}")
            raise

    def analyze_excel_file(self, progress_callback=None):
        try:
            excel_path = os.path.join(self.DATA_FOLDER, "trocas.xlsx")
            if not os.path.exists(excel_path):
                raise FileNotFoundError("Arquivo 'trocas.xlsx' não encontrado na pasta Data")
            
            df = pd.read_excel(excel_path)
            os_column = None
            
            # Tentar encontrar a coluna com números de OS
            possible_columns = ['OS', 'os', 'Ordem de Serviço', 'ordem_servico', 'numero_os']
            for col in possible_columns:
                if col in df.columns:
                    os_column = col
                    break
            
            if os_column is None:
                raise ValueError("Não foi possível encontrar a coluna com números de OS")
            
            results = []
            total_os = len(df[os_column])
            
            for idx, os_number in enumerate(df[os_column]):
                try:
                    result = self.analyze_single_os(str(os_number))
                    results.append(result)
                    
                    if progress_callback:
                        progress = (idx + 1) / total_os * 100
                        progress_callback(progress, f"Analisando OS {os_number} ({idx + 1}/{total_os})")
                        
                except Exception as e:
                    logging.error(f"Error processing OS {os_number}: {str(e)}")
                    results.append({
                        "numero_os": os_number,
                        "error": str(e)
                    })
                
                time.sleep(1)  # evitar sobrecarga no servidor
            
            return results
            
        except Exception as e:
            logging.error(f"Error analyzing Excel file: {str(e)}")
            raise

    def export_results(self, results, format="excel"):
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            
            if format == "excel":
                output_file = os.path.join(self.DATA_FOLDER, f"resultados_{timestamp}.xlsx")
                
                # formatar para excel
                excel_data = []
                for result in results:
                    row = {
                        "Número OS": result.get("numero_os", ""),
                        "Cliente": result.get("cliente", ""),
                        "Referência": result.get("referencia", ""),
                        "Garantia": result.get("garantia", ""),
                        "Status": result.get("status", ""),
                        "Data Compra": result.get("dataCompra", ""),
                        "Imagem": "Sim" if result.get("image_valid", False) else "Não",
                        "Caminho Imagem": result.get("image_path", ""),
                        "Erro": result.get("error", "")
                    }
                    excel_data.append(row)
                
                df = pd.DataFrame(excel_data)
                df.to_excel(output_file, index=False)
                return output_file
                
            elif format == "json":
                output_file = os.path.join(self.DATA_FOLDER, f"resultados_{timestamp}.json")
                with open(output_file, 'w', encoding='utf-8') as f:
                    json.dump(results, f, ensure_ascii=False, indent=2)
                return output_file
                
        except Exception as e:
            logging.error(f"Error exporting results: {str(e)}")
            raise