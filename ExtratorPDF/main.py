# RPA para Extração de Dados de Notas Fiscais
# Usando LlamaIndex + PyMuPDF + OpenAI

import os
import json
import pandas as pd
from pathlib import Path
from datetime import datetime
from typing import Dict, List, Optional
import logging
import consulta_gpt as gpt
import time
from llama_index.core.prompts import PromptTemplate
import fitz  # PyMuPDF

# Configuração de logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class NFExtractor:
    def __init__(self, model_name: str = "llama3.2:3b"):
        """
        Inicializa o extrator de NF
        Args:
            model_name: Nome do modelo Ollama (recomendado: llama3.2:3b para velocidade)
        """
        self.model_name = model_name
        self.llm = None
        self.setup_llm()

        # Template para extração estruturada
        self.extraction_template = PromptTemplate(
            """
            Você é um especialista em análise de Notas Fiscais brasileiras.
            Extraia APENAS as seguintes informações do texto da NF fornecido:

            DADOS OBRIGATÓRIOS:
            - Número da NF
            - Data de emissão
            - CNPJ do emitente
            - Razão social do emitente
            - CNPJ do destinatário (se houver)
            - Razão social do destinatário
            - Valor total da NF
            - Valor do ICMS (se houver)
            - Valor do IPI (se houver)

            PRODUTOS/SERVIÇOS (primeiros 3 itens):
            - Descrição do produto/serviço
            - Quantidade
            - Valor unitário
            - Valor total do item

            FORMATO DE RESPOSTA (JSON):
            {{
                "numero_nf": "",
                "data_emissao": "",
                "cnpj_emitente": "",
                "razao_social_emitente": "",
                "cnpj_destinatario": "",
                "razao_social_destinatario": "",
                "valor_total": "",
                "valor_icms": "",
                "valor_ipi": "",
                "itens": [
                    {{
                        "descricao": "",
                        "quantidade": "",
                        "valor_unitario": "",
                        "valor_total": ""
                    }}
                ]
            }}

            TEXTO DA NOTA FISCAL:
            {document_text}

            RESPOSTA (apenas JSON válido):
            """
        )

    def extract_text_from_pdf(self, pdf_path: str) -> str:
        """
        Extrai texto do PDF usando PyMuPDF, com tratamento robusto de erros.

        Args:
            pdf_path: Caminho completo do arquivo PDF

        Returns:
            Texto extraído do PDF ou string vazia se falhar
        """
        try:
            pdf_path = str(Path(pdf_path))  # Garante formatação consistente
            if not os.path.isfile(pdf_path):
                logger.warning(f"Arquivo não encontrado: {pdf_path}")
                return ""

            with fitz.open(pdf_path) as doc:
                texto_extraido = "\n".join(page.get_text() for page in doc)
                return texto_extraido.strip()

        except Exception as e:
            logger.error(f"Erro ao extrair texto do PDF '{pdf_path}': {e}")
            return ""

    def extract_text_from_pdf(self, pdf_path: str) -> str:
        """
        Extrai texto do PDF usando PyMuPDF
        Args:
            pdf_path: Caminho para o arquivo PDF
        Returns:
            Texto extraído do PDF
        """
        try:
            doc = fitz.open(pdf_path)
            text = ""

            for page_num in range(doc.page_count):
                page = doc[page_num]
                text += page.get_text()

            doc.close()
            #print(f"PDF>>>>>>> {text.strip()}")
            return text.strip()

        except Exception as e:
            logger.error(f"Erro ao extrair texto do PDF {pdf_path}: {e}")
            return ""

    def clean_json_response(self, response: str) -> str:
        #Limpa a resposta para extrair apenas o JSON válido
        try:
            # Procura pelo primeiro { e último }
            start = response.find('{')
            end = response.rfind('}') + 1

            if start != -1 and end != 0:
                return response[start:end]
            else:
                return response
        except:
            return response

    def extract_nf_data(self, pdf_path: str) -> Optional[Dict]:
        """
        Extrai dados estruturados de uma NF
        Args:
            pdf_path: Caminho para o arquivo PDF da NF
        Returns:
            Dicionário com dados extraídos ou None se houver erro
        """
        try:
            logger.info(f"Processando: {pdf_path}")


            text = self.extract_text_from_pdf(pdf_path) # Extrai texto do PDF
            if not text:
                logger.warning(f"Não foi possível extrair texto de {pdf_path}")
                return None

            text = text[:5000] # Limita o texto para evitar context overflow

            prompt = self.extraction_template.format(document_text=text)

            response_text = gpt.ConsultaFarmaceutico(prompt)

            json_str = self.clean_json_response(response_text) # Limpa e parseia a resposta JSON

            try:
                data = json.loads(json_str)
                data['arquivo_origem'] = os.path.basename(pdf_path)
                data['data_processamento'] = datetime.now().isoformat()
                return data

            except json.JSONDecodeError as je:
                logger.error(f"Erro ao parsear JSON de {pdf_path}: {je}")
                logger.debug(f"Resposta recebida: {json_str}")
                return None

        except Exception as e:
            logger.error(f"Erro ao processar {pdf_path}: {e}")
            return None

    def process_folder(self, folder_path: str, output_file: str = "nfs_extraidas.xlsx") -> List[Dict]:
        """
        Processa uma pasta com arquivos PDF de NFs

        Args:
            folder_path: Caminho da pasta com os PDFs
            output_file: Nome do arquivo de saída Excel

        Returns:
            Lista com dados extraídos de todas as NFs
        """
        folder_path = Path(folder_path)

        if not folder_path.exists():
            raise FileNotFoundError(f"Pasta não encontrada: {folder_path}")

        pdf_files = list(folder_path.glob("*.pdf")) # Encontra todos os PDFs

        if not pdf_files:
            logger.warning(f"Nenhum arquivo PDF encontrado em {folder_path}")
            return []

        logger.info(f"Encontrados {len(pdf_files)} arquivos PDF")

        results = []

        for pdf_file in pdf_files:
            try:
                data = self.extract_nf_data(str(pdf_file))
                if data:
                    results.append(data)
                    logger.info(f"✓ Processado: {pdf_file.name}")
                else:
                    logger.warning(f"✗ Falha ao processar: {pdf_file.name}")

            except Exception as e:
                logger.error(f"Erro ao processar {pdf_file}: {e}")
                continue

        # Salva resultados
        if results:
            self.save_results(results, output_file)

        return results

    def save_results(self, results: List[Dict], output_file: str):
        #Salva resultados em Excel e JSON
        try:
            # Prepara dados para DataFrame
            df_data = []

            for result in results:
                row = {
                    'arquivo_origem': result.get('arquivo_origem', ''),
                    'numero_nf': result.get('numero_nf', ''),
                    'data_emissao': result.get('data_emissao', ''),
                    'cnpj_emitente': result.get('cnpj_emitente', ''),
                    'razao_social_emitente': result.get('razao_social_emitente', ''),
                    'cnpj_destinatario': result.get('cnpj_destinatario', ''),
                    'razao_social_destinatario': result.get('razao_social_destinatario', ''),
                    'valor_total': result.get('valor_total', ''),
                    'valor_icms': result.get('valor_icms', ''),
                    'valor_ipi': result.get('valor_ipi', ''),
                    'data_processamento': result.get('data_processamento', '')
                }
                df_data.append(row)

            # Salva Excel
            df = pd.DataFrame(df_data)
            excel_file = output_file
            df.to_excel(excel_file, index=False)
            logger.info(f"Dados salvos em: {excel_file}")

            # Salva JSON completo (com itens detalhados)
            json_file = output_file.replace('.xlsx', '.json')
            with open(json_file, 'w', encoding='utf-8') as f:
                json.dump(results, f, indent=2, ensure_ascii=False)
            logger.info(f"Dados detalhados salvos em: {json_file}")

        except Exception as e:
            logger.error(f"Erro ao salvar resultados: {e}")

def main():

    start_time = time.time()

    # Configurações
    PASTA_NFS = r"C:\Users\israel.ribeiro\PyCharmMiscProject\LeitorPDFChat\NFs"
    ARQUIVO_SAIDA = r"outputs\resultados_nfs.xlsx"
    MODELO_LLM = "llama3.2:3b"

    print("Iniciando RPA de Extração de Notas Fiscais")
    print(f"Pasta: {PASTA_NFS}")
    #print(f"Modelo: {MODELO_LLM}")
    print("-" * 50)

    try:

        extractor = NFExtractor(model_name=MODELO_LLM) # Inicializa extrator

        results = extractor.process_folder(PASTA_NFS, ARQUIVO_SAIDA) # Processa pasta

        # Relatório final
        print(f"\nProcessamento concluído!")
        print(f"Total de NFs processadas: {len(results)}")
        print(f"Arquivo gerado: {ARQUIVO_SAIDA}")

        """
        if results:
            print("\nResumo dos primeiros resultados:")
            for i, result in enumerate(results[:3], 1):
                print(f"{i}. {result.get('arquivo_origem', 'N/A')} - NF: {result.get('numero_nf', 'N/A')}")
        """

    except Exception as e:
        logger.error(f"Erro na execução principal: {e}")
        print(f"Erro: {e}")

    end_time = time.time()
    elapsed_time = end_time - start_time
    print(f"\n Tempo total de execução:\033[1m {elapsed_time:.2f} \033[0msegundos")

if __name__ == "__main__":
    main()
