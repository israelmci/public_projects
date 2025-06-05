import logging
import logging as log

log.basicConfig(
    filename=r"logs\erros_leitorPDF.log",
    level=log.ERROR,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

def log(mensagem):
    logging.error(mensagem)

