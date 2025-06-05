import openai as gpt
import _loggin as log
import os
from dotenv import load_dotenv

load_dotenv()
api_key = os.getenv("OPENAI_API_KEY")

ator = "Você é um especialista em análise de Notas Fiscais brasileiras"

messages = [{"role": "system", "content": f"[{ator}]"}]

try:

    client = gpt.OpenAI(api_key=api_key)

    def ConsultaFarmaceutico(prompt):

       resposta = client.chat.completions.create(
           model="gpt-4o-mini",
           messages=[
               {"role": "system", "content": ator},
               {"role": "user", "content": prompt},
           ],
           temperature=0.1,  # quanto mais alto mais e aleatório/criativo quanto menor mais preciso, vai de 0 a 1.
       )

       assistant_reply = resposta.choices[0].message.content

       return assistant_reply

except Exception as e:
    log.log(f"Erro ao consultar GPT: {e}")
