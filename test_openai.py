import os
import json
from dotenv import load_dotenv
from openai import OpenAI

load_dotenv()

client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))

PROMPT = """
Devuelve EXCLUSIVAMENTE un JSON válido con este formato exacto:

{
  "modulos": [
    "string",
    "string",
    "string",
    "string",
    "string",
    "string",
    "string",
    "string"
  ]
}

REGLAS OBLIGATORIAS:
- Exactamente 8 módulos
- SOLO títulos (NO descripciones)
- Máximo 12 palabras por módulo
- En español
- NO numeración
- NO texto fuera del JSON
- NO explicaciones
"""

response = client.responses.create(
    model="gpt-5-mini",
    input=[
        {"role": "system", "content": "Eres un asistente académico extremadamente estricto."},
        {
            "role": "user",
            "content": f"Diplomado: Auxiliar de Educación\n{PROMPT}"
        }
    ]
)

raw_output = response.output_text

print("RAW OUTPUT:\n")
print(raw_output)

print("\n--- PARSEANDO JSON ---\n")

data = json.loads(raw_output)

for i, modulo in enumerate(data["modulos"], start=1):
    print(f"{i}. {modulo}")
