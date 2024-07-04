import requests
import pandas as pd

# URL da API OLinda com filtro para taxa SELIC
url = "https://api.bcb.gov.br/dados/serie/bcdata.sgs.432/dados?formato=json"

# Fazer a requisição GET
response = requests.get(url)

# Verificar o status da resposta
if response.status_code == 200:
    try:
        # Tentar converter a resposta em JSON
        data = response.json()
        
        # Converter os dados em um DataFrame do pandas
        df = pd.DataFrame(data)
        
        # Convertendo a coluna 'data' para o tipo datetime
        df['data'] = pd.to_datetime(df['data'], format='%d/%m/%Y')
        
        # Convertendo a coluna 'valor' para float
        df['valor'] = df['valor'].astype(float)
        
        # Exibir os primeiros registros
        print(df.head())
    except requests.exceptions.JSONDecodeError:
        print("Erro ao decodificar JSON. Conteúdo da resposta não é um JSON válido.")
        print("Conteúdo da resposta:", response.text)  # Exibir o conteúdo bruto da resposta para depuração
else:
    print(f"Erro na requisição: {response.status_code}")
    print("Conteúdo da resposta:", response.text)  # Mostrar o conteúdo da resposta para depuração