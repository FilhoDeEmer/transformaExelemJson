import json

# Função para carregar e analisar o arquivo JSON
def analisar_json(arquivo):
    # Abrindo o arquivo JSON
    with open(arquivo, 'r',encoding='utf-8') as file:
        dados = json.load(file)  # Lê e converte o JSON para um dicionário Python
    
    # Verifica se há dados na chave 'Planilha1'
    planilha1 = dados.get('Planilha1', [])
    
    if planilha1:  # Se houver dados dentro de 'Planilha1'
        for item in planilha1:
            usuarios = item.get('USUÁRIO')
            atendimentos = item.get('ATENDIMENTOS', [])
            
            print(f"Usuários: {usuarios}")
            
            for atendimento in atendimentos:
                data = atendimento.get('DT ATENDIMENTO')
                lista_atendimentos = atendimento.get('ATENDIMENTOS', [])
                
                print(f"Data: {data}")
                print(f"Atendimentos: {lista_atendimentos}")
    else:
        print("Não foi encontrado o campo 'Planilha1'.")

# Chamar a função passando o caminho do arquivo JSON
analisar_json('arquivo.json')
