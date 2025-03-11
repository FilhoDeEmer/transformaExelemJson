import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

def analisar_excel():
    arquivo = filedialog.askopenfilename(title="Selecione o arquivo Excel", filetypes=[("Arquivos Excel", "*.xlsx")])
    if not arquivo:
        return
    try:
        # Lê o arquivo Excel
        dados = pd.read_excel(arquivo, sheet_name=0, header=None)

        # Dicionário para armazenar os dados organizados
        usuarios_data = []

        usuario = None
        data_atendimento = None

        for index, row in dados.iterrows():
            # Verifica se a linha contém "USUÁRIO:" e extrai o nome do usuário
            if isinstance(row[0], str) and "USUÁRIO:" in row[0]:
                usuario = row[0].replace("USUÁRIO:", "").strip()

            # Verifica se a linha contém "DT ATENDIMENTO:" e extrai a data
            elif isinstance(row[0], str) and "DT ATENDIMENTO:" in row[0]:
                data_atendimento = row[0].replace("DT ATENDIMENTO:", "").strip()
            else:
                # Caso contrário, a linha contém dados que precisam ser organizados
                if pd.notna(row[0]) and data_atendimento is not None:  # Garante que há uma data antes de processar
                    equipe = row[0]
                    hora = row[9] if pd.notna(row[9]) else "Sem horário"
                    data = row[7] if pd.notna(row[7]) else "Sem horário"
                    
                    # Adiciona os dados estruturados na lista
                    usuarios_data.append([usuario, data_atendimento, equipe, hora, data])

        # Criar DataFrame com as seis colunas desejadas
        df_final = pd.DataFrame(usuarios_data, columns=["Usuário", "Atendimentos", "Equipe", "Hora", "Data"])
        print(df_final["Data"]);
        # Salva em um novo arquivo Excel
        df_final.to_excel("Relatorio_Agrupado.xlsx", index=False)

        messagebox.showinfo("Sucesso", "Dados importados e formatados com sucesso!\nSalvo como 'Relatorio_Agrupado.xlsx'")
    except Exception as e:
        messagebox.showerror("Erro", f"Falha ao importar dados.\nErro: {e}")


def calcular_tempo_medio():
    somaMedia = pd.Timedelta(0)
    contagem = 0

    try:
        # Carrega o DataFrame gerado anteriormente
        df = pd.read_excel("Relatorio_Agrupado.xlsx")

        # Converte "Data" para datetime e "Hora" para datetime, mantendo apenas o horário
        
        df["Data"] = pd.to_datetime(df["Data"], format="%Y-%m-%d")
        df["Hora"] = pd.to_datetime(df["Hora"], format="%H:%M", errors="coerce")

        # Criar lista para armazenar os resultados
        resultados = []

        # Agrupar os atendimentos por usuário e data
        for (usuario, data), grupo in df.groupby(["Usuário", "Data"]):
            # Ordena os atendimentos pelo horário
            grupo = grupo.sort_values(by="Hora")

            # Calcula os intervalos de tempo entre atendimentos consecutivos
            diferencas = grupo["Hora"].diff().dropna()

            if not diferencas.empty:
                # Calcula a média do tempo entre atendimentos
                tempo_medio = diferencas.mean()
                
                # Formatar a média como HH:MM:SS
                # Extrai apenas HH:MM:SS do timedelta
                tempo_medio_str = f"{int(tempo_medio.total_seconds() // 3600):02}:{int((tempo_medio.total_seconds() % 3600) // 60):02}:{int(tempo_medio.total_seconds() % 60):02}"


                # Soma os tempos médios para calcular a média geral depois
                somaMedia += tempo_medio
                contagem += 1

                # Adiciona os dados na lista
                resultados.append([usuario, data.strftime("%Y-%m-%d"), tempo_medio_str])

        # Criar um DataFrame final com os tempos médios
        df_resultado = pd.DataFrame(resultados, columns=["Usuário", "Data Atendimento", "Tempo Médio Entre Atendimentos"])

        # Calcula o tempo médio geral
        tempo_medio_geral = somaMedia / contagem if contagem > 0 else pd.Timedelta(0)
        tempo_medio_geral_str = f"{int(tempo_medio_geral.total_seconds() // 3600):02}:{int((tempo_medio_geral.total_seconds() % 3600) // 60):02}:{int(tempo_medio_geral.total_seconds() % 60):02}"


        # Adiciona o tempo médio geral como a última linha
        df_resultado.loc[len(df_resultado)] = ["Média Geral",contagem, tempo_medio_geral_str]

        # Salvar o resultado em um novo Excel
        df_resultado.to_excel("Tempo_Medio_Atendimentos.xlsx", index=False)

        messagebox.showinfo("Sucesso", "Tempo médio calculado!\nSalvo como 'Tempo_Medio_Atendimentos.xlsx'")
    except Exception as e:
        messagebox.showerror("Erro", f"Falha ao calcular tempo médio.\nErro: {e}")

# Executa a função
# calcular_tempo_medio("Relatorio_Agrupado.xlsx")

# Criar a interface gráfica
root = tk.Tk()
root.title("Analisador de Dados Excel")
root.geometry("400x200")

tk.Label(root, text="Escolha uma ação:", font=("Arial", 12)).pack(pady=10)
tk.Button(root, text="Importar Dados", command=analisar_excel, width=20, height=2).pack(pady=5)
tk.Button(root, text="Calcular Tempo Médio", command=calcular_tempo_medio, width=20, height=2).pack(pady=5)

root.mainloop()