import pandas as pd
import tkinter as tk
from datetime import datetime, timedelta
from tkinter import filedialog, messagebox

# Função para importar os dados do Excel
def importar_dados():
    arquivo = filedialog.askopenfilename(title="Selecione o arquivo Excel", filetypes=[("Arquivos Excel", "*.xlsx")])
    if not arquivo:
        return

    try:
        df = pd.read_excel(arquivo)

        # Converter a data e garantir que a coluna "Hora" seja datetime
        df["Data Atendimento"] = pd.to_datetime(df["Data Atendimento"], format="%Y-%m-%d %H:%M:%S").dt.date
        df["Hora"] = pd.to_datetime(df["Hora"], format="%H:%M:%S", errors="coerce").dt.time
        df = df.dropna(subset=["Hora"])

        # Salvar os dados processados
        df.to_excel("Relatorio_Agrupado.xlsx", index=False)
        messagebox.showinfo("Sucesso", "Dados importados e formatados com sucesso!\nSalvo como 'Relatorio_Agrupado.xlsx'")
    except Exception as e:
        messagebox.showerror("Erro", f"Falha ao importar dados.\nErro: {e}")

# Função para calcular tempo médio entre atendimentos
def calcular_tempo_medio():
    try:
        df = pd.read_excel("Relatorio_Agrupado.xlsx")

        if "Hora" not in df.columns:
            messagebox.showerror("Erro", "O arquivo não contém a coluna 'Hora'.")
            return

        df["Data Atendimento"] = pd.to_datetime(df["Data Atendimento"]).dt.date
        df["Hora"] = pd.to_datetime(df["Hora"], format="%H:%M:%S", errors="coerce")
        df = df.dropna(subset=["Hora"])

        resultados = []
        for (usuario, data_atendimento), grupo in df.groupby(["Usuário", "Data Atendimento"]):
            grupo = grupo.sort_values(by="Hora")
            horarios = grupo["Hora"]
            diferencas = horarios.diff().dropna()
            tempo_medio = diferencas.mean()
            resultados.append([usuario, data_atendimento, tempo_medio])

        df_resultado = pd.DataFrame(resultados, columns=["Usuário", "Data Atendimento", "Tempo Médio Entre Atendimentos"])
        df_resultado.to_excel("Tempo_Medio_Atendimentos.xlsx", index=False)

        messagebox.showinfo("Sucesso", "Tempo médio calculado!\nSalvo como 'Tempo_Medio_Atendimentos.xlsx'")
    except Exception as e:
        messagebox.showerror("Erro", f"Falha ao calcular tempo médio.\nErro: {e}")

def calcular_tempo_medio_atualizado():
    contagem = 0
    somaEntrada = timedelta(0)
    somaAtendimento = timedelta(0)
    somaMedia = timedelta(0)
    
    try:
        df = pd.read_excel("Relatorio_atualizado.xlsx")

        if "HR/ MM ATENDIMENTO" not in df.columns:
            messagebox.showerror("Erro", "O arquivo não contém a coluna 'HR/ MM ATENDIMENTO'.")
            return

        # Formata a data e o horário
        df["HR/ MM ATENDIMENTO"] = pd.to_datetime(df["HR/ MM ATENDIMENTO"], format="%H:%M", errors="coerce")
        df = df.dropna(subset=["HR/ MM ATENDIMENTO"])

        resultados = []
        for (usuario, data_atendimento), grupo in df.groupby(["USUÁRIO", "Período  Ref"]):
            contagem += 1
            
            grupo = grupo.sort_values(by="HR/ MM ATENDIMENTO")
            
            # Entrada (primeiro horário do dia)
            entrada = grupo["HR/ MM ATENDIMENTO"].min().time()
            somaEntrada += timedelta(hours=entrada.hour, minutes=entrada.minute)
            
            # Atendimento (último horário do dia)
            atendimento = grupo["HR/ MM ATENDIMENTO"].max().time()
            somaAtendimento += timedelta(hours=atendimento.hour, minutes=atendimento.minute)
            
            # Cálculo do tempo médio
            diferencas = grupo["HR/ MM ATENDIMENTO"].diff().dropna()
            tempo_medio = diferencas.mean() if not diferencas.empty else timedelta(0)
            somaMedia += tempo_medio
            
            # se tiver que excluir o tempo zerado
            if tempo_medio == timedelta(0) :
                contagem -=1
            
            
            # Formatação do tempo médio
            horas, resto = divmod(tempo_medio.total_seconds(), 3600)
            minutos, segundos = divmod(resto, 60)
            tempo_medio_formatado = f"{int(horas):02}:{int(minutos):02}:{int(segundos):02}"
            
            resultados.append([usuario, data_atendimento, entrada, atendimento, tempo_medio_formatado])
        print(contagem)
        # Cálculos das médias evitando divisão por zero
        if contagem > 0:
            media_entrada = somaEntrada / contagem
            media_atendimento = somaAtendimento / contagem
            media_tempo = somaMedia / contagem
        else:
            media_entrada = timedelta(0)
            media_atendimento = timedelta(0)
            media_tempo = timedelta(0)
        


        # Formatar as médias para HH:MM:SS
        def formatar_tempo(td):
            horas, resto = divmod(td.total_seconds(), 3600)
            minutos, segundos = divmod(resto, 60)
            return f"{int(horas):02}:{int(minutos):02}:{int(segundos):02}"

        # Criar linha final com as médias formatadas
        linha_media = pd.DataFrame(
            [
                ["MÉDIA", "-", formatar_tempo(media_entrada), formatar_tempo(media_atendimento), formatar_tempo(media_tempo)]
            ],
            columns=["USUÁRIO", "Período Ref", "Entrada", "Atendimento", "Tempo Médio"]
        )


        # Criar o DataFrame final
        df_resultado = pd.DataFrame(resultados, columns=["USUÁRIO", "Período Ref", "Entrada", "Atendimento", "Tempo Médio"])
        df_resultado = pd.concat([df_resultado, linha_media], ignore_index=True)

        df_resultado.to_excel("Tempo_Medio_Atendimentos_Atualizado.xlsx", index=False)

        messagebox.showinfo("Sucesso", "Tempo médio calculado!\nSalvo como 'Tempo_Medio_Atendimentos_Atualizado.xlsx'")

    except Exception as e:
        import traceback
        erro_detalhado = traceback.format_exc()
        print(erro_detalhado)
        messagebox.showerror("Erro", f"Falha ao calcular tempo médio.\n{e}")



        
# Criar a interface gráfica
root = tk.Tk()
root.title("Analisador de Dados Excel")
root.geometry("400x200")

tk.Label(root, text="Escolha uma ação:", font=("Arial", 12)).pack(pady=10)
tk.Button(root, text="Importar Dados", command=importar_dados, width=20, height=2).pack(pady=5)
tk.Button(root, text="Calcular Tempo Médio", command=calcular_tempo_medio, width=20, height=2).pack(pady=5)
tk.Button(root, text="Calcular Tempo Médio Atualizado", command=calcular_tempo_medio_atualizado, width=30, height=2).pack(pady=5)

root.mainloop()
