# Fazer os imports necessários
import yfinance as yf
import pandas as pd
import matplotlib.pyplot as plt
import mplcyberpunk
import win32com.client as win32


# Passo 2 - Pegar as cotações históricas
# Escolhendo as cotações que eu quero
acoes = ["^BVSP", "^GSPC", "BRL=X"]
# Baixando-as
dados_mercado = yf.download(acoes, period = "6mo")
# Filtrando dados
dados_mercado = dados_mercado["Adj Close"]

# Passo 3 - Tratar dados coletados.
# Excluindo resultados com NaN
dados_mercado = dados_mercado.dropna()
# Nomeando colunas
dados_mercado.columns = ["DOLAR", "IBOVESPA", "S&P500"]

# Passo 4 - Criar gráficos de performance.
# Estilizando o gráfico
plt.style.use("cyberpunk")
# Criando o gráfico 
plt.plot(dados_mercado["IBOVESPA"])
# Entitulando gráfico
plt.title("IBOVESPA")
# Salvando imagem do gráfico na máquina
plt.savefig("ibovespa.png")

plt.plot(dados_mercado["DOLAR"])
plt.title("DOLAR")

plt.savefig("dolar.png")

plt.plot(dados_mercado["S&P500"])
plt.title("S&P500")

plt.savefig("sp500.png")

# Passo 5 - Calcular retornos diários.
retornos_diarios = dados_mercado.pct_change()
# Pegando os dados mais recentes
retorno_dolar = retornos_diarios["DOLAR"].iloc[-1]
retorno_ibovespa = retornos_diarios["IBOVESPA"].iloc[-1]
retorno_sp = retornos_diarios["S&P500"].iloc[-1]

retorno_dolar = str(round(retorno_dolar * 100, 2)) + "%"
retorno_ibovespa = str(round(retorno_ibovespa * 100, 2)) + "%"
retorno_sp = str(round(retorno_sp * 100, 2)) + "%"

#Configurar enviar email 

outlook = win32.Dispatch("outlook.application") 

email = outlook.CreateItem(0)

email.To = "lucasvaz278@gmail.com"
email.Subject = "Relatório de Mercado"
email.Body = f'''Prezado diretor, segue o relatório de mercado com gráficos feito automaticamente com python :

* O Ibovespa teve o retorno de {retorno_ibovespa}.
* O Dólar teve o retorno de {retorno_dolar}.
* O S&P500 teve o retorno de {retorno_sp}.

Segue em anexo a peformance dos ativos nos últimos 6 meses.

Ass,
O melhor programador do mundo, Lucas


'''

anexo_ibovespa = r"C:\Users\PC\Desktop\Estudos programação\projeto_relatorio_mercado\ibovespa.png"
anexo_dolar = r"C:\Users\PC\Desktop\Estudos programação\projeto_relatorio_mercado\dolar.png"
anexo_sp = r"C:\Users\PC\Desktop\Estudos programação\projeto_relatorio_mercado\sp500.png"

email.Attachments.Add(anexo_ibovespa)
email.Attachments.Add(anexo_dolar)
email.Attachments.Add(anexo_sp)

email.Send()
print('Enviado com sucesso!')