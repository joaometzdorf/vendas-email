import pandas as pd
import win32com.client as win32

# Importar a base de dados
tabela_vendas = pd.read_excel("Vendas.xlsx")

# Visualizar a base de dados
pd.set_option("display.max_columns", None)

# Faturamento por loja
faturamento = tabela_vendas[["ID Loja", "Valor Final"]].groupby("ID Loja").sum()

# Quantidade de produtos vendidos por loja
qtd_produtos = tabela_vendas[["ID Loja", "Quantidade"]].groupby("ID Loja").sum()

# Ticket médio por produto em cada loja
ticket_medio = (faturamento["Valor Final"] / qtd_produtos["Quantidade"]).to_frame()
ticket_medio = ticket_medio.rename(columns={0: "Ticket Médio"})

# Enviar um email com o relatório
outlook = win32.Dispatch("outlook.application")
mail = outlook.CreateItem(0)
mail.To = "joaovictorlimametzdorf@gmail.com"
mail.Subject = "Relatório de Vendas por Loja"
mail.HTMLBody = f"""
<p>Prezados,</p>

<p>Segue o relatório de vendas por cada loja.</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade de produtos vendidos:</p>
{qtd_produtos.to_html()}

<p>Ticket médio dos produtos em cada loja:</p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}     

<p>Qualquer dúvida estou à disposição.</p>

<p>Att.,</p>
<p>João Victor</p>
"""

mail.Send()
print("Email enviado")
