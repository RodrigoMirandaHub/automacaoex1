import pandas as pd
import matplotlib.pyplot as plt


arquivo = 'dados/vendas.xlsx'

tabela_vendas = pd.read_excel(arquivo)
print(tabela_vendas.head())

#total de vendas por vendedor
vendas_por_vendedor = tabela_vendas.groupby("Vendedor") ['Total'].sum().sort_values(ascending=False)
print("\nTotal de vendas por vendedor:")
print(vendas_por_vendedor)

#total de vendas por regiao 
vendas_por_regiao = tabela_vendas.groupby("Região")['Total'].sum().sort_values(ascending=False)
print("\nTotal e vendas por regiao:")
print(vendas_por_regiao)

#produto mais vendido (por quantidade)
mais_vendidos = tabela_vendas.groupby("Produto")['Quantidade'].sum().sort_values(ascending=False)
print("\nProdutos mais vendidos (por quantidade):")
print(mais_vendidos)



# Criando um novo arquivo Excel com os resumos
with pd.ExcelWriter("output/relatorio_final.xlsx", engine="openpyxl") as writer:
    tabela_vendas.to_excel(writer, sheet_name="Base de Vendas", index=False)
    vendas_por_vendedor.to_excel(writer, sheet_name="Vendas por Vendedor")
    vendas_por_regiao.to_excel(writer, sheet_name="Vendas por Região")
    mais_vendidos.to_excel(writer, sheet_name="Mais Vendidos")



# Gráfico de barras: vendas por vendedor
plt.figure(figsize=(10, 6))
vendas_por_vendedor.plot(kind='bar', color='skyblue')
plt.title("Vendas por Vendedor")
plt.ylabel("Total em R$")
plt.xlabel("Vendedor")
plt.tight_layout()
plt.savefig("output/vendas_por_vendedor.png")
plt.close()

# Gráfico de pizza: vendas por região
plt.figure(figsize=(8, 8))
vendas_por_regiao.plot(kind='pie', autopct='%1.1f%%', startangle=90)
plt.title("Distribuição de Vendas por Região")
plt.ylabel("")  # remove label do eixo y
plt.tight_layout()
plt.savefig("output/vendas_por_regiao.png")
plt.close()
