#Calcular OnePage de 25 lojas e enviar formatado no E-mail ao gerente de cada loja correspondente.

#                   Tabela modelo

#                valor dia   meta dia   cenário dia
# faturamento       0           0              v
# diversidade       0           0              f
# ticket medio      0           0              v

#                valor ano   meta ano   cenário ano
# faturamento       0           0              v
# diversidade       0           0              f
# ticket medio      0           0              v



#Leitura e formatação de DataFrames, planilhas, csv..
import pandas as pd

#Criar pastas e ler diretórios
import pathlib
from pathlib import Path

#Leitura dos arquivos da base de dados
email_df = pd.read_excel(r'C:\Users\leand\OneDrive\Documentos\hashtag\f_projetos\projeto1\Projeto AutomacaoIndicadores\Bases de Dados\Emails.xlsx')
vendas_df = pd.read_excel(r'C:\Users\leand\OneDrive\Documentos\hashtag\f_projetos\projeto1\Projeto AutomacaoIndicadores\Bases de Dados\Vendas.xlsx')
lojas_df = pd.read_csv(r'C:\Users\leand\OneDrive\Documentos\hashtag\f_projetos\projeto1\Projeto AutomacaoIndicadores\Bases de Dados\Lojas.csv', sep=';', encoding='utf-8')

# Função para corrigir os caracteres corrompidos
def corrigir_caracteres(texto):
    substituicoes = {
        'Uni�o': 'União',
        'Ribeir�o': 'Ribeirão',
        '�guas': 'Águas',
        'Uberl�ndia': 'Uberlândia',
        'Ribeir�o': 'Ribeirão',
        '��ID Loja': 'ID Loja'  # Adicionando a correção do cabeçalho
    }
    for corrompido, correto in substituicoes.items():
        texto = texto.replace(corrompido, correto)
    return texto

# Corrigindo o cabeçalho do DataFrame
lojas_df.columns = [corrigir_caracteres(col) for col in lojas_df.columns]

# Aplica a função de correção na coluna 'Loja'
lojas_df['Loja'] = lojas_df['Loja'].apply(corrigir_caracteres)

# Formatação necessária
vendas_df = vendas_df.merge(lojas_df, on='ID Loja')

#Separando arquivos da planilha em DataFrames
dicio = {}
for loja in lojas_df['Loja']:
    dicio[loja] = vendas_df.loc[vendas_df['Loja']==loja, : ]

# Dia do indicador
dia = vendas_df['Data'].max()

#Criação de lista com nomes das pastas que serão criadas
caminho = pathlib.Path(r'C:\Users\leand\OneDrive\Documentos\hashtag\f_projetos\projeto1\Projeto AutomacaoIndicadores\Backup Arquivos Lojas')
arquivos = caminho.iterdir()
lista_arq = [arquivo.name for arquivo in arquivos]

# - Metas da empresa
meta_fat_dia = 1000
meta_fat_ano = 1650000
meta_diver_dia = 4
meta_diver_ano = 120
meta_media_dia = 500
meta_media_ano = 500

#Criação das pastas com nomes das lojas
for nome in lojas_df['Loja']:
    if nome not in lista_arq:
        pasta = caminho/nome
        pasta.mkdir()

#Convertendo DataFrames em arquivos excel separados
    nome_aquivo = ("{}_{}_{}.xlsx").format(dia.month, dia.day, nome)
    caminho_arq = caminho / nome / nome_aquivo
    dicio[nome].to_excel(caminho_arq)

# - Faturamento
for lojass in dicio:

    vendas_ano = dicio[lojass]  #Definindo DataFrames na variável vendas_ano
    vendas_dia = vendas_ano.loc[vendas_ano['Data']==dia, :]  #Definindo vendas do dia

    #Criando cópia do DataFrame e excluindo coluna 'Data'
    copia = vendas_ano.copy()  
    auxiliar = copia[['Código Venda','ID Loja', 'Produto', 'Quantidade', 'Valor Unitário', 'Valor Final', 'Loja']] 

# - Faturamentos
    fat_ano = vendas_ano['Valor Final'].sum()
    fat_dia = vendas_dia['Valor Final'].sum()
    
# - Diversidade de Produtos Vendidos
    diversi_ano =len(vendas_ano['Produto'].unique())
    diversi_dia = len(vendas_dia['Produto'].unique())

# - Ticket Médio por Venda 
    valor_ano = auxiliar.groupby('Código Venda').sum()
    media_ano = valor_ano['Valor Final'].mean()

    valor_dia = auxiliar.groupby('Código Venda').sum()
    media_dia = valor_dia['Valor Final'].mean()

#Definindo as cores da tabela final do email
    if fat_ano >= meta_fat_ano:
       cor_fat_ano = 'green'
    else:
        cor_fat_ano = 'red'

    if fat_dia >= meta_fat_dia:
        cor_fat_dia = 'green'
    else:
        cor_fat_dia = 'red'

    if diversi_ano >= meta_diver_ano:
        cor_diver_ano = 'green'
    else:
        cor_diver_ano = 'red'

    if diversi_dia >= meta_diver_dia:
      cor_diver_dia = 'green'
    else:
        cor_diver_dia = 'red'
    
    if media_ano >= meta_media_ano:
        cor_media_ano = 'green'
    else:
        cor_media_ano = 'red'

    if media_dia >= meta_media_dia:
        cor_media_dia = 'green'
    else:
        cor_media_dia = 'red'

# - Enviar Email
    import win32com.client as win32

    email = email_df.loc[email_df['Loja']==lojass, 'E-mail'].values[0]  #Definindo variável email

    outlook = win32.Dispatch('outlook.application')  #Código padrão
    mail = outlook.CreateItem(0)  #Código padrão
    mail.To = email  #Definindo destinatário
    mail.Subject = 'OnePage Dia {}/{}/{} - Loja {}'.format(dia.day, dia.month, dia.year, lojass)  #Definindo cabeçalho do email
    mail.HTMLBody = f'''  #Criando tabelas no corpo do email
<table>
  <tr>
    <th>Indicador</th>
    <th>Valor Dia</th>
    <th>Meta Dia</th>
    <th>Cenário Dia</th>
  </tr>
  <tr>
    <td>Faturamento</td>
    <td style="text-align: center">R${fat_dia:.2f}</td>
    <td style="text-align: center">R${meta_fat_dia}</td>
     <td style="text-align: center"><font color="{cor_fat_dia}">◙</font></td>
  </tr>
  <tr>
    <td>Diversidade de Produtos</td>
    <td style="text-align: center">R${diversi_dia:.2f}</td>
    <td style="text-align: center">R${meta_diver_dia}</td>
     <td style="text-align: center"><font color="{cor_diver_dia}">◙</font></td>
  </tr>
  <tr>
    <td>Ticket Médio</td>
    <td style="text-align: center">R${media_dia:.2f}</td>
    <td style="text-align: center">R${meta_media_dia}</td>
     <td style="text-align: center"><font color="{cor_media_dia}">◙</font></td>
  </tr>
</table>
<br>
<table>
  <tr>
    <th>Indicador</th>
    <th>Valor Ano</th>
    <th>Meta Ano</th>
    <th>Cenário Ano</th>
  </tr>
  <tr>
    <td>Faturamento</td>
    <td style="text-align: center">R${fat_ano}</td>
    <td style="text-align: center">R${meta_fat_ano}</td>
     <td style="text-align: center"><font color="{cor_fat_ano}">◙</font></td>
  </tr>
  <tr>
    <td>Diversidade de Produtos</td>
    <td style="text-align: center">R${diversi_ano}</td>
    <td style="text-align: center">R${meta_diver_ano}</td>
     <td style="text-align: center"><font color="{cor_diver_ano}">◙</font></td>
  </tr>
  <tr>
    <td>Ticket Médio</td>
    <td style="text-align: center">R${media_ano:.2f}</td>
    <td style="text-align: center">R${meta_media_ano}</td>
     <td style="text-align: center"><font color="{cor_media_ano}">◙</font></td>
  </tr>
</table>
'''
    mail.Send()  #Enviar email
print('Concluido')