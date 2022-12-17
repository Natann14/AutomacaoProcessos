#importando bibliotecas
import pandas as pd
import pathlib #Permite navegar por arquivos e pastas do computador
#import win32com.client as win32
import yagmail

df_emails = pd.read_excel(r'Bases de Dados\Emails.xlsx')
df_lojas = pd.read_csv(r'Bases de dados\Lojas.csv', encoding='latin1', sep=';')
df_vendas = pd.read_excel(r'Bases de dados\Vendas.xlsx')
display(df_emails.head())
display(df_lojas.head())
display(df_vendas.head())

i=0
for email in df_emails['E-mail']:
    df_emails.loc[i, 'E-mail'] = '21052001+{}@gmail.com'.format(i)
    i+=1

display(df_emails)

novo_vendas =  df_vendas.merge(df_lojas, on='ID Loja')
display(novo_vendas)

dicionario_lojas = {}

for loja in df_lojas['Loja']:
    dicionario_lojas[loja] = novo_vendas.loc[novo_vendas['Loja']==loja, :]
    
display(dicionario_lojas['Shopping União de Osasco'])

dia_indicador = novo_vendas['Data'].max()
print(dia_indicador)

caminhoarquivo = open('path', 'r')

#Verificar se a pasta com o nome da loja existe, caso não exista, criar a pasta
caminho = pathlib.Path(fr'{caminhoarquivo}')

#A princípio isso é uma lista vazia, pois na pasta não havia nada
lista_pastas = [item.name for item in caminho.iterdir()]

for loja in dicionario_lojas:
    if loja not in lista_pastas:
        pasta = caminho / loja #concatena o caminho do arquivo com o nome da pasta
        pasta.mkdir()
    
    #caso haja pasta com nome da loja o programa apenas irá salvar o arquivo 
    
    extensao = f'{dia_indicador.month}_{dia_indicador.day}_{loja}.xlsx' #definindo nome e extensão do arquivo
    dicionario_lojas[loja].to_excel(caminho / loja / extensao) #rf é a mistura de "rowstring" com "fstring"
    lista_pastas.append(loja)
        
#print(lista_pastas)

loja = 'Norte Shopping'
dicionario_vendas = dicionario_lojas[loja]
#criando Dataframe para vendas do dia
dic_vendasdia = dicionario_vendas.loc[dicionario_vendas['Data']==dia_indicador, :]


#faturamento
faturamento_ano = dicionario_vendas['Valor Final'].sum()
print(faturamento_total)

faturamento_dia = dic_vendasdia['Valor Final'].sum()
print(faturamento_dia)

#diversidade de produtos 
qtd_produtos_ano = len(dicionario_vendas['Produto'].unique())
print(qtd_produtos_ano)

qtd_produtos_dia = len(dic_vendasdia['Produto'].unique())
print(qtd_produtos_dia)

#ticket médio ano
valor_venda = dicionario_vendas.groupby('Código Venda').sum() #O 'groupby' agrupa o df pela coluna'Código venda' e o '.sum()' soma
display(valor_venda)

ticket_medio_ano = valor_venda['Valor Final'].mean() #pegando a média dos valores da coluna valor final
print(ticket_medio_ano)

#ticket medio dia
valor_venda_dia = dic_vendasdia.groupby('Código Venda').sum()
ticket_medio_dia = valor_venda_dia['Valor Final'].mean() #pegando a média dos valores da coluna valor final
print(ticket_medio_dia)

faturamento_meta_ano = 1650000
faturamento_meta_dia = 1000
meta_qtdeprodutos_ano = 120 
meta_qtdeprodutos_dia = 4
ticket_meta_ano = 500
ticket_meta_dia = 500

#Exibe o 'caminho' do local onde o programa está rodando
print(pathlib.Path.cwd())

print('-'*50)

print(df_emails.loc[df_emails['Loja']==loja, 'E-mail'].values[0])

outlook = win32.Dispatch('outlook.application')

mail = outlook.CreateItem(0)

#definindo uma variavel para os nomes de quem vou enviar os emails OBS.: o nome vem do dataframe
nome = df_emails.loc[df_emails['Loja']==loja, 'Gerente'].values[0]

#Pegando o email no dataframe, o metodo '.To' define para quem quero enviar o email. Já faz parte do código de enviar o email 
mail.To = df_emails.loc[df_emails['Loja']==loja, 'E-mail'].values[0]

##método '.CC' significa copia
#mail.CC = 'email@gmail.com'

##método '.BCC' significa copia oculta
#mail.BCC = 'email@gmail.com'

#Assunto
mail.Subject = f'One Page Dia: {dia_indicador.day}/{_indicador.month} - Loja: {loja}'

#corpo do email OBS.: Pode ser em HTML
#mail.Body = 'Texto do E-mail'

if faturamento_dia >= faturamento_meta_dia:
    cor_fat_dia = 'green'
else:
    cor_fat_dia = 'red'
if faturamento_ano >= faturamento_meta_ano:
    cor_fat_ano = 'green'
else:
    cor_fat_ano = 'red'
if qtd_produtos_dia >= meta_qtdeprodutos_dia:
    cor_qtde_dia = 'green'
else:
    cor_qtde_dia = 'red'
if qtd_produtos_ano >= meta_qtdeprodutos_ano:
    cor_qtde_ano = 'green'
else:
    cor_qtde_ano = 'red'
if ticket_medio_dia >= ticket_meta_dia:
    cor_ticket_dia = 'green'
else:
    cor_ticket_dia = 'red'
if ticket_medio_ano >= ticket_meta_ano:
    cor_ticket_ano = 'green'
else:
    cor_ticket_ano = 'red'

texto = f'''
<p>Bom dia, {nome}</p>

<p>O resultado de ontem <strong>({dia_indicador.day}/{dia_indicador.month})</strong> da <strong>loja {loja}</strong>:</p>

<table>
  <tr>
    <th>Indicador</th>
    <th>Valor Dia</th>
    <th>Meta Dia</th>
    <th>Cenário Dia</th>
  </tr>
  <tr>
    <td>Faturamento</td>
    <td style="text-align: center">R${faturamento_dia:.2f}</td>
    <td style="text-align: center">R${faturamento_meta_dia:.2f}</td>
    <td style="text-align: center"><font color="{cor_fat_dia}">◙</font></td>
  </tr>
  <tr>
    <td>Diversidade de Produtos</td>
    <td style="text-align: center">{qtd_produtos_dia}</td>
    <td style="text-align: center">{meta_qtdeprodutos_dia}</td>
    <td style="text-align: center"><font color="{cor_qtde_dia}">◙</font></td>
  </tr>
  <tr>
    <td>Ticket Médio</td>
    <td style="text-align: center">R${ticket_medio_dia:.2f}</td>
        <td style="text-align: center">R${ticket_meta_dia:.2f}</td>
    <td style="text-align: center"><font color="{cor_ticket_dia}">◙</font></td>
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
    <td style="text-align: center">R${faturamento_ano:.2f}</td>
    <td style="text-align: center">R${faturamento_meta_ano:.2f}</td>
    <td style="text-align: center"><font color="{cor_fat_ano}">◙</font></td>
  </tr>
  <tr>
    <td>Diversidade de Produtos</td>
    <td style="text-align: center">{qtd_produtos_ano}</td>
    <td style="text-align: center">{meta_qtdeprodutos_ano}</td>
    <td style="text-align: center"><font color="{cor_qtde_ano}">◙</font></td>
  </tr>
  <tr>
    <td>Ticket Médio</td>
    <td style="text-align: center">R${ticket_medio_ano:.2f}</td>
        <td style="text-align: center">R${ticket_meta_ano:.2f}</td>
    <td style="text-align: center"><font color="{cor_ticket_ano}">◙</font></td>
  </tr>
</table>

<p>Segue em anexo a planilha com todos os dados para mais detalhes</p>

<p>Qualquer dúvida, estou à disposição</p>

<p>Att., Natan</p>

'''

mail.HTMLBody = texto

# Anexos (pode colocar quantos quiser):
attachment  = caminho / loja / extensao
mail.Attachments.Add(str(attachment))

mail.Send()