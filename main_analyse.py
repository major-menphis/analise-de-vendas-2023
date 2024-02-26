import streamlit as st
import pandas as pd
import plotly.express as px
import locale

locale.setlocale(locale.LC_MONETARY, 'pt-BR.UTF-8')
# arquivo primeiro semestre
data_file = pd.read_excel('./data/ano 2023/PRODUTOS ST 012023 Á 062023 TODAS AS VENDAS.xlsx')
# arquivo segundo semestre
# data_file = pd.read_excel('./data/ano 2023/PRODUTOS ST 012023 Á 122023 TODAS AS VENDAS.xlsx')
empresa = 'NOME EMPRESARIAL'
data = '1º SEMESTRE DE 2023'
relatorio_excel = f'RELATORIO {empresa} {data}.xlsx'
data_to_save = {}

def salvar_dados_excel(dados_dict):
    writer = pd.ExcelWriter(relatorio_excel, engine='xlsxwriter')
    for sheet_name in dados_dict.keys():
        dados_dict[sheet_name].to_excel(writer, sheet_name=sheet_name, index=False)
    writer.close()

st.set_page_config(layout='wide')
st.title(relatorio_excel.replace('.xlsx', ''))

# identificar e contar os produtos unicos
produtos = data_file['DESCRICAO'].value_counts().reset_index()
data_to_save['quantidade_produtos'] = produtos
# identificar os produtos mais e menos vendidos
quinze_mais_vendidos = produtos.iloc[:15]
quinze_menos_vendidos = produtos.iloc[-15:]
fig_produtos_mais = px.bar(quinze_mais_vendidos, x='count', y='DESCRICAO')
fig_produtos_menos = px.bar(quinze_menos_vendidos, x='count', y='DESCRICAO', barmode='group')

cont_1 = st.container(border=True)
cont_2 = st.container(border=True)
cont_3 = st.container(border=True)
col1, col2 = st.columns([1, 1])
col3, col4 = st.columns([1, 1])
col5, col6 = st.columns([1, 1])
with cont_1:
    col1.subheader('Produtos mais vendidos')
    col1.plotly_chart(fig_produtos_mais, use_container_width=True)

    col2.subheader('Produtos menos vendidos')
    col2.plotly_chart(fig_produtos_menos, use_container_width=True)
    
with cont_2:
    data_file_st_with_icms = data_file[data_file['VLR_ICMS'] != 0]
    col3.write(f'Foram analisadas {data_file.shape[0]} regitros de vendas nesse período.')
    col3.write(f'Das quais {data_file_st_with_icms["DESCRICAO"].value_counts().sum()} foram vendas com destaque de ICMS e classificados como ST.')
    col3.subheader('Lista da quantidade vendida de cada produto com ICMS destacado.')
    col3.table(data_file_st_with_icms['DESCRICAO'].value_counts())

    valor_icms = data_file_st_with_icms['VLR_ICMS'].sum()
    valor_venda = data_file_st_with_icms['TOTAL_VENDA'].sum()
    col4.write(f'Valor total de venda dos produtos com ICMS: {locale.currency(valor_venda, grouping=True)}')
    col4.write(f'Valor total do ICMS destacado indevidamente: {locale.currency(valor_icms, grouping=True)}')
    col4.subheader('Lista do valor de ICMS de cada produto destacado indevidamente.')
    data_file_st_with_icms_descriptions = data_file_st_with_icms[['DESCRICAO', 'VLR_ICMS']].groupby(['DESCRICAO']).sum()
    data_file_st_with_icms_descriptions = data_file_st_with_icms_descriptions.sort_values(by='VLR_ICMS', ascending=False)
    data_file_products_formated = data_file_st_with_icms_descriptions['VLR_ICMS'].apply(lambda x: f"{locale.currency(x, grouping=True)}")
    col4.table(data_file_products_formated)

with cont_3:
    valor_diferenca = data_file_st_with_icms['TOTAL_VENDA'].sum() - data_file_st_with_icms['BASE_ICMS'].sum()
    col5.subheader('Analise de outros indicadores')
    col5.write('\n *Valores arredondados em R$ 0,01 para + ou -')
    col5.write(f'Diferença entre o valor total da venda e base de cálculo do ICMS (Somente do produtos com ICMS destacado): {locale.currency(valor_diferenca, grouping=True)}')
    col5.write('*Valores negativos indicam que a base de cálculo é maior que o total da venda.')

    col5.subheader('Recomendações / Sugestões')
    col5.write('Se faz necessário a verificação dos produtos elencados na lista de destaques indevidos de ICMS, esses produtos não podem sair com ICMS, mesmo que estejam como ST.')
    col5.write('Downloads disponíveis:')
    salvar_dados_excel(data_to_save)
    with open(relatorio_excel, 'rb') as my_file:
        col5.download_button(label='Planilha com a lista de produtos e quantidades',
                             data=my_file,
                             file_name=f'Produtos e quantidade vendidas {empresa} {data}.xlsx',
                             mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')      

    # produtos com valores de venda diferentes da base de calculo do icms
    data_file_icms_dif_venda = data_file_st_with_icms[data_file_st_with_icms['TOTAL_VENDA'] != data_file_st_with_icms['BASE_ICMS']]
    data_file_icms_dif_venda = data_file_icms_dif_venda[['DESCRICAO', 'TOTAL_VENDA', 'BASE_ICMS']].groupby(['DESCRICAO']).sum()
    # converter e comparar os valores de base de calculo e total de venda
    data_file_icms_dif_venda['TOTAL_VENDA'] = data_file_icms_dif_venda['TOTAL_VENDA'].apply(lambda x: f'{locale.currency(x, grouping=True)}')
    data_file_icms_dif_venda['BASE_ICMS'] = data_file_icms_dif_venda['BASE_ICMS'].apply(lambda x: f'{locale.currency(x, grouping=True)}')
    data_file_icms_dif_venda = data_file_icms_dif_venda[data_file_icms_dif_venda['TOTAL_VENDA'] != data_file_icms_dif_venda['BASE_ICMS']]

    col6.subheader('Produtos com diferença entre valor total de venda e base de calculo do ICMS')
    col6.table(data_file_icms_dif_venda)
