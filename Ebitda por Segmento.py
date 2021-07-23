import pandas as pd
from datetime import datetime

pd.options.mode.chained_assignment = None  # default='warn'

desired_width = 320
pd.set_option('display.width', desired_width)
pd.set_option('display.max_columns', 10)


def formatar_dados(dados_inicio):
    dados_inicio_1 = pd.DataFrame(dados_inicio)
    dados_inicio_1['Conta do Razão'] = dados_inicio_1['Conta do Razão'].astype(str)
    dados_inicio_3 = dados_inicio_1[dados_inicio_1['Conta do Razão'].str.contains('6151299020')]
    remover = ['6150', '615129', '6152', '6153', '6154', '6155', '6156', '61604', '6162', '6163',
               '6164', '6165', '6160299010', '6151151020', '6151151010', '6151151020', '6151152010',
               '6160351010', '6160351011']
    dados_inicio_1 = dados_inicio_1[~dados_inicio_1['Conta do Razão'].str.contains('|'.join(remover))]
    dados_inicio_1 = dados_inicio_1[~dados_inicio_1['Elemento PEP'].str.contains('RSG', na=False)]
    dados_inicio_1 = dados_inicio_1[dados_inicio_1['Data de lançamento'].notnull()]
    dados_inicio_2 = dados_inicio_1[dados_inicio_1['Centro custo'].isnull()]
    dados_inicio_2 = dados_inicio_2[~dados_inicio_2['Conta do Razão'].str.contains('6151124130')]
    dados_inicio_2['Ordem'] = dados_inicio_1['Ordem'].astype(str)
    dados_inicio_2.loc[dados_inicio_2['Ordem'].str.contains('100284'), 'Centro custo'] = '11400'
    dados_inicio_2.loc[dados_inicio_2['Ordem'].str.contains('100342'), 'Centro custo'] = '11500'
    dados_inicio_2.loc[dados_inicio_2['Ordem'].str.contains('100283'), 'Centro custo'] = '11300'
    dados_inicio_2.loc[dados_inicio_2['Ordem'].str.contains('100282'), 'Centro custo'] = '11100'
    dados_inicio_2.loc[dados_inicio_2['Conta do Razão'].str.contains('61601'), 'Centro custo'] = '11420'
    dados_inicio_2.loc[dados_inicio_2['Conta do Razão'].str.contains('615'), 'Centro custo'] = '11440'
    dados_inicio_1 = dados_inicio_1[dados_inicio_1['Centro custo'] > 11000.0]
    dados_inicio_1 = dados_inicio_1[dados_inicio_1['Centro custo'] < 14000.0]
    dados_inicio_1 = pd.concat([dados_inicio_1, dados_inicio_2, dados_inicio_3])
    del dados_inicio_1['Ordem']
    del dados_inicio_1['Elemento PEP']
    del dados_inicio_1['Atribuição']
    dados_inicio_1['Centro custo'] = pd.to_numeric(dados_inicio_1['Centro custo'], errors='coerce')
    return dados_inicio_1


dados_inicio_1 = formatar_dados(pd.read_excel('Despesas06.xlsx'))


def completar_dados(dados_2, criterios):
    busca_sigla = pd.DataFrame(dados_2)
    busca_sigla['complemento'] = ['.0' for l in busca_sigla['Conta do Razão']]
    busca_sigla['Conta do Razão'] = busca_sigla['Conta do Razão'].astype(str) + busca_sigla['complemento']
    dados_a_completar = pd.merge(dados_inicio_1, busca_sigla[['Centro custo', 'Sigla']], on=['Centro custo'],
                                 how='left')
    dados_a_completar['busca'] = dados_a_completar['Centro custo'].astype(str) + dados_a_completar[
        'Conta do Razão'].astype(str)
    dados_a_completar = pd.merge(dados_a_completar, busca_sigla[['Conta do Razão', 'nome_conta']],
                                 on=['Conta do Razão'], how='left')
    criterios = criterios[criterios['CONTAS CONTÁBEIS'].notnull()]
    criterios['busca'] = criterios['Unnamed: 4'].astype(str) + criterios['CONTAS CONTÁBEIS'].astype(str)
    dados_a_completar = pd.merge(dados_a_completar, criterios[['busca', 'Cód. ']], on=['busca'], how='left')
    check = dados_a_completar[dados_a_completar['Cód. '].isnull()]
    check.to_excel('erros.xlsx')
    print(check)
    dados_a_completar = dados_a_completar[dados_a_completar['Cód. '].notnull()]
    del dados_a_completar['busca']
    del dados_a_completar['Tipo de documento']
    return dados_a_completar


abas = [5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22,
        23, 24, 25, 26, 27, 28, 29, 30, 31, 32]

dados_a_completar = completar_dados(pd.read_excel('G:\GECOT\Despesas por Segmento\DRIVERS de rateio por segmento.xlsx',
                                                  sheet_name='SIGLAS'),
                                    pd.concat(pd.read_excel('G:\GECOT\Despesas por Segmento\DRIVERS de '
                                    'rateio por segmento.xlsx',
                                    sheet_name=abas, usecols=[4, 5, 6, 7, 8]), ignore_index=True))


def definir_rateios(rateio):
    driver_rateio = pd.DataFrame(rateio)
    mes = input('Digite o mês (mm-aaaa): ')
    global mes_arquivo
    mes_arquivo = mes
    mes = datetime.strptime(mes, '%m-%Y')

    driver_rateio = driver_rateio[driver_rateio['Data'].isin([mes])]
    driver_rateio.reset_index(inplace=True)
    del driver_rateio['index']
    return driver_rateio


driver_rateio = definir_rateios(pd.read_excel('G:\GECOT\Despesas por Segmento\DRIVERS de rateio por segmento.xlsx',
                                              sheet_name='SEGMENTOS'))

def rateio_por_segmento(dados_a_completar, driver_rateio):
    for index, row in dados_a_completar.iterrows():
        for i in range(1, 5):
            if row['Cód. '] == i:
                if i == 1:
                    dados_a_completar.loc[index, 'Residencial'] = row['Montante em moeda interna']
                elif i == 2:
                    dados_a_completar.loc[index, 'Comercial'] = row['Montante em moeda interna']
                elif i == 3:
                    dados_a_completar.loc[index, 'Industrial'] = row['Montante em moeda interna']
                elif i == 4:
                    dados_a_completar.loc[index, 'Gás Natural Veicular - GNV'] = row['Montante em moeda interna']
            else:
                if row['Cód. '] == 7:
                    dados_a_completar.loc[index, 'Residencial'] = row['Montante em moeda interna'] * (driver_rateio.loc[0][9] + driver_rateio.loc[1][9])
                    dados_a_completar.loc[index, 'Comercial'] = row['Montante em moeda interna'] * driver_rateio.loc[2][9]
                    dados_a_completar.loc[index, 'Gás Natural Veicular - GNV'] = row['Montante em moeda interna'] * driver_rateio.loc[4][9]
                    dados_a_completar.loc[index, 'Gás Natural - Frotas'] = row['Montante em moeda interna'] * driver_rateio.loc[6][9]
                elif row['Cód. '] == 8:
                    dados_a_completar.loc[index, 'Residencial'] = row['Montante em moeda interna'] * (
                                driver_rateio.loc[0][7] + driver_rateio.loc[1][7])
                    dados_a_completar.loc[index, 'Comercial'] = row['Montante em moeda interna'] * driver_rateio.loc[2][7]
                    dados_a_completar.loc[index, 'Industrial'] = row['Montante em moeda interna'] * (driver_rateio.loc[3][7] + driver_rateio.loc[10][7])
                    dados_a_completar.loc[index, 'Gás Natural Veicular - GNV'] = row['Montante em moeda interna'] * driver_rateio.loc[4][7]
                    dados_a_completar.loc[index, 'Gás Natural - Frotas'] = row['Montante em moeda interna'] * driver_rateio.loc[6][7]
                elif row['Cód. '] == 9:
                    dados_a_completar.loc[index, 'Residencial'] = row['Montante em moeda interna'] * (
                            driver_rateio.loc[0][3] + driver_rateio.loc[1][3])
                    dados_a_completar.loc[index, 'Comercial'] = row['Montante em moeda interna'] * driver_rateio.loc[2][3]
                    dados_a_completar.loc[index, 'Industrial'] = row['Montante em moeda interna'] * (
                                driver_rateio.loc[3][3] + driver_rateio.loc[10][3])
                    dados_a_completar.loc[index, 'Gás Natural Veicular - GNV'] = row['Montante em moeda interna'] * driver_rateio.loc[4][3]
                    dados_a_completar.loc[index, 'Gás Natural - Frotas'] = row['Montante em moeda interna'] * driver_rateio.loc[6][3]
                elif row['Cód. '] == 10:
                    dados_a_completar.loc[index, 'Industrial'] = row['Montante em moeda interna'] * (
                            driver_rateio.loc[3][11] + driver_rateio.loc[10][11])
                    dados_a_completar.loc[index, 'Gás Natural Veicular - GNV'] = row['Montante em moeda interna'] * driver_rateio.loc[4][11]
                    dados_a_completar.loc[index, 'Gás Natural - Frotas'] = row['Montante em moeda interna'] * driver_rateio.loc[6][11]
                elif row['Cód. '] == 13 or row['Cód. '] == 12:
                    dados_a_completar.loc[index, 'Residencial'] = row['Montante em moeda interna'] * (
                            driver_rateio.loc[0][5] + driver_rateio.loc[1][5])
                    dados_a_completar.loc[index, 'Comercial'] = row['Montante em moeda interna'] * driver_rateio.loc[2][5]
                    dados_a_completar.loc[index, 'Industrial'] = row['Montante em moeda interna'] * (
                            driver_rateio.loc[3][5] + driver_rateio.loc[10][5])
                    dados_a_completar.loc[index, 'Gás Natural Veicular - GNV'] = row['Montante em moeda interna'] * driver_rateio.loc[4][5]
                    dados_a_completar.loc[index, 'Gás Natural - Frotas'] = row['Montante em moeda interna'] * driver_rateio.loc[6][5]
                elif row['Cód. '] == 14:
                    dados_a_completar.loc[index, 'Residencial'] = row['Montante em moeda interna'] * (
                                driver_rateio.loc[0][13] + driver_rateio.loc[1][13])
                    dados_a_completar.loc[index, 'Comercial'] = row['Montante em moeda interna'] * driver_rateio.loc[2][13]

    tabela_pronta = pd.melt(dados_a_completar,
                          id_vars=['Data de lançamento', 'Data do documento', 'Montante em moeda interna',
                                   'Nº documento', 'Texto', 'Conta do Razão', 'Centro custo', 'Sigla', 'nome_conta',
                                   'Cód. '], var_name='Segmento', value_name='Valor')
    for index, row in tabela_pronta.iterrows():
        if row['Sigla'] != 'CUSTO' or row['Sigla'] != 'DESPESA':
            tabela_pronta.loc[index, 'Montante em moeda interna'] = row['Valor']
    del tabela_pronta['Valor']
    return tabela_pronta

tabela_pronta = rateio_por_segmento(dados_a_completar, driver_rateio)

def unir_com_balancete(tabela_pronta, balancete, driver):
    balancete = pd.DataFrame(balancete)
    balancete.drop(['Saldo Inicial', 'Movimentação a Débito', 'Movimentação a Crédito', 'Saldo Acumulado'], axis=1,
                   inplace=True)
    balancete['Conta do Razão'] = pd.to_numeric(balancete['Conta do Razão'], errors='coerce')
    balancete['Conta do Razão'] = balancete[balancete['Conta do Razão'] > 6000000000]
    for index, row in balancete.iterrows():
        if row['Conta do Razão'] < 6130000000:
            balancete.loc[index, 'Sigla'] = 'RECEITA'
        else:
            balancete.loc[index, 'Sigla'] = 'CUSTO'

    selecao_bal = ['611', '612', '615013', '615023', '615029', '615213', '615313', '615413', '615613', '6152192001',
                   '6153192001',
                   '6154192001', '6156192001', '6150192001', '6150292001']
    balancete['Conta do Razão'] = balancete['Conta do Razão'].astype(str)
    balancete['Segmento'] = balancete['Texto Conta do Razão'].str.split().str[-1]

    balancete = balancete[balancete['Conta do Razão'].str.contains('|'.join(selecao_bal))]
    texto = {'GNC': 'Industrial', 'RESIDENCIAL': 'Residencial',
             'INDUSTRIAL': 'Industrial', 'GNV': 'Gás Natural Veicular - GNV', 'COMERCIAL': 'Comercial',
             'PRIMA': 'Industrial', 'MAT.PRIMA': 'Industrial', '-MAP.PRIMA': 'Industrial'}
    for i, j in texto.items():
        balancete['Segmento'] = balancete['Segmento'].replace(i, j)
    balancete.rename(columns={'Total Movimentação': 'Montante em moeda interna', 'Texto Conta do Razão': 'nome_conta'}, inplace=True)
    balancete['Data de lançamento'] = datetime.strptime(mes_arquivo, '%m-%Y')
    balancete['Data do documento'] = balancete['Data de lançamento']
    balancete['Centro custo'] = balancete['Sigla']
    balancete_rateado = pd.DataFrame()
    for index, row in balancete.iterrows():
        if row['Segmento'] == 'Gás Natural Veicular - GNV':
            balancete.loc[index, 'Montante em moeda interna'] = row['Montante em moeda interna'] * \
                                                                driver.loc[4][2] / (driver.loc[4][2] + driver.loc[6][2])
            balancete.loc[[index], ['Segmento']] = 'Gás Natural Veicular - GNV'
            balancete_rateado = balancete_rateado.append(balancete.loc[[index]])
            balancete.loc[index, 'Montante em moeda interna'] = row['Montante em moeda interna'] * \
                                                                driver.loc[6][2] / (driver.loc[4][2] + driver.loc[6][2])
            balancete.loc[[index], ['Segmento']] = 'Gás Natural - Frotas'
            balancete_rateado = balancete_rateado.append(balancete.loc[[index]])
        else:
            balancete_rateado = balancete_rateado.append(balancete.loc[[index]])
    tabela_pronta = pd.concat([tabela_pronta, balancete_rateado]).fillna(0)
    tabela_pronta['Montante em moeda interna'] = tabela_pronta['Montante em moeda interna'] * -1
    return tabela_pronta

tabela_consolidada = unir_com_balancete(tabela_pronta, pd.read_excel('balancete06.xlsx'), driver_rateio)


def resumir_segmento(consolidado):
    tabela = [[], []]
    for s in consolidado['Segmento'].unique():
        tabela[0].append(s)
        tabela[1].append(consolidado[consolidado['Segmento'] == s]['Montante em moeda interna'].sum())
    datanova = pd.DataFrame({'Segmento': tabela[0], 'Montante': tabela[1]})
    datanova.loc['TOTAL GERAL'] = datanova.iloc[:, 1:].sum(axis=0)
    return datanova


por_segmento = resumir_segmento(tabela_consolidada)


def formatar_consolidado(item, segmento):
    writer = pd.ExcelWriter('G:\GECOT\Despesas por Segmento\Despesas por Segmento ' + mes_arquivo + '.xlsx', engine='xlsxwriter')
    item = pd.DataFrame(item)
    item.rename(columns={'Montante em moeda interna': 'Montante'}, inplace=True)
    item['Data de lançamento'] = item['Data de lançamento'].dt.date
    item['Data do documento'] = item['Data do documento'].dt.date
    # item.loc['TOTAL GERAL'] = item.loc[:, 'Montante'].sum(axis=0)
    item.to_excel(writer, sheet_name='Geral', index=False)
    segmento.to_excel(writer, sheet_name='Resumo', index=False)

    workbook = writer.book
    worksheet = writer.sheets['Geral']
    worksheet_segmento = writer.sheets['Resumo']

    format_numero = workbook.add_format({'num_format': '#,##0.00'})
    format_texto = workbook.add_format({'num_format': '0', 'align': 'center'})
    format_texto2 = workbook.add_format({'num_format': '0', 'align': 'left'})
    format_data = workbook.add_format({'num_format': 'dd/mm/yyyy', 'align': 'center'})

    worksheet.set_column('A:A', 17, format_data)
    worksheet.set_column('B:B', 17, format_data)
    worksheet.set_column('C:C', 15, format_numero)
    worksheet.set_column('D:D', 15, format_texto)
    worksheet.set_column('E:E', 25, format_texto2)
    worksheet.set_column('F:F', 15, format_texto)
    worksheet.set_column('G:G', 15, format_texto)
    worksheet.set_column('H:H', 15, format_texto)
    worksheet.set_column('I:I', 28, format_texto2)
    worksheet.set_column('J:J', 8, format_texto)
    worksheet.set_column('K:K', 35, format_texto2)

    worksheet_segmento.set_column('A:A', 35, format_texto2)
    worksheet_segmento.set_column('B:B', 17, format_numero)

    writer.save()

formatar_consolidado(tabela_consolidada, por_segmento)