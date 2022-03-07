from tkinter import *
from tkinter import ttk, messagebox, filedialog as fd
import tkinter
import pandas as pd
from datetime import datetime
from os.path import expanduser
import threading

desired_width = 320
pd.set_option('display.width', desired_width)
pd.set_option('display.max_columns', 10)

pd.options.mode.chained_assignment = None  # default='warn'

class Ebitda:

    def __init__(self, tela_login):
        self.tela_login = tela_login
        self.tela_login.geometry('700x600+350+50')
        self.tela_login.title('Ebitda por Segmento')
        self.tela_login.resizable(0, 0)

        j = 0
        r = 0
        for i in range(100):
            c = str(222222+r)
            Frame(self.tela_login, width=10, height=600, bg='#'+c).place(x=j,y=0)
            j = j+10
            r = r+1

        self.frame1 = Frame(self.tela_login, width=600, height=500, bg='white')
        self.frame1.place(x=50, y=50)

        #label
        l1 = Label(self.tela_login, text='Balancete', bg='white')
        l = ('consolas', 13) #fonte, tamanho
        l1.config(font=l)
        l1.place(x=80, y=80)
        self.b1 = Button(self.tela_login, text='Abrir arquivo', command=self.abre_bal)
        self.b1.place(x=280, y=110)
        self.e1 = Entry(tela_login, font=12, text='', bd=1)
        self.e1.place(x=80, y=110)

        # label 2
        self.l2 = Label(self.tela_login, text='Relatório Despesas (FBL3N)', bg='white')
        l = ('consolas', 13)  # fonte, tamanho
        self.l2.config(font=l)
        self.l2.place(x=80, y=180)
        self.b2 = Button(self.tela_login, text='Abrir arquivo', command=self.abre_desp)
        self.b2.place(x=280, y=210)
        self.e2 = Entry(tela_login, font=13, text='', bd=1)
        self.e2.place(x=80, y=210)

        # label 3
        self.l3 = Label(tela_login, text='Drivers de Rateio', bg='white')
        l = ('consolas', 13)  # fonte, tamanho
        self.l3.config(font=l)
        self.l3.place(x=80, y=280)
        self.b3 = Button(self.tela_login, text='Abrir arquivo', command=self.abre_driver)
        self.b3.place(x=280, y=310)
        self.e3 = Entry(self.tela_login, font=13, text='', bd=1)
        self.e3.place(x=80, y=310)

        self.l4 = Label(self.tela_login, text='Digite o Mês do Relatório: (mm-aaaa) ', bg='white', font=l)
        self.l4.place(x=80, y=370)
        self.mes = Entry(self.tela_login, font=12, text='', bd=1)
        self.mes.place(x=80, y=400)

        self.home = expanduser("~")

        self.progress_frame = Frame(self.tela_login, width=400, height=200, bg='azure', bd=3, relief=RIDGE)
        self.texto = Label(self.progress_frame, text='Gerando Relatório', font=('Goudy old style', 12, 'bold'), bd=0). \
            place(x=120, y=60)
        self.pb = ttk.Progressbar(self.progress_frame, mode='indeterminate', length=280)
        self.pb.place(x=50, y=100)

        def start_foo_thread(ev):
            self.foo_thread = threading.Thread(target=lambda: [self.formatar_dados(), self.completar_dados(),
                    self.definir_rateios(), self.rateio_por_segmento(), self.unir_com_balancete(),
                                        self.resumir_segmento(), self.formatar_consolidado()])
            self.foo_thread.daemon = True
            self.progress_frame.place(x=150, y=200)
            self.pb.start()
            self.foo_thread.start()
            self.tela_login.after(20, check_foo_thread)

        def check_foo_thread():
            if self.foo_thread.is_alive():
                self.tela_login.after(20, check_foo_thread)
            else:
                self.pb.stop()
                self.progress_frame.place_forget()
                tkinter.messagebox.showinfo('', 'Relatório Gerado com Sucesso!')

        Button(self.tela_login, font=('consolas', 11, 'bold'), width=20, height=2, fg='white', bg='#FF5733',
               border=1, text='GERAR RELATÓRIO', command=lambda: start_foo_thread(None)).place(x=87, y=475)

    def abre_bal(self):
        self.bal = fd.askopenfilename(title='Abrir arquivo', initialdir=self.home + '\Desktop')
        self.e1.delete(0, END)
        self.e1.insert(0, self.bal)


    def abre_desp(self):
        self.despesas = fd.askopenfilename(title='Abrir arquivo', initialdir=self.home + '\Desktop')
        self.e2.delete(0, END)
        self.e2.insert(0, self.despesas)

    def abre_driver(self):
        self.driver = fd.askopenfilename(title='Abrir arquivo',
                                     initialdir='G:\GECOT\Despesas por Segmento\\')
        self.e3.delete(0, END)
        self.e3.insert(0, self.driver)


    def formatar_dados(self):
        self.dados_inicio_1 = pd.read_excel(self.despesas)
        self.dados_inicio_1 = pd.DataFrame(self.dados_inicio_1)
        self.dados_inicio_1['Conta do Razão'] = self.dados_inicio_1['Conta do Razão'].astype(str)
        self.dados_inicio_3 = self.dados_inicio_1[self.dados_inicio_1['Conta do Razão'].str.contains('6151299020')]
        remover = ['6150', '615129', '6152', '6153', '6154', '6155', '6156', '61604', '6162', '6163',
                   '6164', '6165', '6160299010', '6151151020', '6151151010', '6151151020', '6151152010',
                   '6160351010', '6160351011']
        self.dados_inicio_1 = self.dados_inicio_1[~self.dados_inicio_1['Conta do Razão'].str.contains('|'.join(remover))]
        self.dados_inicio_1 = self.dados_inicio_1[~self.dados_inicio_1['Elemento PEP'].str.contains('RSG', na=False)]
        self.dados_inicio_1 = self.dados_inicio_1[self.dados_inicio_1['Data de lançamento'].notnull()]
        self.dados_inicio_2 = self.dados_inicio_1[self.dados_inicio_1['Centro custo'].isnull()]
        self.dados_inicio_2 = self.dados_inicio_2[~self.dados_inicio_2['Conta do Razão'].str.contains('6151124130')]
        self.dados_inicio_2 = self.dados_inicio_2[~self.dados_inicio_2['Conta do Razão'].str.contains('6151124030')]
        self.dados_inicio_2['Ordem'] = self.dados_inicio_1['Ordem'].astype(str)
        self.dados_inicio_2.loc[self.dados_inicio_2['Ordem'].str.contains('100284'), 'Centro custo'] = '11400'
        self.dados_inicio_2.loc[self.dados_inicio_2['Ordem'].str.contains('100342'), 'Centro custo'] = '11500'
        self.dados_inicio_2.loc[self.dados_inicio_2['Ordem'].str.contains('100283'), 'Centro custo'] = '11300'
        self.dados_inicio_2.loc[self.dados_inicio_2['Ordem'].str.contains('100282'), 'Centro custo'] = '11100'
        self.dados_inicio_2.loc[self.dados_inicio_2['Conta do Razão'].str.contains('61601'), 'Centro custo'] = '11420'
        self.dados_inicio_2.loc[self.dados_inicio_2['Conta do Razão'].str.contains('615'), 'Centro custo'] = '11440'
        self.dados_inicio_1 = self.dados_inicio_1[self.dados_inicio_1['Centro custo'] > 11000.0]
        self.dados_inicio_1 = self.dados_inicio_1[self.dados_inicio_1['Centro custo'] < 14000.0]
        self.dados_inicio_1 = pd.concat([self.dados_inicio_1, self.dados_inicio_2, self.dados_inicio_3])
        del self.dados_inicio_1['Ordem']
        del self.dados_inicio_1['Elemento PEP']
        del self.dados_inicio_1['Atribuição']
        self.dados_inicio_1['Centro custo'] = pd.to_numeric(self.dados_inicio_1['Centro custo'], errors='coerce')


    def completar_dados(self):
        abas = [5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22,
                     23, 24, 25, 26, 27, 28, 29, 30, 31, 32]
        # self.dados_2 = pd.read_excel(self.driver, sheet_name='SIGLAS')
        # self.criterios = pd.read_excel(self.driver, sheet_name=abas, usecols=[4, 5, 6, 7, 8])
        # self.criterios = pd.DataFrame(self.criterios)
        # print(self.criterios.columns)
        self.dados_2, self.criterios = ((pd.read_excel(self.driver, sheet_name='SIGLAS'),
        pd.concat(pd.read_excel(self.driver, sheet_name=abas, usecols=[4, 5, 6, 7, 8]), ignore_index=True)))
        self.busca_sigla = pd.DataFrame(self.dados_2)
        self.busca_sigla['complemento'] = ['.0' for l in self.busca_sigla['Conta do Razão']]
        self.busca_sigla['Conta do Razão'] = self.busca_sigla['Conta do Razão'].astype(str) + self.busca_sigla[
            'complemento']
        self.dados_a_completar = pd.merge(self.dados_inicio_1, self.busca_sigla[['Centro custo', 'Sigla']],
                                          on=['Centro custo'],
                                          how='left')
        self.dados_a_completar['busca'] = self.dados_a_completar['Centro custo'].astype(str) + self.dados_a_completar[
            'Conta do Razão'].astype(str)
        self.dados_a_completar = pd.merge(self.dados_a_completar, self.busca_sigla[['Conta do Razão', 'nome_conta']],
                                          on=['Conta do Razão'], how='left')
        self.criterios = self.criterios[self.criterios['CONTAS CONTÁBEIS'].notnull()]
        self.criterios['busca'] = self.criterios['Unnamed: 4'].astype(str) + self.criterios['CONTAS CONTÁBEIS'].astype(
            str)
        self.dados_a_completar = pd.merge(self.dados_a_completar, self.criterios[['busca', 'Cód. ']], on=['busca'],
                                          how='left')
        check = self.dados_a_completar[self.dados_a_completar['Cód. '].isnull()]
        check.to_excel('erros.xlsx')
        print(check)
        self.dados_a_completar = self.dados_a_completar[self.dados_a_completar['Cód. '].notnull()]
        del self.dados_a_completar['busca']
        del self.dados_a_completar['Tipo de documento']


    def definir_rateios(self):
        self.rateio = pd.read_excel(self.driver, sheet_name='SEGMENTOS')
        self.driver_rateio = pd.DataFrame(self.rateio)
        self.mes_arquivo = self.mes.get()
        self.mes = datetime.strptime(self.mes.get(), '%m-%Y')
        self.driver_rateio = self.driver_rateio[self.driver_rateio['Data'].isin([self.mes])]
        self.driver_rateio.reset_index(inplace=True)
        del self.driver_rateio['index']

    def rateio_por_segmento(self):
        for index, row in self.dados_a_completar.iterrows():
            for i in range(1, 5):
                if row['Cód. '] == i:
                    if i == 1:
                        self.dados_a_completar.loc[index, 'Residencial'] = row['Montante em moeda interna']
                    elif i == 2:
                        self.dados_a_completar.loc[index, 'Comercial'] = row['Montante em moeda interna']
                    elif i == 3:
                        self.dados_a_completar.loc[index, 'Industrial'] = row['Montante em moeda interna']
                    elif i == 4:
                        self.dados_a_completar.loc[index, 'Gás Natural Veicular - GNV'] = row['Montante em moeda interna']
                else:
                    if row['Cód. '] == 7:
                        self.dados_a_completar.loc[index, 'Residencial'] = row['Montante em moeda interna'] * (
                                    self.driver_rateio.loc[0][9] + self.driver_rateio.loc[1][9])
                        self.dados_a_completar.loc[index, 'Comercial'] = row['Montante em moeda interna'] * \
                                                                         self.driver_rateio.loc[2][9]
                        self.dados_a_completar.loc[index, 'Gás Natural Veicular - GNV'] = row['Montante em moeda interna'] * \
                                                                                          self.driver_rateio.loc[4][9]
                        self.dados_a_completar.loc[index, 'Gás Natural - Frotas'] = row['Montante em moeda interna'] * \
                                                                                    self.driver_rateio.loc[6][9]
                    elif row['Cód. '] == 8:
                        self.dados_a_completar.loc[index, 'Residencial'] = row['Montante em moeda interna'] * (
                                self.driver_rateio.loc[0][7] + self.driver_rateio.loc[1][7])
                        self.dados_a_completar.loc[index, 'Comercial'] = row['Montante em moeda interna'] * \
                                                                         self.driver_rateio.loc[2][7]
                        self.dados_a_completar.loc[index, 'Industrial'] = row['Montante em moeda interna'] * (
                                    self.driver_rateio.loc[3][7] + self.driver_rateio.loc[10][7])
                        self.dados_a_completar.loc[index, 'Gás Natural Veicular - GNV'] = row['Montante em moeda interna'] * \
                                                                                          self.driver_rateio.loc[4][7]
                        self.dados_a_completar.loc[index, 'Gás Natural - Frotas'] = row['Montante em moeda interna'] * \
                                                                                    self.driver_rateio.loc[6][7]
                    elif row['Cód. '] == 9:
                        self.dados_a_completar.loc[index, 'Residencial'] = row['Montante em moeda interna'] * (
                                self.driver_rateio.loc[0][3] + self.driver_rateio.loc[1][3])
                        self.dados_a_completar.loc[index, 'Comercial'] = row['Montante em moeda interna'] * \
                                                                         self.driver_rateio.loc[2][3]
                        self.dados_a_completar.loc[index, 'Industrial'] = row['Montante em moeda interna'] * (
                                self.driver_rateio.loc[3][3] + self.driver_rateio.loc[10][3])
                        self.dados_a_completar.loc[index, 'Gás Natural Veicular - GNV'] = row['Montante em moeda interna'] * \
                                                                                          self.driver_rateio.loc[4][3]
                        self.dados_a_completar.loc[index, 'Gás Natural - Frotas'] = row['Montante em moeda interna'] * \
                                                                                    self.driver_rateio.loc[6][3]
                    elif row['Cód. '] == 10:
                        self.dados_a_completar.loc[index, 'Industrial'] = row['Montante em moeda interna'] * (
                                self.driver_rateio.loc[3][11] + self.driver_rateio.loc[10][11])
                        self.dados_a_completar.loc[index, 'Gás Natural Veicular - GNV'] = row['Montante em moeda interna'] * \
                                                                                          self.driver_rateio.loc[4][11]
                        self.dados_a_completar.loc[index, 'Gás Natural - Frotas'] = row['Montante em moeda interna'] * \
                                                                                    self.driver_rateio.loc[6][11]
                    elif row['Cód. '] == 13 or row['Cód. '] == 12:
                        self.dados_a_completar.loc[index, 'Residencial'] = row['Montante em moeda interna'] * (
                                self.driver_rateio.loc[0][5] + self.driver_rateio.loc[1][5])
                        self.dados_a_completar.loc[index, 'Comercial'] = row['Montante em moeda interna'] * \
                                                                         self.driver_rateio.loc[2][5]
                        self.dados_a_completar.loc[index, 'Industrial'] = row['Montante em moeda interna'] * (
                                self.driver_rateio.loc[3][5] + self.driver_rateio.loc[10][5])
                        self.dados_a_completar.loc[index, 'Gás Natural Veicular - GNV'] = row['Montante em moeda interna'] * \
                                                                                          self.driver_rateio.loc[4][5]
                        self.dados_a_completar.loc[index, 'Gás Natural - Frotas'] = row['Montante em moeda interna'] * \
                                                                                    self.driver_rateio.loc[6][5]
                    elif row['Cód. '] == 14:
                        self.dados_a_completar.loc[index, 'Residencial'] = row['Montante em moeda interna'] * (
                                self.driver_rateio.loc[0][13] + self.driver_rateio.loc[1][13])
                        self.dados_a_completar.loc[index, 'Comercial'] = row['Montante em moeda interna'] * \
                                                                         self.driver_rateio.loc[2][13]

        self.tabela_pronta = pd.melt(self.dados_a_completar,
                            id_vars=['Data de lançamento', 'Data do documento', 'Montante em moeda interna',
                                     'Nº documento', 'Texto', 'Conta do Razão', 'Centro custo', 'Sigla', 'nome_conta',
                                     'Cód. '], var_name='Segmento', value_name='Valor')
        for index, row in self.tabela_pronta.iterrows():
            if row['Sigla'] != 'CUSTO' or row['Sigla'] != 'DESPESA':
                self.tabela_pronta.loc[index, 'Montante em moeda interna'] = row['Valor']
        del self.tabela_pronta['Valor']


    def unir_com_balancete(self):
        self.balancete_2 = pd.read_excel(self.bal)
        self.balancete_2.drop(['Saldo Inicial', 'Movimentação a Débito', 'Movimentação a Crédito', 'Saldo Acumulado'], axis=1,
                            inplace=True)
        self.balancete_2['Conta do Razão'] = pd.to_numeric(self.balancete_2['Conta do Razão'], errors='coerce')
        self.balancete = self.balancete_2.loc[self.balancete_2['Conta do Razão'] > 6000000000]
        for index, row in self.balancete.iterrows():
            if row['Conta do Razão'] < 6130000000:
                self.balancete.loc[index, 'Sigla'] = 'RECEITA'
            else:
                self.balancete.loc[index, 'Sigla'] = 'CUSTO'

        selecao_bal = ['611', '612', '615013', '615023', '615029', '615213', '615313', '615413', '615613', '6152192001',
                       '6153192001',
                       '6154192001', '6156192001', '6150192001', '6150292001']
        self.balancete['Conta do Razão'] = self.balancete['Conta do Razão'].astype(str)
        self.balancete['Segmento'] = self.balancete['Texto Conta do Razão'].str.split().str[-1]

        self.balancete = self.balancete[self.balancete['Conta do Razão'].str.contains('|'.join(selecao_bal))]
        texto = {'GNC': 'Industrial', 'RESIDENCIAL': 'Residencial',
                 'INDUSTRIAL': 'Industrial', 'GNV': 'Gás Natural Veicular - GNV', 'COMERCIAL': 'Comercial',
                 'PRIMA': 'Industrial', 'MAT.PRIMA': 'Industrial', '-MAP.PRIMA': 'Industrial',
                 'REFRIGERAÇÃO': 'Industrial', 'REFRIGERAÇAO': 'Industrial', 'DISTRIBUIDA': 'Residencial',
                 'GER.DISTRIBUIDA': 'Residencial'}
        for i, j in texto.items():
            self.balancete['Segmento'] = self.balancete['Segmento'].replace(i, j)
        self.balancete.rename(
            columns={'Total Movimentação': 'Montante em moeda interna', 'Texto Conta do Razão': 'nome_conta'}, inplace=True)
        self.balancete['Data de lançamento'] = datetime.strptime(self.mes_arquivo, '%m-%Y')
        self.balancete['Data do documento'] = self.balancete['Data de lançamento']
        self.balancete['Centro custo'] = self.balancete['Sigla']
        self.balancete_rateado = pd.DataFrame()
        for index, row in self.balancete.iterrows():
            if row['Segmento'] == 'Gás Natural Veicular - GNV':
                self.balancete.loc[index, 'Montante em moeda interna'] = row['Montante em moeda interna'] * \
                                                                         self.driver_rateio.loc[4][2] / (
                                                                                     self.driver_rateio.loc[4][2] +
                                                                                     self.driver_rateio.loc[6][2])
                self.balancete.loc[[index], ['Segmento']] = 'Gás Natural Veicular - GNV'
                self.balancete_rateado = self.balancete_rateado.append(self.balancete.loc[[index]])
                self.balancete.loc[index, 'Montante em moeda interna'] = row['Montante em moeda interna'] * \
                                                                         self.driver_rateio.loc[6][2] / (
                                                                                     self.driver_rateio.loc[4][2] +
                                                                                     self.driver_rateio.loc[6][2])
                self.balancete.loc[[index], ['Segmento']] = 'Gás Natural - Frotas'
                self.balancete_rateado = self.balancete_rateado.append(self.balancete.loc[[index]])
            else:
                self.balancete_rateado = self.balancete_rateado.append(self.balancete.loc[[index]])
        self.tabela_pronta = pd.concat([self.tabela_pronta, self.balancete_rateado]).fillna(0)
        self.tabela_pronta['Montante em moeda interna'] = self.tabela_pronta['Montante em moeda interna'] * -1


    def resumir_segmento(self):
        tabela = [[], []]
        for s in self.tabela_pronta['Segmento'].unique():
            tabela[0].append(s)
            tabela[1].append(self.tabela_pronta[self.tabela_pronta['Segmento'] == s]['Montante em moeda interna'].sum())
        self.datanova = pd.DataFrame({'Segmento': tabela[0], 'Montante': tabela[1]})
        self.datanova.loc['TOTAL GERAL'] = self.datanova.iloc[:, 1:].sum(axis=0)

    def formatar_consolidado(self):
        writer = pd.ExcelWriter('G:\GECOT\Despesas por Segmento\Despesas por Segmento ' + self.mes_arquivo + '.xlsx',
                                engine='xlsxwriter')
        item = pd.DataFrame(self.tabela_pronta)
        item.rename(columns={'Montante em moeda interna': 'Montante'}, inplace=True)
        item['Data de lançamento'] = item['Data de lançamento'].dt.date
        item['Data do documento'] = item['Data do documento'].dt.date
        # item.loc['TOTAL GERAL'] = item.loc[:, 'Montante'].sum(axis=0)
        item.to_excel(writer, sheet_name='Geral', index=False)
        self.datanova.to_excel(writer, sheet_name='Resumo', index=False)

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



tela_login = Tk()
aplicacao = Ebitda(tela_login)
tela_login.mainloop()
