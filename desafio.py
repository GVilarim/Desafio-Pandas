import pandas as pd


def alerta(bruto, liquido):
    return 'Sim' if bruto != liquido else 'Não'


caminho = 'arquivos'
# Extração dos dados
arquivo = pd.ExcelFile(f'{caminho}\\detalhamento.xlsx')
df_detalhar_am = pd.read_excel(arquivo, sheet_name='AM', header=1)
df_detalhar_rr = pd.read_excel(arquivo, sheet_name='RR', header=1)
df_detalhar_ro = pd.read_excel(arquivo, sheet_name='RO', header=1)
df_detalhar_ac = pd.read_excel(arquivo, sheet_name='AC', header=1)
df_vendas = pd.read_csv(f'{caminho}\\vendas.csv', sep='|', encoding='utf-8')

# Tratamento dos dados
df_detalhar_ac = df_detalhar_ac.rename({'Valor Bruto': 'Vlookup Bruto'}, axis=1)
df_vendas = df_vendas[['Escritório de vendas', 'Fornecedor', 'Valor Liquido']].copy()
df_vendas = df_vendas.rename({'Escritório de vendas': 'escritorio',
                              'Fornecedor': 'fornecedor',
                              'Valor Liquido': 'valor_liquido'}, axis=1)
df_lojas = pd.concat([df_detalhar_ac, df_detalhar_ro, df_detalhar_rr, df_detalhar_am])
del df_detalhar_am, df_detalhar_rr, df_detalhar_ac, df_detalhar_ro
df_lojas = df_lojas.rename({'NomeFantasia': 'nome_fantasia',
                            'Escritório de vendas': 'escritorio',
                            'Operadora': 'fornecedor',
                            'Vlookup Bruto': 'valor_bruto'}, axis=1)
df_vendas = df_vendas.groupby(['escritorio', 'fornecedor']).agg({'valor_liquido': 'sum'}).reset_index(drop=False)
df_lojas['valor_bruto'] = df_lojas['valor_bruto'].astype(float)

# Correlacionar Tabelas
df_resultado = pd.merge(df_vendas, df_lojas, on=['escritorio', 'fornecedor'], how='inner')

# Regra de Negócio
df_resultado['status'] = df_resultado.apply(lambda row: alerta(row['valor_bruto'], row['valor_liquido']), axis=1)

# Formatação Final
df_resultado = df_resultado.drop(['UF', 'escritorio'], axis=1)
df_resultado = df_resultado.rename({'fornecedor': 'Fornecedor',
                                    'valor_liquido': 'Valor Líquido',
                                    'nome_fantasia': 'Nome Fantasia',
                                    'valor_bruto': 'Valor Bruto',
                                    'status': 'Alerta'}, axis=1)
df_resultado = df_resultado[['Nome Fantasia', 'Fornecedor', 'Valor Líquido', 'Valor Bruto', 'Alerta']]
