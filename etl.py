# etl.py
import pandas as pd
import numpy as np

def processar_dados_base(arquivo_up, capital_social):
    """Lê, transforma as abas e gera os DataFrames de evidências."""
    xls = pd.ExcelFile(arquivo_up)
    
    df_outorga = pd.read_excel(xls, sheet_name='Dados da outorga')
    df_mov = pd.read_excel(xls, sheet_name='Histórico de movimentações', header=1)
    df_prev = pd.read_excel(xls, sheet_name='Previsão outorga 2026')
    df_membros = pd.read_excel(xls, sheet_name='Membros')

    # ==========================================
    # LIMPEZA E TRANSFORMAÇÃO GLOBAL
    # ==========================================
    df_mov.columns = df_mov.columns.str.replace('<br/>', ' ').str.replace('\n', '').str.strip()
    df_mov['Órgão Administrativo'] = df_mov['Órgão Administrativo'].replace({'Diretoria': 'Diretoria Estatutária', 'Conselheiro': 'Conselho de Administração'})
    df_mov['Data'] = pd.to_datetime(df_mov['Data'], errors='coerce', dayfirst=True)
    df_mov = df_mov[df_mov['Data'].notna()]
    df_mov['Ano'] = df_mov['Data'].dt.year.astype('Int64')

    df_outorga['Órgão Administrativo'] = df_outorga['Órgão Administrativo'].replace({'Diretoria': 'Diretoria Estatutária', 'Conselheiro': 'Conselho de Administração'})
    df_outorga['Tipo de Plano'] = df_outorga['Tipo de Plano'].fillna('')

    preco_col = 'Preço de Exercício Atual'
    df_outorga[preco_col] = pd.to_numeric(
        df_outorga[preco_col].astype(str).str.replace(',', '.'), errors='coerce'
    ).fillna(0)
    df_sop   = df_outorga[df_outorga[preco_col] > 0].copy()
    df_acoes = df_outorga[df_outorga[preco_col] == 0].copy()

    cols_num = ['Preço de Exercício Atual', 'Fair Value na Outorga', 'Outorgado (original)', 'Fair Value Atualizado']
    for df_temp in [df_sop, df_acoes]:
        for col in cols_num:
            if col in df_temp.columns:
                df_temp[col] = pd.to_numeric(df_temp[col].astype(str).str.replace(',', '.'), errors='coerce').fillna(0)
        df_temp['Data de Outorga']   = pd.to_datetime(df_temp['Data de Outorga'],   errors='coerce', dayfirst=True)
        df_temp['Data da Carência']  = pd.to_datetime(df_temp['Data da Carência'],  errors='coerce', dayfirst=True)
        df_temp['Data de Expiração'] = pd.to_datetime(df_temp['Data de Expiração'], errors='coerce', dayfirst=True)

    df_membros['DATA DE ENTRADA'] = pd.to_datetime(df_membros['DATA DE ENTRADA'], errors='coerce')
    df_membros['DATA DE SAÍDA']   = pd.to_datetime(df_membros['DATA DE SAÍDA'],   errors='coerce', dayfirst=True)

    def calcular_pro_rata(row, ano):
        inicio_ano = pd.Timestamp(f'{ano}-01-01'); fim_ano = pd.Timestamp(f'{ano}-12-31')
        inicio_real = max(row['DATA DE ENTRADA'], inicio_ano) if pd.notnull(row['DATA DE ENTRADA']) else inicio_ano
        fim_real    = min(row['DATA DE SAÍDA'],   fim_ano)   if pd.notnull(row['DATA DE SAÍDA'])   else fim_ano
        if fim_real < inicio_ano or inicio_real > fim_ano: return 0.0
        return max(0, (fim_real - inicio_real).days + 1) / 365.0

    df_membros['Pro_Rata_2025'] = df_membros.apply(lambda x: calcular_pro_rata(x, 2025), axis=1)
    df_membros['Pro_Rata_2026'] = df_membros.apply(lambda x: calcular_pro_rata(x, 2026), axis=1)

    # ==========================================
    # GERAÇÃO DAS EVIDÊNCIAS
    # ==========================================
    col_qtd_mov  = [c for c in df_mov.columns if 'Quantidade de Ações' in c or 'Ações Liberadas' in c][0]
    col_preco_mov = [c for c in df_mov.columns if 'Preço de Exercício' in c][0]
    col_preco_merc = [c for c in df_mov.columns if 'Preço da Ação' in c or 'Mercado' in c][0]

    evid_85, evid_86, evid_87, evid_88 = [], [], [], []
    evid_89, evid_810, evid_811 = [], [], []
    limite_2025 = pd.Timestamp('2025-12-31')

    # ----------------- OPÇÕES (8.5, 8.6, 8.7, 8.8) -----------------
    for _, row in df_sop.iterrows():
        nome, orgao, programa = row['Nome'], row['Órgão Administrativo'], row['Programa']
        qtd_original = row['Outorgado (original)']
        preco_atual  = row['Preço de Exercício Atual']
        preco_inicial = row['Preço de Exercício na Outorga']

        movs = df_mov[(df_mov['Nome'] == nome) & (df_mov['Programa'] == programa)]
        exercidas_pre = movs[(movs['Ano'] <= 2024) & (movs['Status'] == 'Exercido')][col_qtd_mov].sum()
        perdidas_pre  = movs[(movs['Ano'] <= 2024) & (movs['Status'].isin(['Cancelado', 'Prescrito', 'Abandonado']))][col_qtd_mov].sum()

        saldo_ini = qtd_original - exercidas_pre - perdidas_pre
        if saldo_ini > 0:
            evid_85.append([nome, orgao, programa, preco_inicial, saldo_ini, 'Inicial 2025'])

        exercidas_25 = movs[(movs['Ano'] == 2025) & (movs['Status'] == 'Exercido')][col_qtd_mov].sum()
        if exercidas_25 > 0:
            # Preço de exercício das opções exercidas vem do histórico de movimentações (correto)
            preco_ex_mov = movs[(movs['Ano'] == 2025) & (movs['Status'] == 'Exercido')][col_preco_mov].mean()
            evid_85.append([nome, orgao, programa, preco_ex_mov, exercidas_25, 'Exercidas 2025'])

        perdidas_25 = movs[(movs['Ano'] == 2025) & (movs['Status'].isin(['Cancelado', 'Prescrito', 'Abandonado']))][col_qtd_mov].sum()
        if perdidas_25 > 0:
            evid_85.append([nome, orgao, programa, preco_inicial, perdidas_25, 'Perdidas 2025'])

        saldo_fim = saldo_ini - exercidas_25 - perdidas_25
        if saldo_fim > 0:
            evid_85.append([nome, orgao, programa, preco_inicial, saldo_fim, 'Final 2025'])
            evid_85.append([nome, orgao, programa, preco_atual,   saldo_fim, 'Inicial 2026'])
            evid_85.append([nome, orgao, programa, preco_atual,   saldo_fim, 'Final 2026'])
            st_v = 'Exercível' if (pd.notnull(row['Data da Carência']) and row['Data da Carência'] <= limite_2025) else 'Não exercível'
            evid_87.append([orgao, nome, programa, st_v, saldo_fim, preco_inicial, row['Fair Value Atualizado'],
                            row['Data de Outorga'], row['Data da Carência'], row['Data de Expiração']])

        if pd.notnull(row['Data de Outorga']) and row['Data de Outorga'].year == 2025 and qtd_original > 0:
            evid_86.append([f"{programa} ({orgao})", orgao, nome, programa,
                            row['Data de Outorga'], row['Data da Carência'], row['Data de Expiração'],
                            qtd_original, row['Fair Value na Outorga']])

    # Previsões 2026 (8.5 - SOP)
    for i, row in df_prev.iterrows():
        val = str(row.iloc[0])
        if "Quantidade a ser outorgada" in val:
            qtd   = pd.to_numeric(row.iloc[1], errors='coerce')
            preco = pd.to_numeric(df_prev.iloc[i+(2 if "Conselho" in val else 1)].iloc[1], errors='coerce')
            if qtd > 0:
                o_nome = 'Conselho de Administração' if "Conselho" in val else 'Diretoria Estatutária'
                prog_n = 'Novo Prog. CA' if "Conselho" in val else 'Novo Prog. Dir'
                evid_85.extend([
                    [prog_n, o_nome, 'Previsão 2026', preco, qtd, 'Novas 2026'],
                    [prog_n, o_nome, 'Previsão 2026', preco, qtd, 'Final 2026']
                ])

    # Exercidas 2025 (8.8)
    df_ex_25 = df_mov[(df_mov['Ano'] == 2025) & (df_mov['Status'] == 'Exercido')].copy()
    df_ex_25 = df_ex_25[
        df_ex_25['Plano'].str.contains('Stock Options', case=False, na=False) |
        df_ex_25['Programa'].str.contains('Stock Options', case=False, na=False)
    ]
    for _, row in df_ex_25.iterrows():
        qtd = pd.to_numeric(str(row[col_qtd_mov]).replace(',', '.'), errors='coerce')
        if qtd > 0:
            evid_88.append([
                row['Órgão Administrativo'], row['Nome'], row['Programa'],
                row['Data'].strftime('%d/%m/%Y'), qtd,
                pd.to_numeric(str(row[col_preco_mov]).replace(',', '.'),  errors='coerce'),
                pd.to_numeric(str(row[col_preco_merc]).replace(',', '.'), errors='coerce')
            ])

    # ----------------- AÇÕES (8.9, 8.10, 8.11) -----------------
    for _, row in df_acoes.iterrows():
        nome, orgao, programa = row['Nome'], row['Órgão Administrativo'], row['Programa']
        qtd_original, fv_outorga, fv_atual = row['Outorgado (original)'], row['Fair Value na Outorga'], row['Fair Value Atualizado']
        ano_out = row['Data de Outorga'].year if pd.notnull(row['Data de Outorga']) else 0

        movs = df_mov[(df_mov['Nome'] == nome) & (df_mov['Programa'] == programa)]
        termos_entregue = 'Exercido|Liberado|Entregue|Resgatado'
        termos_perdida  = 'Cancelado|Prescrito|Abandonado'

        if ano_out == 2025:
            saldo_ini = 0; outorgadas_25 = qtd_original
        else:
            outorgadas_25 = 0
            entregues_pre = movs[(movs['Ano'] <= 2024) & (movs['Status'].str.contains(termos_entregue, case=False, na=False))][col_qtd_mov].sum()
            perdidas_pre  = movs[(movs['Ano'] <= 2024) & (movs['Status'].str.contains(termos_perdida,  case=False, na=False))][col_qtd_mov].sum()
            saldo_ini = qtd_original - entregues_pre - perdidas_pre

        entregues_25 = movs[(movs['Ano'] == 2025) & (movs['Status'].str.contains(termos_entregue, case=False, na=False))][col_qtd_mov].sum()
        perdidas_25  = movs[(movs['Ano'] == 2025) & (movs['Status'].str.contains(termos_perdida,  case=False, na=False))][col_qtd_mov].sum()
        saldo_fim = saldo_ini + outorgadas_25 - entregues_25 - perdidas_25

        if saldo_ini     > 0: evid_89.append([nome, orgao, programa, fv_outorga, saldo_ini,     'Inicial 2025'])
        if outorgadas_25 > 0: evid_89.append([nome, orgao, programa, fv_outorga, outorgadas_25, 'Outorgadas 2025'])
        if entregues_25  > 0: evid_89.append([nome, orgao, programa, fv_outorga, entregues_25,  'Entregues 2025'])
        if perdidas_25   > 0: evid_89.append([nome, orgao, programa, fv_outorga, perdidas_25,   'Perdidas 2025'])
        if saldo_fim     > 0: evid_89.append([nome, orgao, programa, fv_atual,   saldo_fim,     'Final 2025'])

        if ano_out == 2025 and qtd_original > 0:
            evid_810.append([f"{programa} ({orgao})", orgao, nome, programa,
                             row['Data de Outorga'], row['Data da Carência'], qtd_original, fv_outorga])

    for i, row in df_prev.iterrows():
        val = str(row.iloc[0])
        if "Quantidade a ser outorgada" in val:
            qtd = pd.to_numeric(row.iloc[1], errors='coerce')
            if qtd > 0:
                o_nome = 'Conselho de Administração' if "Conselho" in val else 'Diretoria Estatutária'
                prog_n = 'Novo Prog. CA (RSU)' if "Conselho" in val else 'Novo Prog. Dir (RSU)'
                evid_89.extend([
                    [prog_n, o_nome, 'Previsão 2026', 0, qtd, 'Novas 2026'],
                    [prog_n, o_nome, 'Previsão 2026', 0, qtd, 'Final 2026']
                ])

    # Entregues 2025 (8.11)
    df_entregues_25 = df_mov[
        (df_mov['Ano'] == 2025) &
        (df_mov['Status'].str.contains('Exercido|Liberado|Entregue|Resgatado', case=False, na=False))
    ].copy()
    df_entregues_25 = df_entregues_25[
        df_entregues_25['Plano'].str.contains('Ações Restritas|Performance|Matching', case=False, na=False) |
        df_entregues_25['Programa'].str.contains('Ações Restritas|Performance|Matching', case=False, na=False)
    ]

    for _, row_ent in df_entregues_25.iterrows():
        qtd_ent = pd.to_numeric(str(row_ent[col_qtd_mov]).replace(',', '.'), errors='coerce')
        if qtd_ent > 0:
            p_aq   = pd.to_numeric(str(row_ent[col_preco_mov]).replace(',', '.'),  errors='coerce')
            p_merc = pd.to_numeric(str(row_ent[col_preco_merc]).replace(',', '.'), errors='coerce')
            data_mov = row_ent['Data']

            membro_info = df_membros[df_membros['NOME COMPLETO'] == row_ent['Nome']]
            era_estatutario = False
            if not membro_info.empty:
                entrada = membro_info.iloc[0]['DATA DE ENTRADA']
                saida   = membro_info.iloc[0]['DATA DE SAÍDA']
                era_estatutario = (
                    pd.notnull(entrada) and entrada <= data_mov and
                    (pd.isnull(saida) or saida >= data_mov)
                )

            evid_811.append([
                row_ent['Órgão Administrativo'], row_ent['Nome'], row_ent['Programa'],
                data_mov.strftime('%d/%m/%Y'), qtd_ent, p_aq, p_merc, era_estatutario
            ])

    # Convertendo tudo para DataFrames
    df_evid_85  = pd.DataFrame(evid_85,  columns=['Nome', 'Órgão Administrativo', 'Programa', 'Preço', 'Qtd', 'Status'])
    df_evid_86  = pd.DataFrame(evid_86,  columns=['Coluna_Relatorio', 'Órgão Administrativo', 'Nome', 'Programa', 'Data Outorga', 'Data Carência', 'Data Expiração', 'Qtd_Outorgada', 'Fair_Value'])
    df_evid_87  = pd.DataFrame(evid_87,  columns=['Órgão Administrativo', 'Nome', 'Programa', 'Status_Vesting', 'Qtd_Saldo', 'Preço', 'Fair_Value', 'Data_Outorga', 'Data_Carência', 'Data_Expiração'])
    df_evid_88  = pd.DataFrame(evid_88,  columns=['Órgão Administrativo', 'Nome', 'Programa', 'Data', 'Qtd', 'Preço_Ex', 'Preço_Merc'])
    df_evid_89  = pd.DataFrame(evid_89,  columns=['Nome', 'Órgão Administrativo', 'Programa', 'Fair_Value', 'Qtd', 'Status'])
    df_evid_810 = pd.DataFrame(evid_810, columns=['Coluna_Relatorio', 'Órgão Administrativo', 'Nome', 'Programa', 'Data Outorga', 'Data Carência', 'Qtd_Outorgada', 'Fair_Value'])
    df_evid_811 = pd.DataFrame(evid_811, columns=['Órgão Administrativo', 'Nome', 'Programa', 'Data', 'Qtd', 'Preço_Aquisicao', 'Preço_Mercado', 'Era_Estatutario'])
    df_evid_811 = df_evid_811.drop_duplicates(subset=['Nome', 'Programa', 'Data', 'Qtd'])

    # Marcação de Membros
    df_membros['Tem_85_2025'] = df_membros['NOME COMPLETO'].isin(df_evid_85[df_evid_85['Status'].str.contains('2025')]['Nome'].unique()).astype(int)
    df_membros['Tem_85_2026'] = df_membros['NOME COMPLETO'].isin(df_evid_85[df_evid_85['Status'].str.contains('2026')]['Nome'].unique()).astype(int)
    df_membros['Tem_86']  = df_membros['NOME COMPLETO'].isin(df_evid_86['Nome'].unique()).astype(int)
    df_membros['Tem_87']  = df_membros['NOME COMPLETO'].isin(df_evid_87['Nome'].unique()).astype(int)
    df_membros['Tem_88']  = df_membros['NOME COMPLETO'].isin(df_evid_88['Nome'].unique()).astype(int)
    df_membros['Tem_89']  = df_membros['NOME COMPLETO'].isin(df_evid_89['Nome'].unique()).astype(int)
    df_membros['Tem_810'] = df_membros['NOME COMPLETO'].isin(df_evid_810['Nome'].unique()).astype(int)
    df_membros['Tem_811'] = df_membros['NOME COMPLETO'].isin(df_evid_811['Nome'].unique()).astype(int)

    return {
        "membros": df_membros, "df_outorga": df_outorga,
        "8.5": df_evid_85, "8.6": df_evid_86, "8.7": df_evid_87,
        "8.8": df_evid_88, "8.9": df_evid_89, "8.10": df_evid_810, "8.11": df_evid_811
    }
