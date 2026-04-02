import pandas as pd
import numpy as np

def processar_dados_base(arquivo_up, capital_social):
    """Lê, transforma as abas e gera os DataFrames de evidências."""
    xls = pd.ExcelFile(arquivo_up)

    df_outorga = pd.read_excel(xls, sheet_name='Dados da outorga')
    df_mov     = pd.read_excel(xls, sheet_name='Histórico de movimentações', header=1)
    df_prev    = pd.read_excel(xls, sheet_name='Previsão outorga 2026')
    df_membros = pd.read_excel(xls, sheet_name='Membros')

    # ==========================================
    # LIMPEZA E TRANSFORMAÇÃO GLOBAL
    # ==========================================
    df_mov.columns = df_mov.columns.str.replace('<br/>', ' ').str.replace('\n', '').str.strip()
    df_mov['Órgão Administrativo'] = df_mov['Órgão Administrativo'].replace(
        {'Diretoria': 'Diretoria Estatutária', 'Conselheiro': 'Conselho de Administração'})
    df_mov['Data'] = pd.to_datetime(df_mov['Data'], errors='coerce', dayfirst=True)
    df_mov = df_mov[df_mov['Data'].notna()].copy()
    df_mov['Ano'] = df_mov['Data'].dt.year.astype('Int64')

    df_outorga['Órgão Administrativo'] = df_outorga['Órgão Administrativo'].replace(
        {'Diretoria': 'Diretoria Estatutária', 'Conselheiro': 'Conselho de Administração'})
    df_outorga['Tipo de Plano'] = df_outorga['Tipo de Plano'].fillna('')

    preco_col = 'Preço de Exercício na Outorga'
    df_outorga[preco_col] = pd.to_numeric(
        df_outorga[preco_col].astype(str).str.replace(',', '.'), errors='coerce').fillna(0)
    df_sop   = df_outorga[df_outorga[preco_col] > 0].copy()
    df_acoes = df_outorga[df_outorga[preco_col] == 0].copy()

    cols_num = ['Preço de Exercício Atual', 'Fair Value na Outorga',
                'Outorgado (original)', 'Fair Value Atualizado']
    for df_temp in [df_sop, df_acoes]:
        for col in cols_num:
            if col in df_temp.columns:
                df_temp[col] = pd.to_numeric(
                    df_temp[col].astype(str).str.replace(',', '.'), errors='coerce').fillna(0)
        df_temp['Data de Outorga']   = pd.to_datetime(df_temp['Data de Outorga'],   errors='coerce', dayfirst=True)
        df_temp['Data da Carência']  = pd.to_datetime(df_temp['Data da Carência'],  errors='coerce', dayfirst=True)
        df_temp['Data de Expiração'] = pd.to_datetime(df_temp['Data de Expiração'], errors='coerce', dayfirst=True)

    df_membros['DATA DE ENTRADA'] = pd.to_datetime(df_membros['DATA DE ENTRADA'], errors='coerce', dayfirst=True)
    df_membros['DATA DE SAÍDA']   = pd.to_datetime(df_membros['DATA DE SAÍDA'],   errors='coerce', dayfirst=True)

    # CORREÇÃO NOME (Bug Bruno): normaliza casing para todas as comparações
    df_membros['_nome_lower'] = df_membros['NOME COMPLETO'].str.lower().str.strip()
    df_mov['_nome_lower']     = df_mov['Nome'].str.lower().str.strip()
    df_sop['_nome_lower']     = df_sop['Nome'].str.lower().str.strip()
    df_acoes['_nome_lower']   = df_acoes['Nome'].str.lower().str.strip()

    def calcular_pro_rata(row, ano):
        inicio_ano  = pd.Timestamp(f'{ano}-01-01')
        fim_ano     = pd.Timestamp(f'{ano}-12-31')
        inicio_real = max(row['DATA DE ENTRADA'], inicio_ano) if pd.notnull(row['DATA DE ENTRADA']) else inicio_ano
        fim_real    = min(row['DATA DE SAÍDA'],   fim_ano)   if pd.notnull(row['DATA DE SAÍDA'])   else fim_ano
        if fim_real < inicio_ano or inicio_real > fim_ano:
            return 0.0
        return max(0, (fim_real - inicio_real).days + 1) / 365.0

    df_membros['Pro_Rata_2025'] = df_membros.apply(lambda x: calcular_pro_rata(x, 2025), axis=1)
    df_membros['Pro_Rata_2026'] = df_membros.apply(lambda x: calcular_pro_rata(x, 2026), axis=1)

    # ==========================================
    # IDENTIFICADORES DAS COLUNAS DO HISTÓRICO
    # ==========================================
    col_qtd_mov   = [c for c in df_mov.columns if 'Quantidade de Ações' in c or 'Ações Liberadas' in c][0]
    col_preco_mov  = [c for c in df_mov.columns if 'Preço de Exercício' in c][0]
    col_preco_merc = [c for c in df_mov.columns if 'Preço da Ação' in c or 'Mercado' in c][0]

    evid_85, evid_86, evid_87, evid_88 = [], [], [], []
    evid_89, evid_810, evid_811 = [], [], []
    limite_2025 = pd.Timestamp('2025-12-31')

    # ==========================================
    # SOP — ITENS 8.5 / 8.6 / 8.7 / 8.8
    #
    # CORREÇÃO PRINCIPAL (duplicação/saldo fictício):
    #   O ETL anterior iterava lote a lote e cruzava as movimentações
    #   do programa inteiro contra a qtd de cada lote individualmente,
    #   gerando saldos fantasmas (ex: 233 para André).
    #   Agora agrupamos por (Nome, Programa) e calculamos o saldo
    #   consolidado do programa de uma só vez.
    # ==========================================
    programas_sop = set(df_sop['Programa'].unique())

    for (nome_lower, programa), grupo in df_sop.groupby(['_nome_lower', 'Programa'], sort=False):

        row0     = grupo.iloc[0]
        nome     = row0['Nome']
        orgao    = row0['Órgão Administrativo']
        data_out = row0['Data de Outorga']

        preco_inicial = row0['Preço de Exercício na Outorga']
        preco_atual   = row0['Preço de Exercício Atual']

        # Quantidade total do programa (soma dos lotes)
        qtd_total = grupo['Outorgado (original)'].sum()

        # Movimentações — match por nome normalizado
        movs = df_mov[(df_mov['_nome_lower'] == nome_lower) & (df_mov['Programa'] == programa)]

        exercidas_pre = movs[(movs['Ano'] <= 2024) & (movs['Status'] == 'Exercido')][col_qtd_mov].sum()
        perdidas_pre  = movs[(movs['Ano'] <= 2024) & (movs['Status'].isin(['Cancelado', 'Prescrito', 'Abandonado']))][col_qtd_mov].sum()
        saldo_ini     = qtd_total - exercidas_pre - perdidas_pre

        if saldo_ini > 0:
            evid_85.append([nome, orgao, programa, preco_inicial, saldo_ini, 'Inicial 2025'])

        exercidas_25 = movs[(movs['Ano'] == 2025) & (movs['Status'] == 'Exercido')][col_qtd_mov].sum()
        if exercidas_25 > 0:
            preco_ex_mov = movs[(movs['Ano'] == 2025) & (movs['Status'] == 'Exercido')][col_preco_mov].mean()
            evid_85.append([nome, orgao, programa, preco_ex_mov, exercidas_25, 'Exercidas 2025'])

        perdidas_25 = movs[(movs['Ano'] == 2025) & (movs['Status'].isin(['Cancelado', 'Prescrito', 'Abandonado']))][col_qtd_mov].sum()
        if perdidas_25 > 0:
            evid_85.append([nome, orgao, programa, preco_inicial, perdidas_25, 'Perdidas 2025'])

        saldo_fim = saldo_ini - exercidas_25 - perdidas_25

        # CORREÇÃO Leandro_2026: programas outorgados em 2026 não compõem o saldo de 2025
        if saldo_fim > 0 and pd.notnull(data_out) and data_out.year <= 2025:
            evid_85.append([nome, orgao, programa, preco_inicial, saldo_fim, 'Final 2025'])
            evid_85.append([nome, orgao, programa, preco_atual,   saldo_fim, 'Inicial 2026'])
            evid_85.append([nome, orgao, programa, preco_atual,   saldo_fim, 'Final 2026'])

            st_v = ('Exercível'
                    if (pd.notnull(row0['Data da Carência']) and row0['Data da Carência'] <= limite_2025)
                    else 'Não exercível')
            evid_87.append([
                orgao, nome, programa, st_v, saldo_fim, preco_inicial,
                row0['Fair Value Atualizado'],
                data_out, row0['Data da Carência'], row0['Data de Expiração']
            ])

        # 8.6 — detalhamento de outorgas de 2025
        if pd.notnull(data_out) and data_out.year == 2025 and qtd_total > 0:
            row_last = grupo.sort_values('Data da Carência').iloc[-1]
            evid_86.append([
                f"{programa} ({orgao})", orgao, nome, programa,
                data_out,
                row_last['Data da Carência'],
                row_last['Data de Expiração'],
                qtd_total,
                grupo['Fair Value na Outorga'].mean()
            ])

    # Previsões 2026 para SOP (8.5)
    for i, row in df_prev.iterrows():
        val = str(row.iloc[0])
        if 'Quantidade a ser outorgada' in val:
            qtd = pd.to_numeric(row.iloc[1], errors='coerce')
            try:
                preco = pd.to_numeric(df_prev.iloc[i + 1].iloc[1], errors='coerce')
            except Exception:
                preco = 0
            if pd.notnull(qtd) and qtd > 0:
                o_nome = 'Conselho de Administração' if 'Conselho' in val else 'Diretoria Estatutária'
                prog_n = 'Novo Prog. CA' if 'Conselho' in val else 'Novo Prog. Dir'
                evid_85.extend([
                    [prog_n, o_nome, 'Previsão 2026', preco if pd.notnull(preco) else 0, qtd, 'Novas 2026'],
                    [prog_n, o_nome, 'Previsão 2026', preco if pd.notnull(preco) else 0, qtd, 'Final 2026'],
                ])

    # 8.8 — exercidas 2025 vindas do histórico
    df_ex_25 = df_mov[
        (df_mov['Ano'] == 2025) &
        (df_mov['Status'] == 'Exercido') &
        (df_mov['Programa'].isin(programas_sop))
    ].copy()
    for _, row in df_ex_25.iterrows():
        qtd = pd.to_numeric(str(row[col_qtd_mov]).replace(',', '.'), errors='coerce')
        if pd.notnull(qtd) and qtd > 0:
            evid_88.append([
                row['Órgão Administrativo'], row['Nome'], row['Programa'],
                row['Data'].strftime('%d/%m/%Y'),
                qtd,
                pd.to_numeric(str(row[col_preco_mov]).replace(',', '.'),  errors='coerce'),
                pd.to_numeric(str(row[col_preco_merc]).replace(',', '.'), errors='coerce')
            ])

    # ==========================================
    # RSU — ITENS 8.9 / 8.10 / 8.11
    # Mesma correção: agrupamento por (Nome, Programa)
    # ==========================================
    programas_rsu = set(df_acoes['Programa'].unique())

    for (nome_lower, programa), grupo in df_acoes.groupby(['_nome_lower', 'Programa'], sort=False):
        row0     = grupo.iloc[0]
        nome     = row0['Nome']
        orgao    = row0['Órgão Administrativo']
        data_out = row0['Data de Outorga']
        fv_outorga = grupo['Fair Value na Outorga'].mean()
        fv_atual   = grupo['Fair Value Atualizado'].mean()
        qtd_total  = grupo['Outorgado (original)'].sum()
        ano_out    = data_out.year if pd.notnull(data_out) else 0

        movs = df_mov[(df_mov['_nome_lower'] == nome_lower) & (df_mov['Programa'] == programa)]
        termos_entregue = 'Exercido|Liberado|Entregue|Resgatado'
        termos_perdida  = 'Cancelado|Prescrito|Abandonado'

        if ano_out == 2025:
            saldo_ini    = 0
            outorgadas_25 = qtd_total
        else:
            outorgadas_25 = 0
            entregues_pre = movs[(movs['Ano'] <= 2024) & (movs['Status'].str.contains(termos_entregue, case=False, na=False))][col_qtd_mov].sum()
            perdidas_pre  = movs[(movs['Ano'] <= 2024) & (movs['Status'].str.contains(termos_perdida,  case=False, na=False))][col_qtd_mov].sum()
            saldo_ini     = qtd_total - entregues_pre - perdidas_pre

        entregues_25 = movs[(movs['Ano'] == 2025) & (movs['Status'].str.contains(termos_entregue, case=False, na=False))][col_qtd_mov].sum()
        perdidas_25  = movs[(movs['Ano'] == 2025) & (movs['Status'].str.contains(termos_perdida,  case=False, na=False))][col_qtd_mov].sum()
        saldo_fim    = saldo_ini + outorgadas_25 - entregues_25 - perdidas_25

        if saldo_ini     > 0: evid_89.append([nome, orgao, programa, fv_outorga, saldo_ini,    'Inicial 2025'])
        if outorgadas_25 > 0: evid_89.append([nome, orgao, programa, fv_outorga, outorgadas_25,'Outorgadas 2025'])
        if entregues_25  > 0: evid_89.append([nome, orgao, programa, fv_outorga, entregues_25, 'Entregues 2025'])
        if perdidas_25   > 0: evid_89.append([nome, orgao, programa, fv_outorga, perdidas_25,  'Perdidas 2025'])
        if saldo_fim     > 0: evid_89.append([nome, orgao, programa, fv_atual,   saldo_fim,    'Final 2025'])

        if ano_out == 2025 and qtd_total > 0:
            row_last = grupo.sort_values('Data da Carência').iloc[-1]
            evid_810.append([
                f"{programa} ({orgao})", orgao, nome, programa,
                data_out, row_last['Data da Carência'], qtd_total, fv_outorga
            ])

    for i, row in df_prev.iterrows():
        val = str(row.iloc[0])
        if 'Quantidade a ser outorgada' in val:
            qtd = pd.to_numeric(row.iloc[1], errors='coerce')
            if pd.notnull(qtd) and qtd > 0:
                o_nome = 'Conselho de Administração' if 'Conselho' in val else 'Diretoria Estatutária'
                prog_n = 'Novo Prog. CA (RSU)' if 'Conselho' in val else 'Novo Prog. Dir (RSU)'
                evid_89.extend([
                    [prog_n, o_nome, 'Previsão 2026', 0, qtd, 'Novas 2026'],
                    [prog_n, o_nome, 'Previsão 2026', 0, qtd, 'Final 2026'],
                ])

    # 8.11 — entregues 2025 (RSU)
    df_ent_25 = df_mov[
        (df_mov['Ano'] == 2025) &
        (df_mov['Status'].str.contains('Exercido|Liberado|Entregue|Resgatado', case=False, na=False)) &
        (df_mov['Programa'].isin(programas_rsu))
    ].copy()

    for _, row_ent in df_ent_25.iterrows():
        qtd_ent = pd.to_numeric(str(row_ent[col_qtd_mov]).replace(',', '.'), errors='coerce')
        if pd.notnull(qtd_ent) and qtd_ent > 0:
            p_aq   = pd.to_numeric(str(row_ent[col_preco_mov]).replace(',', '.'),  errors='coerce')
            p_merc = pd.to_numeric(str(row_ent[col_preco_merc]).replace(',', '.'), errors='coerce')
            data_mov = row_ent['Data']

            nome_lower_ent = str(row_ent['Nome']).lower().strip()
            membro_info = df_membros[df_membros['_nome_lower'] == nome_lower_ent]
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

    # ==========================================
    # DATAFRAMES FINAIS
    # ==========================================
    df_evid_85  = pd.DataFrame(evid_85,  columns=['Nome', 'Órgão Administrativo', 'Programa', 'Preço', 'Qtd', 'Status'])
    df_evid_86  = pd.DataFrame(evid_86,  columns=['Coluna_Relatorio', 'Órgão Administrativo', 'Nome', 'Programa', 'Data Outorga', 'Data Carência', 'Data Expiração', 'Qtd_Outorgada', 'Fair_Value'])
    df_evid_87  = pd.DataFrame(evid_87,  columns=['Órgão Administrativo', 'Nome', 'Programa', 'Status_Vesting', 'Qtd_Saldo', 'Preço', 'Fair_Value', 'Data_Outorga', 'Data_Carência', 'Data_Expiração'])
    df_evid_88  = pd.DataFrame(evid_88,  columns=['Órgão Administrativo', 'Nome', 'Programa', 'Data', 'Qtd', 'Preço_Ex', 'Preço_Merc'])
    df_evid_89  = pd.DataFrame(evid_89,  columns=['Nome', 'Órgão Administrativo', 'Programa', 'Fair_Value', 'Qtd', 'Status'])
    df_evid_810 = pd.DataFrame(evid_810, columns=['Coluna_Relatorio', 'Órgão Administrativo', 'Nome', 'Programa', 'Data Outorga', 'Data Carência', 'Qtd_Outorgada', 'Fair_Value'])
    df_evid_811 = pd.DataFrame(evid_811, columns=['Órgão Administrativo', 'Nome', 'Programa', 'Data', 'Qtd', 'Preço_Aquisicao', 'Preço_Mercado', 'Era_Estatutario'])
    df_evid_811 = df_evid_811.drop_duplicates(subset=['Nome', 'Programa', 'Data', 'Qtd'])

    # ==========================================
    # MARCAÇÃO DE MEMBROS (nome normalizado — corrige bug do Bruno)
    # ==========================================
    def nomes_lower(df, col): return set(df[col].str.lower().str.strip().unique()) if not df.empty else set()

    nomes_85_2025 = nomes_lower(df_evid_85[df_evid_85['Status'].str.contains('2025')], 'Nome')
    nomes_85_2026 = nomes_lower(df_evid_85[df_evid_85['Status'].str.contains('2026')], 'Nome')

    df_membros['Tem_85_2025'] = df_membros['_nome_lower'].isin(nomes_85_2025).astype(int)
    df_membros['Tem_85_2026'] = df_membros['_nome_lower'].isin(nomes_85_2026).astype(int)
    df_membros['Tem_86']      = df_membros['_nome_lower'].isin(nomes_lower(df_evid_86,  'Nome')).astype(int)
    df_membros['Tem_87']      = df_membros['_nome_lower'].isin(nomes_lower(df_evid_87,  'Nome')).astype(int)
    df_membros['Tem_88']      = df_membros['_nome_lower'].isin(nomes_lower(df_evid_88,  'Nome')).astype(int)
    df_membros['Tem_89']      = df_membros['_nome_lower'].isin(nomes_lower(df_evid_89,  'Nome')).astype(int)
    df_membros['Tem_810']     = df_membros['_nome_lower'].isin(nomes_lower(df_evid_810, 'Nome')).astype(int)
    df_membros['Tem_811']     = df_membros['_nome_lower'].isin(nomes_lower(df_evid_811, 'Nome')).astype(int)

    # Remove coluna auxiliar
    df_membros = df_membros.drop(columns=['_nome_lower'])

    return {
        'membros':    df_membros,
        'df_outorga': df_outorga,
        '8.5':  df_evid_85,
        '8.6':  df_evid_86,
        '8.7':  df_evid_87,
        '8.8':  df_evid_88,
        '8.9':  df_evid_89,
        '8.10': df_evid_810,
        '8.11': df_evid_811,
    }
