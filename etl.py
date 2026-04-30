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
    df_sop   = df_outorga[(df_outorga[preco_col] > 0) & (df_outorga[preco_col] < 1e8)].copy()
    df_acoes = df_outorga[(df_outorga[preco_col] == 0) | (df_outorga[preco_col] >= 1e8)].copy()

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

    programas_sop = set(df_sop['Programa'].unique())

    for _, lote in df_sop.iterrows():
        nome     = lote['Nome']
        nlow     = lote['_nome_lower']
        orgao    = lote['Órgão Administrativo']
        prog     = lote['Programa']
        lote_num = lote['Lote']
        qtd      = lote['Outorgado (original)']
        data_out = lote['Data de Outorga']

        preco_inicial = lote['Preço de Exercício na Outorga']
        preco_atual   = lote['Preço de Exercício Atual']
        fv_atualizado = lote['Fair Value Atualizado']

        # Filtra movimentações exatamente deste lote (nome + programa + número do lote)
        movs = df_mov[
            (df_mov['_nome_lower'] == nlow) &
            (df_mov['Programa']    == prog) &
            (df_mov['Lote']        == lote_num)
        ]

        # Saldo antes de 2025
        exercidas_pre = movs[(movs['Ano'] <= 2024) & (movs['Status'] == 'Exercido')][col_qtd_mov].sum()
        perdidas_pre  = movs[(movs['Ano'] <= 2024) & (movs['Status'].isin(['Cancelado', 'Prescrito', 'Abandonado']))][col_qtd_mov].sum()
        saldo_ini = qtd - exercidas_pre - perdidas_pre

        # Lote encerrado antes de 2025 — ignora
        if saldo_ini <= 0:
            continue

        if pd.notnull(data_out) and data_out.year > 2025:
            continue

        # Linha Inicial 2025
        evid_85.append([nome, orgao, prog, lote_num, preco_inicial, saldo_ini, 'Inicial 2025'])

        # Movimentações em 2025
        exercidas_25 = movs[(movs['Ano'] == 2025) & (movs['Status'] == 'Exercido')][col_qtd_mov].sum()
        perdidas_25  = movs[(movs['Ano'] == 2025) & (movs['Status'].isin(['Cancelado', 'Prescrito', 'Abandonado']))][col_qtd_mov].sum()

        if exercidas_25 > 0:
            # PE das exercidas vem do histórico (já atualizado na data do exercício)
            preco_ex = movs[(movs['Ano'] == 2025) & (movs['Status'] == 'Exercido')][col_preco_mov].mean()
            evid_85.append([nome, orgao, prog, lote_num, preco_ex, exercidas_25, 'Exercidas 2025'])

        if perdidas_25 > 0:
            evid_85.append([nome, orgao, prog, lote_num, preco_inicial, perdidas_25, 'Perdidas 2025'])

        saldo_fim = saldo_ini - exercidas_25 - perdidas_25

        if saldo_fim > 0:
            # Ignora opções já expiradas antes do final de 2025
            if pd.notnull(lote['Data de Expiração']) and lote['Data de Expiração'] < limite_2025:
                pass
            else:
                evid_85.append([nome, orgao, prog, lote_num, preco_inicial, saldo_fim, 'Final 2025'])
                evid_85.append([nome, orgao, prog, lote_num, preco_atual,   saldo_fim, 'Inicial 2026'])
                evid_85.append([nome, orgao, prog, lote_num, preco_atual,   saldo_fim, 'Final 2026'])

                # 8.7 — opções em aberto no final de 2025
                st_v = ('Exercível'
                        if (pd.notnull(lote['Data da Carência']) and lote['Data da Carência'] <= limite_2025)
                        else 'Não exercível')
                evid_87.append([
                    orgao, nome, prog, lote_num, st_v, saldo_fim, preco_inicial,
                    fv_atualizado, data_out, lote['Data da Carência'], lote['Data de Expiração']
                ])

        # 8.6 — detalhamento de outorgas realizadas em 2025
        if pd.notnull(data_out) and data_out.year == 2025 and qtd > 0:
            evid_86.append([
                f"{prog} ({orgao})", orgao, nome, prog, lote_num,
                data_out, lote['Data da Carência'], lote['Data de Expiração'],
                qtd, lote['Fair Value na Outorga']
            ])

    # Previsões 2026 para SOP (8.5)
    # Suporta duas estruturas de aba:
    #   Formato A (antigo): "Quantidade a ser outorgada" com órgão no mesmo texto
    #   Formato B (novo):   linhas separadas por órgão com descrição completa
    def _ler_prev_col(col_idx=1):
        """Retorna dict {orgao: (qtd, preco, nome_prog)} lido da aba de previsão."""
        result = {}
        n_cols = len(df_prev.columns)
        if col_idx >= n_cols:
            return result
        # Tenta detectar o nome do programa
        nome_prog_prev = 'Novo Programa 2026'
        for i, row in df_prev.iterrows():
            val = str(row.iloc[col_idx]).strip() if pd.notnull(row.iloc[col_idx]) else ''
            desc = str(row.iloc[col_idx - 1]).strip() if col_idx > 0 else ''
            if 'Nome do Programa' in desc and val not in ('nan',''):
                nome_prog_prev = val
        # Lê quantidades por órgão
        mapa_orgao = {
            'Conselho Administrativo': 'Conselho de Administração',
            'Conselho de Administração': 'Conselho de Administração',
            'Diretoria Estatutária': 'Diretoria Estatutária',
            'Diretoria': 'Diretoria Estatutária',
            'Conselho Fiscal': 'Conselho Fiscal',
        }
        preco_prev = 0
        for i, row in df_prev.iterrows():
            desc = str(row.iloc[col_idx - 1]).strip() if col_idx > 0 else ''
            val_raw = row.iloc[col_idx]
            val_num = pd.to_numeric(str(val_raw).replace(',','.'), errors='coerce')
            # Preço da ação na data da outorga
            if 'Preço da ação' in desc or 'Estimativa do Preço' in desc:
                if pd.notnull(val_num):
                    preco_prev = val_num
            # Quantidade por órgão
            if 'Quantidade a ser outorgada' in desc and pd.notnull(val_num) and val_num > 0:
                for chave, orgao_nome in mapa_orgao.items():
                    if chave in desc:
                        result[orgao_nome] = (val_num, preco_prev, nome_prog_prev)
                        break
                else:
                    # Formato antigo: órgão inferido pelo texto
                    if 'Conselho' in desc and 'Fiscal' not in desc:
                        result['Conselho de Administração'] = (val_num, preco_prev, nome_prog_prev)
                    elif 'Fiscal' in desc:
                        result['Conselho Fiscal'] = (val_num, preco_prev, nome_prog_prev)
                    else:
                        result['Diretoria Estatutária'] = (val_num, preco_prev, nome_prog_prev)
        return result

    prev_sop = _ler_prev_col(col_idx=2)  # coluna C (índice 2)
    for orgao_prev, (qtd_prev, preco_prev, nome_prog_prev) in prev_sop.items():
        if qtd_prev > 0:
            evid_85.extend([
                [nome_prog_prev, orgao_prev, 'Previsão 2026', None, preco_prev, qtd_prev, 'Novas 2026'],
                [nome_prog_prev, orgao_prev, 'Previsão 2026', None, preco_prev, qtd_prev, 'Final 2026'],
            ])

    # 8.8 — exercidas 2025 vindas do histórico (por lote)
    df_ex_25 = df_mov[
        (df_mov['Ano'] == 2025) &
        (df_mov['Status'] == 'Exercido') &
        (df_mov['Programa'].isin(programas_sop))
    ].copy()
    for _, row in df_ex_25.iterrows():
        qtd_ex = pd.to_numeric(str(row[col_qtd_mov]).replace(',', '.'), errors='coerce')
        if pd.notnull(qtd_ex) and qtd_ex > 0:
            evid_88.append([
                row['Órgão Administrativo'], row['Nome'], row['Programa'], row['Lote'],
                row['Data'].strftime('%d/%m/%Y'),
                qtd_ex,
                pd.to_numeric(str(row[col_preco_mov]).replace(',', '.'),  errors='coerce'),
                pd.to_numeric(str(row[col_preco_merc]).replace(',', '.'), errors='coerce')
            ])

    # ==========================================
    # RSU — ITENS 8.9 / 8.10 / 8.11
    # Mesma lógica: itera lote a lote
    # ==========================================
    programas_rsu = set(df_acoes['Programa'].unique())

    for _, lote in df_acoes.iterrows():
        nome     = lote['Nome']
        nlow     = lote['_nome_lower']
        orgao    = lote['Órgão Administrativo']
        prog     = lote['Programa']
        lote_num = lote['Lote']
        qtd      = lote['Outorgado (original)']
        data_out = lote['Data de Outorga']
        fv_out   = lote['Fair Value na Outorga']
        fv_atu   = lote['Fair Value Atualizado']
        ano_out  = data_out.year if pd.notnull(data_out) else 0

        movs = df_mov[
            (df_mov['_nome_lower'] == nlow) &
            (df_mov['Programa']    == prog) &
            (df_mov['Lote']        == lote_num)
        ]
        termos_entregue = 'Exercido|Liberado|Entregue|Resgatado'
        termos_perdida  = 'Cancelado|Prescrito|Abandonado'

        if ano_out == 2025:
            saldo_ini    = 0
            outorgadas_25 = qtd
        elif ano_out > 2025:
            continue  # outorga futura, não existia em 2025
        else:
            outorgadas_25 = 0
            entregues_pre = movs[(movs['Ano'] <= 2024) & (movs['Status'].str.contains(termos_entregue, case=False, na=False))][col_qtd_mov].sum()
            perdidas_pre  = movs[(movs['Ano'] <= 2024) & (movs['Status'].str.contains(termos_perdida,  case=False, na=False))][col_qtd_mov].sum()
            saldo_ini     = qtd - entregues_pre - perdidas_pre
            if saldo_ini <= 0:
                continue

        entregues_25 = movs[(movs['Ano'] == 2025) & (movs['Status'].str.contains(termos_entregue, case=False, na=False))][col_qtd_mov].sum()
        perdidas_25  = movs[(movs['Ano'] == 2025) & (movs['Status'].str.contains(termos_perdida,  case=False, na=False))][col_qtd_mov].sum()
        saldo_fim    = saldo_ini + outorgadas_25 - entregues_25 - perdidas_25

        if saldo_ini     > 0: evid_89.append([nome, orgao, prog, lote_num, fv_out, saldo_ini,    'Inicial 2025'])
        if outorgadas_25 > 0: evid_89.append([nome, orgao, prog, lote_num, fv_out, outorgadas_25,'Outorgadas 2025'])
        if entregues_25  > 0: evid_89.append([nome, orgao, prog, lote_num, fv_out, entregues_25, 'Entregues 2025'])
        if perdidas_25   > 0: evid_89.append([nome, orgao, prog, lote_num, fv_out, perdidas_25,  'Perdidas 2025'])
        if saldo_fim     > 0: evid_89.append([nome, orgao, prog, lote_num, fv_atu, saldo_fim,    'Final 2025'])

        if ano_out == 2025 and qtd > 0:
            evid_810.append([
                f"{prog} ({orgao})", orgao, nome, prog, lote_num,
                data_out, lote['Data da Carência'], qtd, fv_out
            ])

    prev_rsu = _ler_prev_col(col_idx=2)
    for orgao_prev, (qtd_prev, _, nome_prog_prev) in prev_rsu.items():
        if qtd_prev > 0:
            prog_n = f"{nome_prog_prev} (RSU)"
            evid_89.extend([
                [prog_n, orgao_prev, 'Previsão 2026', None, 0, qtd_prev, 'Novas 2026'],
                [prog_n, orgao_prev, 'Previsão 2026', None, 0, qtd_prev, 'Final 2026'],
            ])

    # 8.11 — entregues 2025 (RSU) por lote
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

            nlow_ent = str(row_ent['Nome']).lower().strip()
            membro_info = df_membros[df_membros['_nome_lower'] == nlow_ent]
            era_estatutario = False
            if not membro_info.empty:
                entrada = membro_info.iloc[0]['DATA DE ENTRADA']
                saida   = membro_info.iloc[0]['DATA DE SAÍDA']
                era_estatutario = (
                    pd.notnull(entrada) and entrada <= data_mov and
                    (pd.isnull(saida) or saida >= data_mov)
                )

            evid_811.append([
                row_ent['Órgão Administrativo'], row_ent['Nome'], row_ent['Programa'], row_ent['Lote'],
                data_mov.strftime('%d/%m/%Y'), qtd_ent, p_aq, p_merc, era_estatutario
            ])

    # ==========================================
    # DATAFRAMES FINAIS — com coluna Lote
    # ==========================================
    df_evid_85  = pd.DataFrame(evid_85,  columns=['Nome', 'Órgão Administrativo', 'Programa', 'Lote', 'Preço', 'Qtd', 'Status'])
    df_evid_86  = pd.DataFrame(evid_86,  columns=['Coluna_Relatorio', 'Órgão Administrativo', 'Nome', 'Programa', 'Lote', 'Data Outorga', 'Data Carência', 'Data Expiração', 'Qtd_Outorgada', 'Fair_Value'])
    df_evid_87  = pd.DataFrame(evid_87,  columns=['Órgão Administrativo', 'Nome', 'Programa', 'Lote', 'Status_Vesting', 'Qtd_Saldo', 'Preço', 'Fair_Value', 'Data_Outorga', 'Data_Carência', 'Data_Expiração'])
    df_evid_88  = pd.DataFrame(evid_88,  columns=['Órgão Administrativo', 'Nome', 'Programa', 'Lote', 'Data', 'Qtd', 'Preço_Ex', 'Preço_Merc'])
    df_evid_89  = pd.DataFrame(evid_89,  columns=['Nome', 'Órgão Administrativo', 'Programa', 'Lote', 'Fair_Value', 'Qtd', 'Status'])
    df_evid_810 = pd.DataFrame(evid_810, columns=['Coluna_Relatorio', 'Órgão Administrativo', 'Nome', 'Programa', 'Lote', 'Data Outorga', 'Data Carência', 'Qtd_Outorgada', 'Fair_Value'])
    df_evid_811 = pd.DataFrame(evid_811, columns=['Órgão Administrativo', 'Nome', 'Programa', 'Lote', 'Data', 'Qtd', 'Preço_Aquisicao', 'Preço_Mercado', 'Era_Estatutario'])
    df_evid_811 = df_evid_811.drop_duplicates(subset=['Nome', 'Programa', 'Lote', 'Data', 'Qtd'])

    # ==========================================
    # MARCAÇÃO DE MEMBROS
    # ==========================================
    def nomes_lower(df, col):
        return set(df[col].str.lower().str.strip().unique()) if not df.empty else set()

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
