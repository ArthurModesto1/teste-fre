# export.py
import pandas as pd
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Alignment
from io import BytesIO

# Dicionários de tradução de modelos de precificação e volatilidade
MODEL_OPTIONS = {
    0: "Binomial",
    1: "Black & Scholes",
    2: "Bjerkund-Stensland",
    3: "Último Pregão",
    4: "Média dos últimos Pregões",
    5: "Monte Carlo"
}
MODEL_VOLATILITY = {
    0: "N/A",
    1: "Desvio Padrão"
}

def gerar_excel_final(dados, capital_social):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    orgaos = ['Conselho de Administração', 'Diretoria Estatutária', 'Conselho Fiscal']
    def get_t(orgao): return "*administração*" if "Administração" in orgao else ("*fiscal*" if "Fiscal" in orgao else "*diretoria*")
    def get_col(i): return openpyxl.utils.get_column_letter(i)

    df_membros, df_outorga = dados['membros'], dados['df_outorga']
    df_evid_85, df_evid_86, df_evid_87 = dados['8.5'], dados['8.6'], dados['8.7']
    df_evid_88, df_evid_89, df_evid_810, df_evid_811 = dados['8.8'], dados['8.9'], dados['8.10'], dados['8.11']

    # 4.1 Abas de Evidências (Ocultas/Auxiliares no Excel)
    col_org_membros = 'Orgão' if 'Orgão' in df_membros.columns else 'Órgão Administrativo'
    colunas_membros_export = [
        col_org_membros,
        'NOME COMPLETO', 'DATA DE ENTRADA', 'DATA DE SAÍDA',
        'Pro_Rata_2025', 'Pro_Rata_2026',
        'Tem_85_2025', 'Tem_85_2026', 'Tem_86', 'Tem_87', 'Tem_88', 'Tem_89', 'Tem_810', 'Tem_811'
    ]

    dfs_evid = [
        ('Evid_Membros', df_membros[colunas_membros_export]),
        ('Evid_85',  df_evid_85),
        ('Evid_86',  df_evid_86),
        ('Evid_87',  df_evid_87),
        ('Evid_89',  df_evid_89),
        ('Evid_810', df_evid_810)
    ]
    max_rows = {}
    for nome_aba, df_e in dfs_evid:
        ws_e = wb.create_sheet(nome_aba)
        for r in dataframe_to_rows(df_e, index=False, header=True): ws_e.append(r)
        max_rows[nome_aba] = ws_e.max_row

    # Evidências com fórmulas de ganho (8.8 e 8.11)
    # 8.8
    ws_e88 = wb.create_sheet('Evid_88')
    ws_e88.append(list(df_evid_88.columns) + ['Ganho_Bruto_Formula'])
    for i, r in enumerate(dataframe_to_rows(df_evid_88, index=False, header=False), start=2):
        ws_e88.append(list(r) + [f"=E{i}*(G{i}-F{i})"])
    max_rows['Evid_88'] = ws_e88.max_row

    # 8.11 inclui coluna Era_Estatutario (coluna H) + fórmula de ganho (coluna I)
    ws_e811 = wb.create_sheet('Evid_811')
    ws_e811.append(list(df_evid_811.columns) + ['Ganho_Bruto_Formula'])
    for i, r in enumerate(dataframe_to_rows(df_evid_811, index=False, header=False), start=2):
        ws_e811.append(list(r) + [f"=E{i}*(G{i}-F{i})"])
    max_rows['Evid_811'] = ws_e811.max_row

    def formatar_prazos(df_filtrado, tipo_data, acoes=False):
        if df_filtrado.empty: return "-"
        linhas = []
        for (dout, dalvo), grp in df_filtrado.groupby(['Data_Outorga', tipo_data]):
            if pd.isnull(dout) or pd.isnull(dalvo): continue
            # Usa o nome do programa em vez da data de outorga
            nome_prog = grp['Programa'].iloc[0] if 'Programa' in grp.columns else dout.strftime('%d/%m/%Y')
            qtd = grp['Qtd' if acoes else 'Qtd_Saldo'].sum()
            linhas.append(
                f"{nome_prog}: {qtd:,.0f} {'ações' if acoes else 'opções'} até {dalvo.strftime('%d/%m/%Y')}".replace(',', '.')
            )
        return "\n".join(linhas) if linhas else "-"

    # ==================== QUADROS ====================

    # 8.5 e 8.9
    def criar_quadro_mov(num, df_e, max_r, ano, is_rsu=False):
        ws = wb.create_sheet(f'Quadro_{num}')
        ws.append(['Grupos'] + orgaos)
        aba_evid = f'Evid_{num.replace("_2025","").replace("_2026","").replace(".", "")}'
        labels = [
            "Nº total de membros", "Nº de membros remunerados", "Diluição potencial (%)", "Esclarecimento",
            "Em aberto no início do exercício (R$)", "Perdidas e expiradas no exercício (R$)",
            "Exercidas no exercício (R$)", "Em aberto no final do exercício (R$)"
        ]
        for i, lbl in enumerate(labels, start=2):
            ws.append([lbl])
            for c_idx, org in enumerate(orgaos):
                cel = f"{get_col(c_idx + 2)}{i}"
                if "total de membros" in lbl:
                    col_p = "E" if ano == 2025 else "F"
                    ws[cel] = f'=SUMIF(Evid_Membros!$A$2:$A${max_rows["Evid_Membros"]}, "{get_t(org)}", Evid_Membros!${col_p}$2:${col_p}${max_rows["Evid_Membros"]})'
                    ws[cel].number_format = '0.00'
                elif "remunerados" in lbl:
                    col_p   = "E" if ano == 2025 else "F"
                    col_flag = "G" if ano == 2025 else "H"
                    ws[cel] = (
                        f'=SUMPRODUCT(--(Evid_Membros!$A$2:$A${max_rows["Evid_Membros"]}="{get_t(org)}"), '
                        f'Evid_Membros!${col_p}$2:${col_p}${max_rows["Evid_Membros"]}, '
                        f'--(Evid_Membros!${col_flag}$2:${col_flag}${max_rows["Evid_Membros"]}=1))'
                    )
                    ws[cel].number_format = '0.00'
                elif "R$" in lbl:
                    st = (f"Inicial {ano}" if "início" in lbl else
                          (f"Final {ano}" if "final" in lbl else
                           (f"Perdidas {ano}" if "Perdidas" in lbl else f"Exercidas {ano}")))
                    ws[cel] = (
                        f'=IFERROR(SUMPRODUCT('
                        f'--({aba_evid}!$B$2:$B${max_r}="{org}"), '
                        f'--({aba_evid}!$F$2:$F${max_r}="{st}"), '
                        f'{aba_evid}!$D$2:$D${max_r}, '
                        f'{aba_evid}!$E$2:$E${max_r}) / '
                        f'SUMIFS({aba_evid}!$E$2:$E${max_r}, {aba_evid}!$B$2:$B${max_r}, "{org}", '
                        f'{aba_evid}!$F$2:$F${max_r}, "{st}"), 0)'
                    )
                    ws[cel].number_format = 'R$ #,##0.00'
                elif "Diluição" in lbl:
                    ws[cel] = (
                        f'=IFERROR(SUMIFS({aba_evid}!$E$2:$E${max_r}, '
                        f'{aba_evid}!$B$2:$B${max_r}, "{org}", '
                        f'{aba_evid}!$F$2:$F${max_r}, "Final {ano}") / {capital_social}, 0)'
                    )
                    ws[cel].number_format = '0.0000%'
                else:
                    ws[cel] = "-"

    criar_quadro_mov('8.5_2025', df_evid_85, max_rows['Evid_85'], 2025)
    criar_quadro_mov('8.5_2026', df_evid_85, max_rows['Evid_85'], 2026)

    criar_quadro_mov('8.9_2025', df_evid_89, max_rows['Evid_89'], 2025, is_rsu=True)
    criar_quadro_mov('8.9_2026', df_evid_89, max_rows['Evid_89'], 2026, is_rsu=True)

    # 8.6 e 8.10
    def criar_quadro_outorga(num, df_e, max_r):
        ws = wb.create_sheet(f'Quadro_{num}')
        cols_rel = [c for c in df_e['Coluna_Relatorio'].unique() if pd.notnull(c)] if not df_e.empty else []
        ws.append(['Detalhamento das Outorgas'] + list(cols_rel))
        lbls = [
            "Nº total de membros", "N° de membros remunerados", "Data de outorga",
            "Quantidade outorgada", "Prazo para opções/ações", "Prazo máximo",
            "Prazo de restrição", "Valor justo na data", "Multiplicação"
        ]
        aba = f'Evid_{num.replace(".", "")}'
        # Flag de membros remunerados para 8.6 (Tem_86 = col I) e 8.10 (Tem_810 = col M)
        col_flag_rem = "I" if num == '8.6' else "M"

        for i, lbl in enumerate(lbls, start=2):
            ws.append([lbl])
            for c_idx, col_n in enumerate(cols_rel):
                cel = f"{get_col(c_idx + 2)}{i}"
                df_g = df_e[df_e['Coluna_Relatorio'] == col_n]
                if df_g.empty:
                    ws[cel] = "-"
                    continue

                org_val = df_g['Órgão Administrativo'].iloc[0]

                if "total de membros" in lbl:
                    ws[cel] = (
                        f'=SUMIF(Evid_Membros!$A$2:$A${max_rows["Evid_Membros"]}, '
                        f'"{get_t(org_val)}", Evid_Membros!$E$2:$E${max_rows["Evid_Membros"]})'
                    )
                    ws[cel].number_format = '0.00'
                elif "remunerados" in lbl:
                    ws[cel] = (
                        f'=SUMPRODUCT(--(Evid_Membros!$A$2:$A${max_rows["Evid_Membros"]}="{get_t(org_val)}"), '
                        f'Evid_Membros!$E$2:$E${max_rows["Evid_Membros"]}, '
                        f'--(Evid_Membros!${col_flag_rem}$2:${col_flag_rem}${max_rows["Evid_Membros"]}=1))'
                    )
                    ws[cel].number_format = '0.00'
                elif "Data de outorga" in lbl:
                    ws[cel] = df_g['Data Outorga'].iloc[0].strftime('%d/%m/%Y') if pd.notnull(df_g['Data Outorga'].iloc[0]) else "-"
                elif "Quantidade" in lbl:
                    col_q = "$G" if num == '8.10' else "$H"
                    ws[cel] = f'=SUMIF({aba}!$A$2:$A${max_r}, "{col_n}", {aba}!{col_q}$2:{col_q}${max_r})'
                    ws[cel].number_format = '#,##0'
                elif "Prazo para" in lbl:
                    ws[cel] = df_g['Data Carência'].iloc[0].strftime('%d/%m/%Y') if pd.notnull(df_g['Data Carência'].iloc[0]) else "-"
                elif "Prazo máximo" in lbl:
                    ws[cel] = df_g['Data Expiração'].iloc[0].strftime('%d/%m/%Y') if '8.6' in num and pd.notnull(df_g['Data Expiração'].iloc[0]) else "-"
                elif "restrição" in lbl:
                    ws[cel] = "N/A"
                elif "Valor justo" in lbl:
                    col_v, col_q = ("$H", "$G") if num == '8.10' else ("$I", "$H")
                    ws[cel] = (
                        f'=IFERROR(SUMPRODUCT(--({aba}!$A$2:$A${max_r}="{col_n}"), '
                        f'{aba}!{col_q}$2:{col_q}${max_r}, {aba}!{col_v}$2:{col_v}${max_r}) / '
                        f'SUMIF({aba}!$A$2:$A${max_r}, "{col_n}", {aba}!{col_q}$2:{col_q}${max_r}), 0)'
                    )
                    ws[cel].number_format = 'R$ #,##0.00'
                elif "Multiplicação" in lbl:
                    col_v, col_q = ("$H", "$G") if num == '8.10' else ("$I", "$H")
                    ws[cel] = (
                        f'=SUMPRODUCT(--({aba}!$A$2:$A${max_r}="{col_n}"), '
                        f'{aba}!{col_q}$2:{col_q}${max_r}, {aba}!{col_v}$2:{col_v}${max_r})'
                    )
                    ws[cel].number_format = 'R$ #,##0.00'

    criar_quadro_outorga('8.6',  df_evid_86, max_rows['Evid_86'])
    criar_quadro_outorga('8.10', df_evid_810, max_rows['Evid_810'])

    # 8.7
    ws_87 = wb.create_sheet('Quadro_8.7')
    ws_87.append(['2025'] + orgaos + ['Total'])
    lbls_87 = [
        "Nº total de membros", "N° de membros remunerados",
        "Opções não exercíveis", "Quantidade (Não exercíveis)", "Data em que se tornarão",
        "Prazo máximo (Não exercíveis)", "Restrição (Não)", "Preço médio (Não)", "Valor justo (Não)",
        "Opções exercíveis", "Quantidade (Exercíveis)", "Prazo máximo (Exercíveis)",
        "Restrição (Exercíveis)", "Preço médio (Exercíveis)", "Valor justo (Exercíveis)", "Valor justo do TOTAL"
    ]
    for i, lbl in enumerate(lbls_87, start=2):
        ws_87.append([lbl]); ws_87[f"A{i}"].alignment = Alignment(wrapText=True, vertical='top')
        for c_idx, org in enumerate(orgaos + ['Total']):
            cel = f"{get_col(c_idx + 2)}{i}"
            df_t = df_evid_87 if org == 'Total' else df_evid_87[df_evid_87['Órgão Administrativo'] == org]
            st = "Não exercível" if "Não" in lbl else "Exercível"
            mr = max_rows["Evid_87"]
            if "total de membros" in lbl:
                ws_87[cel] = (f'=SUM(Evid_Membros!$E$2:$E${max_rows["Evid_Membros"]})' if org == 'Total' else
                              f'=SUMIF(Evid_Membros!$A$2:$A${max_rows["Evid_Membros"]}, "{get_t(org)}", Evid_Membros!$E$2:$E${max_rows["Evid_Membros"]})')
                ws_87[cel].number_format = '0.00'
            elif "remunerados" in lbl:
                ws_87[cel] = (
                    f'=SUMPRODUCT(Evid_Membros!$E$2:$E${max_rows["Evid_Membros"]}, --(Evid_Membros!$J$2:$J${max_rows["Evid_Membros"]}=1))' if org == 'Total' else
                    f'=SUMPRODUCT(--(Evid_Membros!$A$2:$A${max_rows["Evid_Membros"]}="{get_t(org)}"), Evid_Membros!$E$2:$E${max_rows["Evid_Membros"]}, --(Evid_Membros!$J$2:$J${max_rows["Evid_Membros"]}=1))'
                )
                ws_87[cel].number_format = '0.00'
            elif "Quantidade" in lbl:
                ws_87[cel] = (f'=SUMIF(Evid_87!$D$2:$D${mr}, "{st}", Evid_87!$E$2:$E${mr})' if org == 'Total' else
                              f'=SUMIFS(Evid_87!$E$2:$E${mr}, Evid_87!$A$2:$A${mr}, "{org}", Evid_87!$D$2:$D${mr}, "{st}")')
                ws_87[cel].number_format = '#,##0'
            elif "Data" in lbl:
                ws_87[cel] = formatar_prazos(df_t[df_t['Status_Vesting'] == 'Não exercível'], 'Data_Carência')
            elif "Prazo" in lbl:
                ws_87[cel] = formatar_prazos(df_t[df_t['Status_Vesting'] == st], 'Data_Expiração')
            elif "Preço" in lbl or "Valor justo (" in lbl:
                col_v = "$F" if "Preço" in lbl else "$G"
                n = (f'SUMPRODUCT(--(Evid_87!$D$2:$D${mr}="{st}"), Evid_87!$E$2:$E${mr}, Evid_87!{col_v}$2:{col_v}${mr})' if org == 'Total' else
                     f'SUMPRODUCT(--(Evid_87!$A$2:$A${mr}="{org}"), --(Evid_87!$D$2:$D${mr}="{st}"), Evid_87!$E$2:$E${mr}, Evid_87!{col_v}$2:{col_v}${mr})')
                d = (f'SUMIF(Evid_87!$D$2:$D${mr}, "{st}", Evid_87!$E$2:$E${mr})' if org == 'Total' else
                     f'SUMIFS(Evid_87!$E$2:$E${mr}, Evid_87!$A$2:$A${mr}, "{org}", Evid_87!$D$2:$D${mr}, "{st}")')
                ws_87[cel] = f'=IFERROR({n} / {d}, 0)'
                ws_87[cel].number_format = 'R$ #,##0.00'
            elif "TOTAL" in lbl:
                ws_87[cel] = (f'=SUMPRODUCT(Evid_87!$E$2:$E${mr}, Evid_87!$G$2:$G${mr})' if org == 'Total' else
                              f'=SUMPRODUCT(--(Evid_87!$A$2:$A${mr}="{org}"), Evid_87!$E$2:$E${mr}, Evid_87!$G$2:$G${mr})')
                ws_87[cel].number_format = 'R$ #,##0.00'
            else:
                ws_87[cel] = "N/A" if "restrição" in lbl.lower() else ""

    # 8.8
    ws_88 = wb.create_sheet('Quadro_8.8')
    ws_88.append(['2025'] + orgaos + ['Total'])
    lbls_88 = ["Nº total de membros", "N° de membros remunerados", "Número de ações",
               "Preço médio de exercício", "Preço médio de mercado", "Multiplicação (Ganho Total)"]
    for i, lbl in enumerate(lbls_88, start=2):
        ws_88.append([lbl]); mr = max_rows["Evid_88"]
        for c_idx, org in enumerate(orgaos + ['Total']):
            cel = f"{get_col(c_idx + 2)}{i}"
            if "total de membros" in lbl:
                ws_88[cel] = (f'=SUM(Evid_Membros!$E$2:$E${max_rows["Evid_Membros"]})' if org == 'Total' else
                              f'=SUMIF(Evid_Membros!$A$2:$A${max_rows["Evid_Membros"]}, "{get_t(org)}", Evid_Membros!$E$2:$E${max_rows["Evid_Membros"]})')
                ws_88[cel].number_format = '0.00'
            elif "remunerados" in lbl:
                ws_88[cel] = (
                    f'=SUMPRODUCT(Evid_Membros!$E$2:$E${max_rows["Evid_Membros"]}, --(Evid_Membros!$K$2:$K${max_rows["Evid_Membros"]}=1))' if org == 'Total' else
                    f'=SUMPRODUCT(--(Evid_Membros!$A$2:$A${max_rows["Evid_Membros"]}="{get_t(org)}"), Evid_Membros!$E$2:$E${max_rows["Evid_Membros"]}, --(Evid_Membros!$K$2:$K${max_rows["Evid_Membros"]}=1))'
                )
                ws_88[cel].number_format = '0.00'
            elif "Número" in lbl:
                ws_88[cel] = (f'=SUM(Evid_88!$E$2:$E${mr})' if org == 'Total' else
                              f'=SUMIF(Evid_88!$A$2:$A${mr}, "{org}", Evid_88!$E$2:$E${mr})')
                ws_88[cel].number_format = '#,##0'
            elif "exercício" in lbl or "mercado" in lbl:
                col_v = "$F" if "exercício" in lbl else "$G"
                n = (f'SUMPRODUCT(Evid_88!$E$2:$E${mr}, Evid_88!{col_v}$2:{col_v}${mr})' if org == 'Total' else
                     f'SUMPRODUCT(--(Evid_88!$A$2:$A${mr}="{org}"), Evid_88!$E$2:$E${mr}, Evid_88!{col_v}$2:{col_v}${mr})')
                d = (f'SUM(Evid_88!$E$2:$E${mr})' if org == 'Total' else
                     f'SUMIF(Evid_88!$A$2:$A${mr}, "{org}", Evid_88!$E$2:$E${mr})')
                ws_88[cel] = f'=IFERROR({n} / {d}, 0)'
                ws_88[cel].number_format = 'R$ #,##0.00'
            elif "Multiplicação" in lbl:
                ws_88[cel] = (f'=SUM(Evid_88!$H$2:$H${mr})' if org == 'Total' else
                              f'=SUMIF(Evid_88!$A$2:$A${mr}, "{org}", Evid_88!$H$2:$H${mr})')
                ws_88[cel].number_format = 'R$ #,##0.00'

    # 8.11
    ws_811 = wb.create_sheet('Quadro_8.11')
    ws_811.append(['2025'] + orgaos + ['Total'])
    lbls_811 = [
        "Nº total de membros", "N° de membros remunerados", "Número de ações",
        "Preço médio ponderado de aquisição",
        "Preço médio ponderado de mercado das ações adquiridas",
        "Multiplicação do total das ações adquiridas",
        "N° de membros que eram Estatutários na data do exercício"  
    ]
    for i, lbl in enumerate(lbls_811, start=2):
        ws_811.append([lbl])
        ws_811[f"A{i}"].alignment = Alignment(wrapText=True, vertical='top')
        mr = max_rows["Evid_811"]
        for c_idx, org in enumerate(orgaos + ['Total']):
            cel = f"{get_col(c_idx + 2)}{i}"
            if "total de membros" in lbl:
                ws_811[cel] = (f'=SUM(Evid_Membros!$E$2:$E${max_rows["Evid_Membros"]})' if org == 'Total' else
                               f'=SUMIF(Evid_Membros!$A$2:$A${max_rows["Evid_Membros"]}, "{get_t(org)}", Evid_Membros!$E$2:$E${max_rows["Evid_Membros"]})')
                ws_811[cel].number_format = '0.00'
            elif "remunerados" in lbl:
                ws_811[cel] = (
                    f'=SUMPRODUCT(Evid_Membros!$E$2:$E${max_rows["Evid_Membros"]}, --(Evid_Membros!$N$2:$N${max_rows["Evid_Membros"]}=1))' if org == 'Total' else
                    f'=SUMPRODUCT(--(Evid_Membros!$A$2:$A${max_rows["Evid_Membros"]}="{get_t(org)}"), Evid_Membros!$E$2:$E${max_rows["Evid_Membros"]}, --(Evid_Membros!$N$2:$N${max_rows["Evid_Membros"]}=1))'
                )
                ws_811[cel].number_format = '0.00'
            elif "Número de ações" in lbl:
                ws_811[cel] = (f'=SUM(Evid_811!$E$2:$E${mr})' if org == 'Total' else
                               f'=SUMIF(Evid_811!$A$2:$A${mr}, "{org}", Evid_811!$E$2:$E${mr})')
                ws_811[cel].number_format = '#,##0'
            elif "Multiplicação" in lbl:
                ws_811[cel] = (f'=SUM(Evid_811!$I$2:$I${mr})' if org == 'Total' else
                               f'=SUMIF(Evid_811!$A$2:$A${mr}, "{org}", Evid_811!$I$2:$I${mr})')
                ws_811[cel].number_format = 'R$ #,##0.00'
            elif "aquisição" in lbl or "mercado" in lbl:
                col_v = "$F" if "aquisição" in lbl else "$G"
                n = (f'SUMPRODUCT(Evid_811!$E$2:$E${mr}, Evid_811!{col_v}$2:{col_v}${mr})' if org == 'Total' else
                     f'SUMPRODUCT(--(Evid_811!$A$2:$A${mr}="{org}"), Evid_811!$E$2:$E${mr}, Evid_811!{col_v}$2:{col_v}${mr})')
                d = (f'SUM(Evid_811!$E$2:$E${mr})' if org == 'Total' else
                     f'SUMIF(Evid_811!$A$2:$A${mr}, "{org}", Evid_811!$E$2:$E${mr})')
                ws_811[cel] = f'=IFERROR({n} / {d}, 0)'
                ws_811[cel].number_format = 'R$ #,##0.00'
            elif "Estatutários" in lbl:
                ws_811[cel] = (f'=COUNTIF(Evid_811!$H$2:$H${mr}, TRUE)' if org == 'Total' else
                               f'=COUNTIFS(Evid_811!$A$2:$A${mr}, "{org}", Evid_811!$H$2:$H${mr}, TRUE)')
                ws_811[cel].number_format = '0'

    # 8.12
    ws_812 = wb.create_sheet('Quadro_8.12')
    df_outorga['Data de Outorga'] = pd.to_datetime(df_outorga['Data de Outorga'], errors='coerce', dayfirst=True)

    df_25 = df_outorga[df_outorga['Data de Outorga'].dt.year < 2025].copy()

    if df_25.empty:
        ws_812.append(["Variável", "Nenhum programa em aberto no início de 2025"])
    else:
        df_25['Chave_Coluna'] = df_25['Programa'] + " (" + df_25['Órgão Administrativo'] + ")"
        col_progs = df_25['Chave_Coluna'].unique()
        ws_812.append(["Variável"] + list(col_progs))

        labels_812 = [
            "Modelo de Precificação", "Preço Médio Ponderado das Ações (R$)", "Preço de Exercício (R$)",
            "Volatilidade Esperada (%)", "Prazo de vida da opção", "Dividendos Esperados (%)",
            "Taxa de juros livre de riscos (%)", "Método utilizado (exercício antecipado)",
            "Forma de determinação da volatilidade", "Outra característica incorporada"
        ]
        for i, lbl in enumerate(labels_812, start=2):
            linha = [lbl]
            ws_812.column_dimensions['A'].width = 50
            ws_812[f"A{i}"].alignment = Alignment(wrapText=True, vertical='top')
            for col_n in col_progs:
                dados_prog = df_25[df_25['Chave_Coluna'] == col_n].iloc[0]

                if lbl == "Modelo de Precificação":
                    modelo_raw = pd.to_numeric(str(dados_prog.get('Model Options', '')).replace(',', '.'), errors='coerce')
                    val = MODEL_OPTIONS.get(int(modelo_raw), "A preencher") if pd.notnull(modelo_raw) else "A preencher"

                elif "Preço Médio Ponderado" in lbl:
                    v = pd.to_numeric(str(dados_prog.get('Preço da Ação / Opção', 0)).replace(',', '.'), errors='coerce')
                    val = f"R$ {v:,.2f}".replace('.', ',') if pd.notnull(v) and v != 0 else "A preencher"

                elif "Preço de Exercício" in lbl:
                    v = pd.to_numeric(str(dados_prog.get('Preço de Exercício na Outorga', 0)).replace(',', '.'), errors='coerce')
                    val = f"R$ {v:,.2f}".replace('.', ',') if pd.notnull(v) and v != 0 else "-"

                elif "Volatilidade Esperada" in lbl:
                    v = pd.to_numeric(str(dados_prog.get('Volatilidade', 0)).replace(',', '.'), errors='coerce')
                    val = f"{v*100:,.2f}%".replace('.', ',') if pd.notnull(v) and v != 0 else "A preencher"

                elif "Prazo de vida" in lbl:
                    dout = dados_prog.get('Data de Outorga')
                    dexp = pd.to_datetime(dados_prog.get('Data de Expiração'), errors='coerce', dayfirst=True)
                    if pd.notnull(dout) and pd.notnull(dexp):
                        # Se expiração = 9999, calcular vida em relação à data de carência
                        if dexp.year >= 9999:
                            dcar = pd.to_datetime(dados_prog.get('Data da Carência'), errors='coerce', dayfirst=True)
                            val = f"{(dcar - dout).days / 365.25:.1f} anos" if pd.notnull(dcar) else "A preencher"
                        else:
                            val = f"{(dexp - dout).days / 365.25:.1f} anos"
                    else:
                        val = "A preencher"

                elif "Dividendos Esperados" in lbl:
                    v = pd.to_numeric(str(dados_prog.get('Dividendos Esperados', 0)).replace(',', '.'), errors='coerce')
                    val = f"{v*100:,.2f}%".replace('.', ',') if pd.notnull(v) and v != 0 else "A preencher"

                elif "Taxa de juros" in lbl:
                    v = pd.to_numeric(str(dados_prog.get('Taxa de Juros Livre de Risco', 0)).replace(',', '.'), errors='coerce')
                    val = f"{v*100:,.2f}%".replace('.', ',') if pd.notnull(v) and v != 0 else "A preencher"

                elif "antecipado" in lbl:
                    v = dados_prog.get('Proporção de Exercício Antecipado', '')
                    val = str(v) if pd.notnull(v) and str(v).strip() != '' else "A preencher"

                elif "volatilidade" in lbl.lower():
                    vol_raw = pd.to_numeric(str(dados_prog.get('ModelVolatility', '')).replace(',', '.'), errors='coerce')
                    val = MODEL_VOLATILITY.get(int(vol_raw), "A preencher") if pd.notnull(vol_raw) else "A preencher"

                else:
                    val = "A preencher"

                linha.append(val)
            ws_812.append(linha)

    # Retorna o arquivo binário em memória
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output
