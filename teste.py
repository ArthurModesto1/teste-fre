import pandas as pd

def criar_template_teste_completo():
    arquivo_saida = 'Template_Teste_Completo.xlsx'
    
    with pd.ExcelWriter(arquivo_saida, engine='openpyxl') as writer:
        
        # ==========================================
        # 1. MEMBROS (4 por Órgão com nomes fictícios)
        # ==========================================
        df_membros = pd.DataFrame({
            'Orgão': [
                'Diretoria Estatutária', 'Diretoria Estatutária', 'Diretoria Estatutária', 'Diretoria Estatutária',
                'Conselho de Administração', 'Conselho de Administração', 'Conselho de Administração', 'Conselho de Administração',
                'Conselho Fiscal', 'Conselho Fiscal', 'Conselho Fiscal', 'Conselho Fiscal'
            ],
            'CARGO': [
                'CEO', 'CFO', 'CTO', 'COO',
                'Presidente CA', 'Membro CA', 'Membro CA', 'Membro CA',
                'Presidente CF', 'Membro CF', 'Membro CF', 'Membro CF'
            ],
            'NOME COMPLETO': [
                'Bruce Wayne', 'Tony Stark', 'Lex Luthor', 'Norman Osborn',
                'Charles Xavier', 'Albus Dumbledore', 'Mestre Yoda', 'Gandalf o Cinzento',
                'Sherlock Holmes', 'Hercule Poirot', 'Hermione Granger', 'Senhor Spock'
            ],
            'CPF/CNPJ': [
                '111.111.111-01', '111.111.111-02', '111.111.111-03', '111.111.111-04',
                '222.222.222-01', '222.222.222-02', '222.222.222-03', '222.222.222-04',
                '333.333.333-01', '333.333.333-02', '333.333.333-03', '333.333.333-04'
            ],
            'DATA DE ENTRADA': ['01/01/2020'] * 12,
            # Simulando que o Senhor Spock saiu no meio do ano para testar a função Pro-Rata
            'DATA DE SAÍDA': ['31/12/2026'] * 11 + ['30/06/2025'], 
            'PRAZO DE MANDATO': ['2 anos'] * 12
        })
        df_membros.to_excel(writer, sheet_name='Membros', index=False)

        # ==========================================
        # 2. DADOS DA OUTORGA (Múltiplos Planos)
        # ==========================================
        df_outorga = pd.DataFrame({
            'Programa': ['SOP 2023', 'RS 2023', 'SOP 2024', 'RS 2024', 'SOP 2025', 'RS 2025', 'SOP 2023', 'RS 2024'],
            'Nome': ['Bruce Wayne', 'Bruce Wayne', 'Tony Stark', 'Tony Stark', 'Lex Luthor', 'Norman Osborn', 'Charles Xavier', 'Albus Dumbledore'],
            'CPF': ['111.111.111-01', '111.111.111-01', '111.111.111-02', '111.111.111-02', '111.111.111-03', '111.111.111-04', '222.222.222-01', '222.222.222-02'],
            'Órgão Administrativo': ['Diretoria Estatutária', 'Diretoria Estatutária', 'Diretoria Estatutária', 'Diretoria Estatutária', 'Diretoria Estatutária', 'Diretoria Estatutária', 'Conselho de Administração', 'Conselho de Administração'],
            'Tipo de Plano': ['Stock Options', 'Ações Restritas', 'Stock Options', 'Ações Restritas', 'Stock Options', 'Ações Restritas', 'Stock Options', 'Ações Restritas'],
            'Data de Outorga': ['01/01/2023', '01/01/2023', '01/01/2024', '01/01/2024', '01/01/2025', '01/01/2025', '01/01/2023', '01/01/2024'],
            'Data da Carência': ['01/01/2024', '01/01/2024', '01/01/2025', '01/01/2025', '01/01/2028', '01/01/2028', '01/01/2024', '01/01/2025'],
            'Data de Expiração': ['01/01/2030', '01/01/2030', '01/01/2031', '01/01/2031', '01/01/2032', '01/01/2032', '01/01/2030', '01/01/2031'],
            'Outorgado (original)': [10000, 5000, 12000, 6000, 15000, 7000, 8000, 4000],
            'Preço de Exercício na Outorga': [10.0, 0.0, 12.0, 0.0, 15.0, 0.0, 10.0, 0.0], # Ações Restritas geralmente tem preço de exercício = 0
            'Fair Value na Outorga': [2.5, 10.0, 3.0, 12.0, 4.0, 15.0, 2.5, 12.0],
            'Preço de Exercício Atual': [10.0, 0.0, 12.0, 0.0, 15.0, 0.0, 10.0, 0.0],
            'Fair Value Atualizado': [3.0, 11.0, 3.5, 13.0, 4.0, 15.0, 3.0, 13.0],
            'Preço da Ação / Opção': [22.0, 22.0, 24.0, 24.0, 26.0, 26.0, 22.0, 24.0],
            'Volatilidade': [0.35, 0.0, 0.38, 0.0, 0.40, 0.0, 0.35, 0.0],
            'Dividendos Esperados': [0.03, 0.0, 0.04, 0.0, 0.05, 0.0, 0.03, 0.0],
            'Taxa de Juros Livre de Risco': [0.10, 0.0, 0.11, 0.0, 0.12, 0.0, 0.10, 0.0],
            'Proporção de Exercício Antecipado': ['N/A', 'N/A', 'N/A', 'N/A', '10%', 'N/A', 'N/A', 'N/A']
        })
        df_outorga.to_excel(writer, sheet_name='Dados da outorga', index=False)

        # ==========================================
        # 3. HISTÓRICO DE MOVIMENTAÇÕES
        # ==========================================
        cols_mov = ['Programa', 'Plano', 'Nome', 'CPF', 'Órgão Administrativo', 'Data', 'Quantidade de Ações', 'Preço de Exercício', 'Preço da Ação (Mercado)', 'Status']
        df_mov = pd.DataFrame([
            # Exercício no ano passado (Garante que o "Saldo Inicial 2025" do Bruce Wayne será 8.000 ao invés de 10.000)
            ['SOP 2023', 'Stock Options', 'Bruce Wayne', '111.111.111-01', 'Diretoria Estatutária', '15/06/2024', 2000, 10.0, 18.0, 'Exercido'],
            
            # Movimentações correntes em 2025 (Exercícios, Entregas e Cancelamentos)
            ['SOP 2023', 'Stock Options', 'Bruce Wayne', '111.111.111-01', 'Diretoria Estatutária', '15/03/2025', 3000, 10.0, 25.0, 'Exercido'],
            ['RS 2023', 'Ações Restritas', 'Bruce Wayne', '111.111.111-01', 'Diretoria Estatutária', '10/05/2025', 1000, 0.0, 26.0, 'Entregue'],
            ['SOP 2024', 'Stock Options', 'Tony Stark', '111.111.111-02', 'Diretoria Estatutária', '20/04/2025', 4000, 12.0, 25.0, 'Exercido'],
            ['RS 2024', 'Ações Restritas', 'Tony Stark', '111.111.111-02', 'Diretoria Estatutária', '15/08/2025', 2000, 0.0, 27.0, 'Cancelado'],
            ['SOP 2023', 'Stock Options', 'Charles Xavier', '222.222.222-01', 'Conselho de Administração', '10/09/2025', 2000, 10.0, 28.0, 'Exercido'],
            ['RS 2024', 'Ações Restritas', 'Albus Dumbledore', '222.222.222-02', 'Conselho de Administração', '10/10/2025', 1000, 0.0, 29.0, 'Entregue']
        ], columns=cols_mov)
        
        # Simula a estrutura exata exigida pelo ETL: O cabeçalho real começa na linha 2.
        df_mov_excel = pd.DataFrame(columns=[f'Dummy_{i}' for i in range(len(cols_mov))])
        df_mov_excel.loc[0] = ['Extrato Oficial de Movimentações Emitido pelo Sistema RH'] + [''] * (len(cols_mov) - 1)
        df_mov_excel.loc[1] = cols_mov
        for i, row in df_mov.iterrows():
            df_mov_excel.loc[i+2] = row.values
        df_mov_excel.to_excel(writer, sheet_name='Histórico de movimentações', index=False, header=False)

        # ==========================================
        # 4. PREVISÃO OUTORGA 2026
        # ==========================================
        df_prev = pd.DataFrame([
            ['Previsão Outorga 2026 (Conselho de Administração)', ''],
            ['Quantidade a ser outorgada', 50000],
            ['Preço de Exercício', 28.50],
            ['', ''],
            ['Previsão Outorga 2026 (Diretoria Estatutária)', ''],
            ['Quantidade a ser outorgada', 120000],
            ['Preço de Exercício', 28.50]
        ])
        df_prev.to_excel(writer, sheet_name='Previsão outorga 2026', index=False, header=False)

    print(f"✅ Arquivo de Testes '{arquivo_saida}' com personagens gerado com sucesso!")

if __name__ == "__main__":
    criar_template_teste_completo()