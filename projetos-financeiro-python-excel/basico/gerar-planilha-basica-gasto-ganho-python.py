import pandas as pd
import os

def gerar_planilha(transacoes, nome_arquivo='financas.xlsx', salvar_area_trabalho=True):
    """
    Gera um Excel com as transações e o saldo final.

    Parâmetros:
    - transacoes: lista de dicts com chaves 'Tipo', 'Categoria', 'Valor'.
    - nome_arquivo: nome do arquivo Excel a ser salvo (você pode trocar aqui ou no final).
    - salvar_area_trabalho: se True, salva na Área de Trabalho; se False, salva na pasta atual.
    """
    # Define o caminho final
    if salvar_area_trabalho:
        desktop = os.path.join(os.path.expanduser('~'), 'Desktop')  # Windows, Mac e Linux
        caminho = os.path.join(desktop, nome_arquivo)
    else:
        caminho = nome_arquivo  # salva na pasta atual do script

    df = pd.DataFrame(transacoes)
    df['Delta'] = df.apply(lambda x: x['Valor'] if x['Tipo']=='Ganho' else -x['Valor'], axis=1)
    saldo_final = df['Delta'].sum()
    df.drop(columns='Delta', inplace=True)

    with pd.ExcelWriter(caminho, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Transações', index=False)
        wb = writer.book
        ws = writer.sheets['Transações']
        fmt = wb.add_format({'num_format': 'R$ #,##0.00'})
        ws.set_column('C:C', 15, fmt)

        linha = len(df) + 2  # se quiser mudar a linha, altere aqui
        ws.write(f'B{linha}', 'Saldo final:')
        ws.write(f'C{linha}', saldo_final, fmt)

    print(f'Planilha salva em: {caminho}')

if __name__ == '__main__':
    # ← Se quiser outro nome, troque em "exemplo":
    nome_arquivo = 'exemplo.xlsx'
    transacoes = [
        {'Tipo': 'Ganho', 'Categoria': 'Salário',     'Valor': 5000.00},
        {'Tipo': 'Ganho', 'Categoria': 'Adiantamento', 'Valor': 600.00},
        {'Tipo': 'Gasto', 'Categoria': 'Supermercado', 'Valor': 600.50},
        {'Tipo': 'Ganho', 'Categoria': 'Freelance',    'Valor': 1200.00},
        {'Tipo': 'Gasto', 'Categoria': 'Transporte',   'Valor': 300.00},
        # pra adicionar mais linhas, copie daqui ↑
    ]

    # Se quiser salvar direto na Área de Trabalho, mantenha como True.
    gerar_planilha(transacoes, nome_arquivo, salvar_area_trabalho=True)
