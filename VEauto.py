import pandas as pd
import numpy as np
import re

def limpar_extrato_contabil(caminho_arquivo, caminho_saida):
    """
    Processa arquivos Excel contábeis, extraindo contas de cabeçalhos 
    e normalizando as colunas de Débito, Crédito e Histórico.
    """
    
    # 1. Carregamento (Tratando possíveis variações de colunas)
    # Nota: Se for .xlsx, use engine='openpyxl'
    df_raw = pd.read_csv(caminho_arquivo) 
    
    # Lista para armazenar as linhas processadas
    dados_limpos = []
    conta_atual = None
    
    # 2. Processamento Linha a Linha (Lógica de Identificação de Conta)
    for _, linha in df_raw.iterrows():
        # Converte a linha inteira em string para buscar o padrão "Conta"
        linha_str = " ".join([str(val) for val in linha.values]).lower()
        
        # Identifica se a linha é um cabeçalho de nova conta
        if 'conta' in linha_str or 'conta analitica' in linha_str:
            # Extrai o nome da conta (ajuste o índice se a conta estiver em coluna específica)
            # Aqui pegamos o conteúdo após "Nome:" ou após o código da conta
            conta_atual = linha_str.strip() 
            continue
            
        # 3. Filtro de Transações
        # Só processa se tivermos uma conta ativa e se a linha parecer uma transação (tem data)
        data_valor = str(linha.get('Data', ''))
        if conta_atual and re.search(r'\d{4}-\d{2}-\d{2}', data_valor):
            nova_linha = {
                'Data': linha.get('Data'),
                'Lote': linha.get('Lote'),
                'Sq': linha.get('Sq'),
                'Contra Partida': linha.get('Contra Partida'),
                'Nro. Doc.': linha.get('Nro. Doc.'),
                'Histórico': linha.get('Histórico'),
                'Origem': linha.get('Origem'),
                'Débito': linha.get('Débito'),
                'Crédito': linha.get('Crédito'),
                'Conta': conta_atual
            }
            dados_limpos.append(nova_linha)

    # 4. Criação do DataFrame Final
    df_final = pd.DataFrame(dados_limpos)

    # 5. Higienização de Valores Numéricos
    # Garante que Débito e Crédito sejam números, preenchendo vazios com 0
    for col in ['Débito', 'Crédito']:
        if col in df_final.columns:
            df_final[col] = pd.to_numeric(df_final[col], errors='coerce').fillna(0)

    # 6. Exportação
    df_final.to_csv(caminho_saida, index=False, encoding='utf-8-sig')
    print(f"Processamento concluído! Arquivo salvo em: {caminho_saida}")

# Exemplo de uso
if __name__ == "__main__":
    arquivo_input = 'comoe.xlsx - Planilha1.csv'
    arquivo_output = 'extrato_limpo_final.csv'
    limpar_extrato_contabil(arquivo_input, arquivo_output)