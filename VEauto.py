import pandas as pd
import customtkinter as ctk
from tkinter import filedialog, messagebox
import re

# Configuração visual do CustomTkinter
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class AutoCleanApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("AutoClean Contábil Pro")
        self.geometry("600x500")
        
        # Elementos da Interface
        self.label = ctk.CTkLabel(self, text="Limpeza de Razão Contábil", font=("Roboto", 22, "bold"))
        self.label.pack(pady=30)

        self.btn_process = ctk.CTkButton(self, text="Selecionar Arquivo Excel ou CSV", 
                                        height=50, width=300, font=("Roboto", 14, "bold"),
                                        command=self.run_process)
        self.btn_process.pack(pady=30)

        self.status = ctk.CTkLabel(self, text="Status: Aguardando seleção...", text_color="gray")
        self.status.pack(pady=10)

    def extrair_conta(self, linha_texto):
        """Limpa a string da conta para o formato 'Código - Nome'."""
        texto = str(linha_texto).replace('nan', '').strip()
        # Busca sequência numérica (código) e o que vem após 'Nome:'
        cod_match = re.search(r'(\d{5,})', texto)
        nome_match = re.search(r'Nome:\s*(.*)', texto, re.IGNORECASE)
        
        if cod_match and nome_match:
            return f"{cod_match.group(1)} - {nome_match.group(1).strip()}"
        return texto

    def limpar_num(self, val):
        """Converte valores contábeis para float puro."""
        if pd.isna(val) or str(val).strip() == "": return 0.0
        if isinstance(val, (int, float)): return float(val)
        # Remove pontos de milhar e ajusta vírgula decimal
        texto = re.sub(r'[^\d,]', '', str(val)).replace(',', '.')
        try: return float(texto)
        except: return 0.0

    def run_process(self):
        file_path = filedialog.askopenfilename(filetypes=[("Arquivos Suportados", "*.xlsx *.xls *.csv")])
        if not file_path: return

        try:
            self.status.configure(text="Processando dados...", text_color="gray")
            self.update()

            # Carregamento do arquivo
            if file_path.endswith('.csv'):
                df_raw = pd.read_csv(file_path, sep=None, engine='python', encoding='utf-8-sig')
            else:
                df_raw = pd.read_excel(file_path)
            
            dados_finais = []
            conta_ativa = ""
            
            # Mapeia colunas existentes (ignora colunas fantasmas do Excel)
            colunas_existentes = [c for c in df_raw.columns if "Unnamed" not in str(c)]
            
            for _, row in df_raw.iterrows():
                # Transforma a linha em texto para detectar a conta
                linha_texto = " ".join([str(v) for v in row.values if pd.notna(v)])
                
                # Detecta nova conta no cabeçalho
                if 'conta' in linha_texto.lower():
                    conta_ativa = self.extrair_conta(linha_texto)
                    continue
                
                # Identifica se é uma linha de transação (se a coluna Data tem número)
                data_val = row.get('Data')
                if pd.notna(data_val) and re.search(r'\d', str(data_val)):
                    # Pula linhas de saldo anterior que costumam vir no meio
                    if "saldo anterior" in linha_texto.lower(): continue
                    
                    # Cria o dicionário da linha com TODAS as colunas que vieram no arquivo
                    linha_dict = {col: row.get(col, "") for col in colunas_existentes}
                    
                    # Injeta/Atualiza a coluna Conta e limpa valores
                    linha_dict['Conta'] = conta_ativa
                    if 'Débito' in linha_dict: linha_dict['Débito'] = self.limpar_num(linha_dict['Débito'])
                    if 'Crédito' in linha_dict: linha_dict['Crédito'] = self.limpar_num(linha_dict['Crédito'])
                    
                    dados_finais.append(linha_dict)

            # Montagem do DataFrame Final
            df_res = pd.DataFrame(dados_finais)

            # Garantia das 5 colunas obrigatórias
            obrigatorias = ['Data', 'Crédito', 'Débito', 'Histórico', 'Conta']
            for col in obrigatorias:
                if col not in df_res.columns:
                    df_res[col] = ""

            # Organiza: Obrigatórias primeiro, Extras depois
            outras_cols = [c for c in df_res.columns if c not in obrigatorias]
            df_res = df_res[obrigatorias + outras_cols]

            # Salvar
            save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", 
                                                    filetypes=[("Excel", "*.xlsx")])
            if save_path:
                df_res.to_excel(save_path, index=False)
                messagebox.showinfo("Sucesso", "Arquivo limpo e processado com sucesso!")
                self.status.configure(text="Concluído!", text_color="green")
            else:
                self.status.configure(text="Operação cancelada", text_color="white")

        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro: {str(e)}")
            self.status.configure(text="Erro no processamento", text_color="red")

if __name__ == "__main__":
    app = AutoCleanApp()
    app.mainloop()