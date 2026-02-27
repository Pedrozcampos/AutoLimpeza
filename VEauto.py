import pandas as pd
import customtkinter as ctk
from tkinter import filedialog, messagebox
import re

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

progress_callback = None

class AutoCleanApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("AutoClean Contábil")
        self.geometry("650x550")

        self.progress_callback = progress_callback

        # --- interface do gui ---
        self.label = ctk.CTkLabel(self, text="Limpeza de Razão Contábil", font=("Roboto", 22, "bold"))
        self.label.pack(pady=30)
        
        self.btn_process = ctk.CTkButton(self, text="Selecionar Arquivo Excel ou CSV", 
                                        height=50, width=300, font=("Roboto", 14, "bold"),
                                        command=self.run_process)
        self.btn_process.pack(pady=30)
        
        self.status = ctk.CTkLabel(self, text="Status: Aguardando seleção...", text_color="gray")
        self.status.pack(pady=10)
        
        # - Barra de progresso
        self.progress_bar = ctk.CTkProgressBar(self, width=400)
        self.progress_bar.pack(pady=10)
        self.progress_bar.set(0)

    def update_progress(self, value, label_text=None):
        # - Atualiza o valor da barra
        self.progress_bar.set(value)

        # - Se um texto específico for enviado, usa ele. 
        # - Caso contrário, só atualiza a porcentagem se o processo já tiver começado (>0)
        if label_text:
            self.status.configure(text=label_text)
        elif value > 0 and value < 1.0:
            self.status.configure(text=f"Processando... {int(value*100)}%")
            
        self.update_idletasks()

    def extrair_conta(self, linha_texto):
        texto = str(linha_texto).replace('nan', '').strip()

        # - Busca sequência numérica (código).
        cod_match = re.search(r'(\d{5,})', texto)
        nome_match = re.search(r'Nome:\s*(.*)', texto, re.IGNORECASE)
        
        if cod_match and nome_match:
            return f"{cod_match.group(1)} - {nome_match.group(1).strip()}"
        return texto

    def limpar_num(self, val):
        # - Converte valores contábeis para float puro, lidando com diferentes formatos e casos de erro.

        if pd.isna(val) or str(val).strip() == "": return 0.0
        if isinstance(val, (int, float)): return float(val)
        # - Remove pontos de milhar e ajusta vírgula decimal.
        texto = re.sub(r'[^\d,]', '', str(val)).replace(',', '.')
        try: return float(texto)
        except: return 0.0

    def run_process(self):
        file_path = filedialog.askopenfilename(filetypes=[("Arquivos Suportados", "*.xlsx *.xls *.csv")])
        if not file_path: return

        try:
            # Inicia o visual do processamento apenas após o arquivo ser selecionado
            self.update_progress(0.05, "Status: Carregando arquivo...")
            
            # - Carregamento do arquivo.
            if file_path.endswith('.csv'):
                df_raw = pd.read_csv(file_path, sep=None, engine='python', encoding='utf-8-sig')
            else:
                df_raw = pd.read_excel(file_path)
            
            dados_finais = []
            conta_ativa = ""
            
            # - Mapeia colunas existentes.
            colunas_existentes = [c for c in df_raw.columns if "Unnamed" not in str(c)]
            
            total_linhas = len(df_raw)
            
            for i, row in df_raw.iterrows():
                # Atualiza a barra dinamicamente baseada nas linhas (0.1 a 0.8 do progresso)
                if i % 50 == 0:
                    prog_iter = 0.1 + (i / total_linhas) * 0.7
                    self.update_progress(prog_iter)
                    
                # - Transforma a linha em texto para detectar a conta.
                linha_texto = " ".join([str(v) for v in row.values if pd.notna(v)])
                
                # - Detecta nova conta no cabeçalho.
                if 'conta' in linha_texto.lower():
                    conta_ativa = self.extrair_conta(linha_texto)
                    continue
                
                # - Identifica se é uma linha de transação válida. (presença de data e números)
                data_val = row.get('Data')
                if pd.notna(data_val) and re.search(r'\d', str(data_val)):
                    # - Pula linhas de saldo anterior que costumam vir no meio.
                    if "saldo anterior" in linha_texto.lower(): continue
                    
                    # - Cria o dicionário da linha com TODAS as colunas que vieram no arquivo.
                    linha_dict = {col: row.get(col, "") for col in colunas_existentes}
                    
                    # - Injeta/Atualiza a coluna Conta e limpa valores.
                    linha_dict['Conta'] = conta_ativa
                    if 'Débito' in linha_dict: linha_dict['Débito'] = self.limpar_num(linha_dict['Débito'])
                    if 'Crédito' in linha_dict: linha_dict['Crédito'] = self.limpar_num(linha_dict['Crédito'])
                    
                    dados_finais.append(linha_dict)
                    
            self.update_progress(0.9, "Status: Organizando planilhas...")
            
            # --- Montagem do DataFrame Final ---
            df_res = pd.DataFrame(dados_finais)

            # - Garantia das 5 colunas obrigatórias.
            obrigatorias = ['Data', 'Crédito', 'Débito', 'Histórico', 'Conta']
            for col in obrigatorias:
                if col not in df_res.columns:
                    df_res[col] = ""

            # - Organiza: Obrigatórias primeiro, Extras depois.
            outras_cols = [c for c in df_res.columns if c not in obrigatorias]
            df_res = df_res[obrigatorias + outras_cols]
            
            # - Salva o resultado.
            save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", 
                                                    filetypes=[("Excel", "*.xlsx")])
            if save_path:
                df_res.to_excel(save_path, index=False)
                self.update_progress(1.0, "Status: Concluído!")
                messagebox.showinfo("Sucesso", "Arquivo limpo e processado com sucesso!")
            else:
                self.status.configure(text="Status: Operação cancelada", text_color="black")
                self.progress_bar.set(0)

        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro: {str(e)}")
            self.status.configure(text="Status: Erro no processamento", text_color="red")
            self.progress_bar.set(0)

    def atualizar_interface_progresso(self, valor):
        # - Atualiza a barra de progresso com segurança
        self.progress_bar.set(valor)
        self.status_label.configure(text=f"Progresso: {int(valor*100)}%")


if __name__ == "__main__":
    app = AutoCleanApp()
    app.mainloop()