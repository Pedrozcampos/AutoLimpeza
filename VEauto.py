import pandas as pd
import customtkinter as ctk
from tkinter import filedialog, messagebox
import re
import os

# Configuração visual da interface
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("AutoClean Contábil - Gemini Edition")
        self.geometry("500x300")

        # Elementos da Interface
        self.label = ctk.CTkLabel(self, text="Conversor de Extratos Contábeis", font=("Roboto", 20))
        self.label.pack(pady=20)

        self.btn_selecionar = ctk.CTkButton(self, text="Selecionar Arquivo (Excel/CSV)", command=self.processar_arquivo)
        self.btn_selecionar.pack(pady=10)

        self.status_label = ctk.CTkLabel(self, text="Aguardando arquivo...", text_color="gray")
        self.status_label.pack(pady=20)

    def processar_arquivo(self):
        caminho_input = filedialog.askopenfilename(
            filetypes=[("Arquivos de Excel", "*.xlsx *.xls"), ("Arquivos CSV", "*.csv")]
        )

        if not caminho_input:
            return

        try:
            self.status_label.configure(text="Processando... aguarde.", text_color="yellow")
            self.update()

            # 1. Carregamento inteligente
            if caminho_input.endswith('.csv'):
                df_raw = pd.read_csv(caminho_input, sep=None, engine='python', encoding='utf-8-sig')
            else:
                df_raw = pd.read_excel(caminho_input)

            dados_limpos = []
            conta_atual = "Não Identificada"

            # 2. Algoritmo de Extração e Limpeza
            for _, linha in df_raw.iterrows():
                # Transforma a linha em string para busca
                linha_texto = " ".join([str(v) for v in linha.values]).strip()
                linha_texto_lower = linha_texto.lower()

                # Identifica troca de conta (Conta ou Conta Analitica)
                if linha_texto_lower.startswith(('conta', 'conta analitica')):
                    conta_atual = linha_texto
                    continue

                # Pula linhas de saldo ou vazias
                if "saldo anterior" in linha_texto_lower or pd.isna(linha.get('Data')):
                    continue

                # Valida se a linha é uma transação (se tem data no formato esperado)
                data_str = str(linha.get('Data'))
                if re.search(r'\d', data_str): # Verifica se contém números (data)
                    nova_linha = {
                        'Data': linha.get('Data'),
                        'Lote': linha.get('Lote'),
                        'Sq': linha.get('Sq'),
                        'Contra Partida': linha.get('Contra Partida'),
                        'Nro. Doc.': linha.get('Nro. Doc.'),
                        'Histórico': linha.get('Histórico'),
                        'Origem': linha.get('Origem'),
                        'Débito': self.limpar_valor(linha.get('Débito')),
                        'Crédito': self.limpar_valor(linha.get('Crédito')),
                        'Conta': conta_atual
                    }
                    dados_limpos.append(nova_linha)

            # 3. Gerar DataFrame Final
            df_final = pd.DataFrame(dados_limpos)

            # 4. Salvar Resultado
            caminho_saida = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel workbook", "*.xlsx")],
                initialfile="Extrato_Limpo.xlsx"
            )

            if caminho_saida:
                df_final.to_excel(caminho_saida, index=False)
                self.status_label.configure(text="Sucesso! Arquivo salvo.", text_color="green")
                messagebox.showinfo("Concluído", f"Arquivo processado com sucesso!\nSalvo em: {caminho_saida}")
            else:
                self.status_label.configure(text="Operação cancelada.", text_color="white")

        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao processar: {str(e)}")
            self.status_label.configure(text="Erro no processamento.", text_color="red")

    def limpar_valor(self, valor):
        """Converte valores contábeis (string com vírgula) para float."""
        if pd.isna(valor) or valor == "" or valor == " ":
            return 0.0
        if isinstance(valor, str):
            valor = valor.replace('.', '').replace(',', '.')
            valor = re.sub(r'[^\d.]', '', valor)
        return float(valor) if valor else 0.0

if __name__ == "__main__":
    app = App()
    app.mainloop()