import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pdfplumber
import pandas as pd
import os
import re
from datetime import datetime

# --- LÓGICA DE EXTRAÇÃO (A MESMA QUE APROVAMOS NA V12) ---

def limpar_valor(valor_str):
    if not valor_str: return 0.0
    v = str(valor_str).replace('.', '').replace(',', '.')
    try: return float(v)
    except: return 0.0

def carregar_gabarito(caminho_pdf):
    gabarito = {}
    codigo_atual = "534"
    try:
        with pdfplumber.open(caminho_pdf) as pdf:
            for page in pdf.pages:
                text = page.extract_text() or ""
                linhas = text.split('\n')
                for linha in linhas:
                    if "Condomínio:" in linha and "-" in linha:
                        match_cod = re.search(r'Condomínio:\s*0?(\d{3})', linha)
                        if match_cod: codigo_atual = match_cod.group(1)
                    if "Unidade:" in linha:
                        parts = linha.split("Unidade:")
                        if len(parts) > 1:
                            conteudo = parts[1].strip()
                            if "-" in conteudo:
                                unidade_raw, nome_raw = conteudo.split("-", 1)
                                unidade_limpa = unidade_raw.strip().replace(" ", "")
                                gabarito[unidade_limpa] = {"codigo": codigo_atual, "proprietario": nome_raw.strip()}
        return gabarito
    except Exception as e:
        return {}

def processar_boletos(caminhos_boletos, gabarito):
    dados_extraidos = []
    
    for caminho in caminhos_boletos:
        with pdfplumber.open(caminho) as pdf:
            primeira_pag = pdf.pages[0].extract_text() or ""
            condominio_fallback = "534"
            if "LOJA" in primeira_pag.upper(): condominio_fallback = "536"
            elif "NR" in primeira_pag.upper(): condominio_fallback = "535"
            
            unidade_atual = ""
            nome_no_boleto = ""
            vencimento_atual = None
            
            for page in pdf.pages:
                text = page.extract_text()
                if not text: continue
                linhas = text.split('\n')
                total_linhas = len(linhas)
                i = 0
                while i < total_linhas:
                    linha = linhas[i].strip()
                    
                    # 1. RADAR UNIDADE
                    if "Unidade" in linha:
                        found_unit = False
                        buffer_nomes = []
                        match_same_line = re.search(r'(00\d{3,}|LOJA\s*\d+|[A-Z]+\d+)', linha, re.IGNORECASE)
                        if match_same_line:
                            unidade_atual = match_same_line.group(1).replace(" ", "")
                            found_unit = True
                        else:
                            for offset in range(1, 5):
                                if i + offset >= total_linhas: break
                                prox_linha = linhas[i+offset].strip()
                                match_prox = re.search(r'(00\d{3,}|LOJA\s*\d+|[A-Z]+\d+)', prox_linha, re.IGNORECASE)
                                if match_prox:
                                    unidade_atual = match_prox.group(1).replace(" ", "")
                                    found_unit = True; break
                                else:
                                    if len(prox_linha) > 3 and not re.search(r'\d{2}/\d{2}/\d{4}', prox_linha):
                                        buffer_nomes.append(prox_linha)
                        if found_unit:
                            nome_no_boleto = " ".join(buffer_nomes) if buffer_nomes else "Na linha"

                    # 2. DATA
                    if "Período" not in linha and "Emitido" not in linha:
                        match_boleto_header = re.search(r'^(\d{7,}).*?(\d{2}/\d{2}/\d{4})', linha)
                        if match_boleto_header:
                            vencimento_atual = match_boleto_header.group(2)
                        elif re.search(r'(\d{2}/\d{2}/\d{4})', linha):
                            temp_data = re.search(r'(\d{2}/\d{2}/\d{4})', linha).group(1)
                            if "1900" not in temp_data and re.search(r'[\d\.,]+$', linha): pass

                    # 3. ITENS
                    match_item = re.search(r'^(\d{4,5})\s+(.*?)\s+([\d\.,]+)$', linha)
                    if match_item:
                        conta = match_item.group(1)
                        descricao = match_item.group(2)
                        valor_str = match_item.group(3)
                        
                        if "Total" in descricao or "RESUMO" in descricao or "Empresa" in descricao: 
                            i += 1; continue
                        
                        codigo_final = condominio_fallback
                        nome_gabarito = "Não encontrado"
                        status_obs = "Atenção"
                        if unidade_atual in gabarito:
                            info = gabarito[unidade_atual]
                            codigo_final = info['codigo']
                            nome_gabarito = info['proprietario']
                            status_obs = "OK"
                        
                        dados_extraidos.append({
                            "condominio": codigo_final, "bloco": "1", "unidade": unidade_atual,
                            "vencimento": vencimento_atual if vencimento_atual else "Verificar",
                            " ": "", "cod conta contabil": conta, "descrição": descricao, "  ": "",
                            "valor": limpar_valor(valor_str),
                            "AUDITORIA_BOLETO": nome_no_boleto, "AUDITORIA_GABARITO": nome_gabarito, "STATUS": status_obs
                        })
                    i += 1
    return dados_extraidos

# --- INTERFACE GRÁFICA (TKINTER) ---

class LelloApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Robô Lello - Inadimplência")
        self.root.geometry("500x450")
        
        # Variáveis
        self.path_gabarito = ""
        self.files_boletos = []

        # Estilo
        style = ttk.Style()
        style.configure("TButton", padding=6, relief="flat", background="#ccc")
        
        # --- BLOCO 1: GABARITO ---
        lbl_step1 = tk.Label(root, text="1. Selecione o Relatório de Unidades (Gabarito)", font=("Arial", 10, "bold"))
        lbl_step1.pack(pady=(15, 5))
        
        self.btn_gabarito = ttk.Button(root, text="Selecionar Gabarito (PDF)", command=self.select_gabarito)
        self.btn_gabarito.pack()
        
        self.lbl_status_gabarito = tk.Label(root, text="Nenhum arquivo selecionado", fg="red", font=("Arial", 8))
        self.lbl_status_gabarito.pack(pady=5)

        # Divisor
        ttk.Separator(root, orient='horizontal').pack(fill='x', pady=10)

        # --- BLOCO 2: BOLETOS ---
        lbl_step2 = tk.Label(root, text="2. Selecione os Boletos de Inadimplência", font=("Arial", 10, "bold"))
        lbl_step2.pack(pady=(5, 5))
        
        self.btn_boletos = ttk.Button(root, text="Selecionar PDFs (Vários)", command=self.select_boletos)
        self.btn_boletos.pack()
        
        self.lbl_status_boletos = tk.Label(root, text="Nenhum arquivo selecionado", fg="red", font=("Arial", 8))
        self.lbl_status_boletos.pack(pady=5)

        # Divisor
        ttk.Separator(root, orient='horizontal').pack(fill='x', pady=10)

        # --- BLOCO 3: AÇÃO ---
        self.btn_processar = tk.Button(root, text="GERAR EXCEL", bg="green", fg="white", font=("Arial", 12, "bold"), height=2, width=20, command=self.run_process)
        self.btn_processar.pack(pady=20)
        
        self.lbl_final_status = tk.Label(root, text="Aguardando...", fg="blue")
        self.lbl_final_status.pack()

    def select_gabarito(self):
        file = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if file:
            self.path_gabarito = file
            self.lbl_status_gabarito.config(text=os.path.basename(file), fg="green")

    def select_boletos(self):
        files = filedialog.askopenfilenames(filetypes=[("PDF Files", "*.pdf")])
        if files:
            self.files_boletos = files
            self.lbl_status_boletos.config(text=f"{len(files)} arquivos selecionados", fg="green")

    def run_process(self):
        if not self.path_gabarito:
            messagebox.showwarning("Atenção", "Selecione o arquivo de Gabarito primeiro!")
            return
        if not self.files_boletos:
            messagebox.showwarning("Atenção", "Selecione os boletos de inadimplência!")
            return

        self.lbl_final_status.config(text="Processando... Aguarde...", fg="orange")
        self.root.update() # Força atualização da tela

        try:
            # 1. Carrega Gabarito
            gabarito = carregar_gabarito(self.path_gabarito)
            
            # 2. Processa
            dados = processar_boletos(self.files_boletos, gabarito)
            
            if dados:
                df = pd.DataFrame(dados)
                
                # Formata Data V12
                df['vencimento'] = pd.to_datetime(df['vencimento'], format='%d/%m/%Y', errors='coerce')
                df['vencimento'] = df['vencimento'].dt.strftime('%d/%m/%Y').fillna("Verificar")
                df = df[df['valor'] > 0]
                
                colunas = ["condominio", "bloco", "unidade", "vencimento", " ", "cod conta contabil", "descrição", "  ", "valor", "AUDITORIA_BOLETO", "AUDITORIA_GABARITO", "STATUS"]
                df = df[colunas]

                # Salvar
                nome_arquivo = f"Importacao_Lello_{datetime.now().strftime('%H%M%S')}.xlsx"
                pasta_destino = os.path.dirname(self.path_gabarito)
                caminho_salvar = os.path.join(pasta_destino, nome_arquivo)
                
                with pd.ExcelWriter(caminho_salvar, engine='xlsxwriter') as writer:
                    df.to_excel(writer, index=False, sheet_name='Importacao')
                    workbook = writer.book
                    worksheet = writer.sheets['Importacao']
                    fmt_num = workbook.add_format({'num_format': '#,##0.00'})
                    
                    # Layout
                    worksheet.set_column('A:C', 10)
                    worksheet.set_column('D:D', 15)
                    worksheet.set_column('E:E', 2)
                    worksheet.set_column('F:F', 15)
                    worksheet.set_column('G:G', 40)
                    worksheet.set_column('H:H', 2)
                    worksheet.set_column('I:I', 15, fmt_num)
                    worksheet.set_column('J:L', 20)

                self.lbl_final_status.config(text="Concluído!", fg="green")
                messagebox.showinfo("Sucesso", f"Arquivo gerado com sucesso em:\n{caminho_salvar}")
                os.startfile(pasta_destino)
            else:
                self.lbl_final_status.config(text="Nenhum dado encontrado", fg="red")
                messagebox.showwarning("Vazio", "O robô não encontrou dados nos PDFs enviados.")

        except Exception as e:
            self.lbl_final_status.config(text="Erro", fg="red")
            messagebox.showerror("Erro Fatal", str(e))

if __name__ == "__main__":
    root = tk.Tk()
    app = LelloApp(root)
    root.mainloop()