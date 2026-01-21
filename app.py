import customtkinter as ctk
from tkinter import filedialog, messagebox
import pdfplumber
import pandas as pd
import os
import re
from datetime import datetime
import threading

# --- CONFIGURAÇÃO VISUAL ---
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

# --- LÓGICA V26 (TOTALMENTE CONFIGURÁVEL) ---

def limpar_valor(valor_str):
    if not valor_str: return 0.0
    v = str(valor_str).replace('.', '').replace(',', '.')
    try: return float(v)
    except: return 0.0

def formatar_unidade_custom(texto_unidade, modo_formatacao):
    """
    Formata a unidade conforme a escolha do usuário na interface.
    """
    texto = texto_unidade.upper().replace(" ", "")
    
    # Se for LOJA, mantém o padrão LJ + 4 digitos
    if "LOJA" in texto or "LJ" in texto:
        numeros = re.findall(r'\d+', texto)
        if numeros:
            return f"LJ{numeros[0].zfill(4)}"
        return texto.replace("LOJA", "LJ")
    
    # Se for APARTAMENTO/CASA
    nums = re.findall(r'\d+', texto)
    if not nums: return texto 
    
    numero_limpo = nums[0] 
    
    if modo_formatacao == "Limpo (602)":
        return numero_limpo.lstrip("0")
    elif modo_formatacao == "3 Dígitos (006)":
        return numero_limpo.zfill(3)
    elif modo_formatacao == "4 Dígitos (0602)":
        return numero_limpo.zfill(4)
    elif modo_formatacao == "6 Dígitos (000602)":
        return numero_limpo.zfill(6)
    else: # "Original do PDF"
        return texto

def formatar_bloco_custom(bloco_original, modo_bloco):
    if modo_bloco == "Fixo: 1":
        return "1"
    elif modo_bloco == "Fixo: 01":
        return "01"
    else: # "Original do PDF"
        return bloco_original

def traduzir_conta_contabil(descricao):
    """
    Mapeamento baseado no arquivo validado.
    """
    desc = descricao.upper().strip()
    
    if "GAS" in desc or "GÁS" in desc: return "472"
    if "AGUA" in desc or "ÁGUA" in desc: return "470"   
    if "ENERGIA" in desc or "LUZ" in desc: return "2823"
    if "FUNDO" in desc and "RESERVA" in desc: return "13" 
    if "IPTU" in desc: return "596"
    if "SEGURANCA" in desc or "LAUDO" in desc or "SISTEMA" in desc: return "597"
    if "EVENTUAIS" in desc: return "497"
    if "MULTA" in desc: return "14"
    if "JUROS" in desc: return "13"
    if "ATUALIZA" in desc: return "15"
    
    return "3" 

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
                                chave = unidade_raw.strip().upper().replace(" ", "")
                                gabarito[chave] = {
                                    "codigo": codigo_atual, 
                                    "proprietario": nome_raw.strip(),
                                    "unidade_raw": unidade_raw.strip().upper() 
                                }
        return gabarito
    except Exception as e:
        return {}

def processar_boletos(caminhos_boletos, gabarito, config_unidade, config_bloco, conta_bancaria):
    dados_extraidos = []
    
    for caminho in caminhos_boletos:
        with pdfplumber.open(caminho) as pdf:
            primeira_pag = pdf.pages[0].extract_text() or ""
            condominio_fallback = "534"
            if "LOJA" in primeira_pag.upper(): condominio_fallback = "536"
            elif "NR" in primeira_pag.upper(): condominio_fallback = "535"
            
            unidade_atual_chave = ""
            unidade_para_excel = ""
            bloco_lido = "1"
            vencimento_atual = None
            ignorar_unidade = False
            
            for page in pdf.pages:
                text = page.extract_text()
                if not text: continue
                linhas = text.split('\n')
                total_linhas = len(linhas)
                i = 0
                while i < total_linhas:
                    linha = linhas[i].strip()
                    
                    # 1. IDENTIFICAÇÃO DA UNIDADE
                    if "Unidade" in linha:
                        found_unit = False
                        match = re.search(r'(00\d{3,}|LOJA\s*\d+|[A-Z]+\d+)', linha, re.IGNORECASE)
                        
                        raw_lido = ""
                        if match:
                            raw_lido = match.group(1).upper()
                            found_unit = True
                        else:
                            for offset in range(1, 5):
                                if i + offset >= total_linhas: break
                                prox = linhas[i+offset].strip()
                                match_prox = re.search(r'(00\d{3,}|LOJA\s*\d+|[A-Z]+\d+)', prox, re.IGNORECASE)
                                if match_prox:
                                    raw_lido = match_prox.group(1).upper()
                                    found_unit = True; break
                        
                        if found_unit:
                            unidade_atual_chave = raw_lido.replace(" ", "")
                            unidade_para_excel = formatar_unidade_custom(raw_lido, config_unidade)
                            
                            if "LOJA" in raw_lido or "LJ" in raw_lido:
                                ignorar_unidade = True 
                            else:
                                ignorar_unidade = False

                    # 2. DATA
                    if "Período" not in linha and "Emitido" not in linha:
                        match_header = re.search(r'^(\d{7,}).*?(\d{2}/\d{2}/\d{4})', linha)
                        if match_header:
                            vencimento_atual = match_header.group(2)
                        elif re.search(r'(\d{2}/\d{2}/\d{4})', linha):
                            temp = re.search(r'(\d{2}/\d{2}/\d{4})', linha).group(1)
                            if "1900" not in temp and re.search(r'[\d\.,]+$', linha): pass

                    # 3. ITENS FINANCEIROS
                    match_item = re.search(r'^(\d{4,5})\s+(.*?)\s+([\d\.,]+)$', linha)
                    if match_item:
                        if ignorar_unidade: 
                            i += 1; continue

                        descricao = match_item.group(2)
                        valor_str = match_item.group(3)
                        
                        if "Total" in descricao or "RESUMO" in descricao or "Empresa" in descricao: 
                            i += 1; continue
                        
                        codigo_final = condominio_fallback
                        if unidade_atual_chave in gabarito:
                            info = gabarito[unidade_atual_chave]
                            codigo_final = info['codigo']
                        
                        codigo_ahreas = traduzir_conta_contabil(descricao)
                        bloco_final = formatar_bloco_custom(bloco_lido, config_bloco)

                        dados_extraidos.append({
                            "Cód. Condomínio": codigo_final, 
                            "Cód. Bloco": bloco_final,        
                            "Cód. Unidade": unidade_para_excel,
                            "Vencimento": vencimento_atual if vencimento_atual else "",
                            "Cód. Conta Bancária": conta_bancaria, # <--- Usa o valor configurado
                            "Cód. Conta Contábil": codigo_ahreas,
                            "Descrição": descricao, 
                            "Complemento": "",                 
                            "Valor": limpar_valor(valor_str),
                            "Percentual Multa": "",            
                            "Nro. Bancário": ""                
                        })
                    i += 1
    return dados_extraidos

# --- INTERFACE MODERNA ---

class LelloAppModerno(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Robô Lello - V26 (Total Config)")
        self.geometry("850x650") 
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.path_gabarito = ""
        self.files_boletos = []

        # --- MENU LATERAL (CONFIGURAÇÕES) ---
        self.frame_menu = ctk.CTkFrame(self, width=220, corner_radius=0)
        self.frame_menu.grid(row=0, column=0, sticky="nsew")
        self.frame_menu.grid_rowconfigure(8, weight=1)

        self.lbl_menu = ctk.CTkLabel(self.frame_menu, text="Configurações", font=("Roboto", 20, "bold"))
        self.lbl_menu.grid(row=0, column=0, padx=20, pady=(30, 20))

        # 1. Conta Bancária
        self.lbl_conta = ctk.CTkLabel(self.frame_menu, text="Conta Bancária (Fixo):", anchor="w")
        self.lbl_conta.grid(row=1, column=0, padx=20, pady=(10, 0), sticky="ew")
        self.entry_conta = ctk.CTkEntry(self.frame_menu, placeholder_text="Ex: 6000")
        self.entry_conta.grid(row=2, column=0, padx=20, pady=(0, 15), sticky="ew")
        self.entry_conta.insert(0, "6000") # Valor padrão

        # 2. Formato Unidade
        self.lbl_unidade = ctk.CTkLabel(self.frame_menu, text="Formato da Unidade:", anchor="w")
        self.lbl_unidade.grid(row=3, column=0, padx=20, pady=(10, 0), sticky="ew")
        self.option_unidade = ctk.CTkOptionMenu(self.frame_menu, values=["6 Dígitos (000602)", "Limpo (602)", "3 Dígitos (006)", "4 Dígitos (0602)", "Original do PDF"])
        self.option_unidade.grid(row=4, column=0, padx=20, pady=(0, 15), sticky="ew")
        self.option_unidade.set("6 Dígitos (000602)")

        # 3. Formato Bloco
        self.lbl_bloco = ctk.CTkLabel(self.frame_menu, text="Formato do Bloco:", anchor="w")
        self.lbl_bloco.grid(row=5, column=0, padx=20, pady=(10, 0), sticky="ew")
        self.option_bloco = ctk.CTkOptionMenu(self.frame_menu, values=["Fixo: 01", "Fixo: 1", "Original do PDF"])
        self.option_bloco.grid(row=6, column=0, padx=20, pady=(0, 10), sticky="ew")
        self.option_bloco.set("Fixo: 01")

        # --- ÁREA PRINCIPAL ---
        self.frame_main = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.frame_main.grid(row=0, column=1, sticky="nsew")
        self.frame_main.grid_columnconfigure(0, weight=1)

        self.lbl_titulo = ctk.CTkLabel(self.frame_main, text="Extrator Lello - V26", font=("Roboto", 24, "bold"))
        self.lbl_titulo.grid(row=0, column=0, pady=(30, 10), sticky="ew")
        self.lbl_sub = ctk.CTkLabel(self.frame_main, text="Configure a conta e formatos no menu lateral", font=("Roboto", 14), text_color="gray")
        self.lbl_sub.grid(row=1, column=0, pady=(0, 20), sticky="ew")

        # Seletores
        self.frame_gabarito = ctk.CTkFrame(self.frame_main)
        self.frame_gabarito.grid(row=2, column=0, padx=20, pady=10, sticky="ew")
        self.frame_gabarito.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(self.frame_gabarito, text="1. Relatório de Unidades (Gabarito)", font=("Roboto", 14, "bold")).grid(row=0, column=0, pady=10)
        self.btn_gabarito = ctk.CTkButton(self.frame_gabarito, text="Selecionar PDF", command=self.select_gabarito, height=40)
        self.btn_gabarito.grid(row=1, column=0, pady=10, padx=20, sticky="ew")
        self.lbl_status_gabarito = ctk.CTkLabel(self.frame_gabarito, text="Pendente", text_color="#FF5555")
        self.lbl_status_gabarito.grid(row=2, column=0, pady=(0, 10))

        self.frame_boletos = ctk.CTkFrame(self.frame_main)
        self.frame_boletos.grid(row=3, column=0, padx=20, pady=10, sticky="ew")
        self.frame_boletos.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(self.frame_boletos, text="2. PDFs de Inadimplência", font=("Roboto", 14, "bold")).grid(row=0, column=0, pady=10)
        self.btn_boletos = ctk.CTkButton(self.frame_boletos, text="Selecionar Boletos", command=self.select_boletos, fg_color="transparent", border_width=2, text_color=("gray10", "#DCE4EE"), height=40)
        self.btn_boletos.grid(row=1, column=0, pady=10, padx=20, sticky="ew")
        self.lbl_status_boletos = ctk.CTkLabel(self.frame_boletos, text="Pendente", text_color="#FF5555")
        self.lbl_status_boletos.grid(row=2, column=0, pady=(0, 10))

        self.btn_processar = ctk.CTkButton(self.frame_main, text="GERAR EXCEL", command=self.start_thread, height=50, font=("Roboto", 16, "bold"), fg_color="#2CC985", hover_color="#229965")
        self.btn_processar.grid(row=4, column=0, padx=20, pady=30, sticky="ew")
        self.barra_progresso = ctk.CTkProgressBar(self.frame_main, mode="indeterminate")

    def select_gabarito(self):
        file = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if file:
            self.path_gabarito = file
            self.lbl_status_gabarito.configure(text=f"✅ {os.path.basename(file)}", text_color="#55FF55")

    def select_boletos(self):
        files = filedialog.askopenfilenames(filetypes=[("PDF Files", "*.pdf")])
        if files:
            self.files_boletos = files
            self.lbl_status_boletos.configure(text=f"✅ {len(files)} arquivos selecionados", text_color="#55FF55")

    def start_thread(self):
        self.btn_processar.configure(state="disabled", text="Processando...")
        self.barra_progresso.grid(row=5, column=0, padx=20, pady=(0, 20), sticky="ew")
        self.barra_progresso.start()
        thread = threading.Thread(target=self.run_process)
        thread.start()

    def run_process(self):
        if not self.path_gabarito or not self.files_boletos:
            messagebox.showwarning("Atenção", "Selecione todos os arquivos!")
            self.reset_ui(); return

        try:
            # PEGA AS CONFIGURAÇÕES
            conf_unidade = self.option_unidade.get()
            conf_bloco = self.option_bloco.get()
            conf_conta = self.entry_conta.get().strip() # Pega o valor da conta bancária

            gabarito = carregar_gabarito(self.path_gabarito)
            # Passa a conta bancária para a função
            dados = processar_boletos(self.files_boletos, gabarito, conf_unidade, conf_bloco, conf_conta)
            
            if dados:
                df = pd.DataFrame(dados)
                df['Vencimento'] = pd.to_datetime(df['Vencimento'], format='%d/%m/%Y', errors='coerce')
                df['Vencimento'] = df['Vencimento'].dt.strftime('%d/%m/%Y').fillna("")
                df = df[df['Valor'] > 0]
                
                colunas_finais = [
                    "Cód. Condomínio", "Cód. Bloco", "Cód. Unidade", "Vencimento", 
                    "Cód. Conta Bancária", "Cód. Conta Contábil", "Descrição", 
                    "Complemento", "Valor", "Percentual Multa", "Nro. Bancário"
                ]
                df = df[colunas_finais]

                nome_arquivo = f"Importacao_Lello_{datetime.now().strftime('%H%M%S')}.xlsx"
                pasta_destino = os.path.dirname(self.files_boletos[0])
                caminho_salvar = os.path.join(pasta_destino, nome_arquivo)
                
                with pd.ExcelWriter(caminho_salvar, engine='xlsxwriter') as writer:
                    df.to_excel(writer, index=False, sheet_name='Inadimplência')
                    workbook = writer.book
                    worksheet = writer.sheets['Inadimplência']
                    fmt_num = workbook.add_format({'num_format': '#,##0.00'})
                    
                    worksheet.set_column('A:K', 15)
                    worksheet.set_column('G:G', 40)
                    worksheet.set_column('I:I', 15, fmt_num)

                messagebox.showinfo("Sucesso", f"Arquivo Gerado!\n{caminho_salvar}")
                try: os.startfile(pasta_destino)
                except: pass
            else:
                messagebox.showwarning("Vazio", "Nenhum dado encontrado.")

        except Exception as e:
            messagebox.showerror("Erro", str(e))
        
        self.reset_ui()

    def reset_ui(self):
        self.barra_progresso.stop()
        self.barra_progresso.grid_forget()
        self.btn_processar.configure(state="normal", text="GERAR EXCEL")

if __name__ == "__main__":
    app = LelloAppModerno()
    app.mainloop()