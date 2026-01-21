import pdfplumber
import pandas as pd
import os
import re
import json

# --- CONFIGURAÇÕES ---
PASTA_ATUAL = os.path.dirname(os.path.abspath(__file__))
PASTA_PDFS = PASTA_ATUAL 
NOME_ARQUIVO_SAIDA = "Relatorio_Final_Data_Texto_v12.xlsx"
ARQUIVO_GABARITO_PDF = "RelatorioUnidades (1).pdf"
ARQUIVO_CACHE_JSON = "banco_dados_unidades.json"

def limpar_valor(valor_str):
    if not valor_str: return 0.0
    v = str(valor_str).replace('.', '').replace(',', '.')
    try: return float(v)
    except: return 0.0

def carregar_gabarito_inteligente(caminho_pdf, caminho_json):
    if os.path.exists(caminho_json):
        try:
            with open(caminho_json, 'r', encoding='utf-8') as f:
                return json.load(f)
        except: pass

    if not os.path.exists(caminho_pdf):
        print(f"[AVISO] Gabarito '{caminho_pdf}' não encontrado.")
        return {}

    print("--> Criando banco de dados das Unidades...")
    gabarito = {}
    codigo_atual = "534"
    
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
                            gabarito[unidade_limpa] = {
                                "codigo": codigo_atual,
                                "proprietario": nome_raw.strip()
                            }
    
    with open(caminho_json, 'w', encoding='utf-8') as f:
        json.dump(gabarito, f, ensure_ascii=False, indent=4)
    return gabarito

def extrair_dados_final(caminho_pdf, gabarito):
    dados_extraidos = []
    
    with pdfplumber.open(caminho_pdf) as pdf:
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
                
                # --- 1. RADAR DE UNIDADE ---
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
                                found_unit = True
                                break
                            else:
                                if len(prox_linha) > 3 and not re.search(r'\d{2}/\d{2}/\d{4}', prox_linha):
                                    buffer_nomes.append(prox_linha)
                    
                    if found_unit:
                        nome_no_boleto = " ".join(buffer_nomes) if buffer_nomes else "Na linha"

                # --- 2. DATA DE VENCIMENTO (CORRIGIDA) ---
                # Ignora linhas de "Período" ou "Emitido em"
                if "Período" not in linha and "Emitido" not in linha:
                    # Prioridade: Linha que começa com código do boleto (ex: 4978...)
                    match_boleto_header = re.search(r'^(\d{7,}).*?(\d{2}/\d{2}/\d{4})', linha)
                    
                    if match_boleto_header:
                        vencimento_atual = match_boleto_header.group(2)
                    
                    # Fallback: Se tiver data e valor financeiro na linha, pode ser vencimento
                    elif re.search(r'(\d{2}/\d{2}/\d{4})', linha):
                        temp_data = re.search(r'(\d{2}/\d{2}/\d{4})', linha).group(1)
                        if "1900" not in temp_data and re.search(r'[\d\.,]+$', linha):
                             pass # Mantém o vencimento que já pegou ou atualiza se tiver certeza

                # --- 3. ITENS ---
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
                        "condominio": codigo_final,
                        "bloco": "1",
                        "unidade": unidade_atual,
                        "vencimento": vencimento_atual if vencimento_atual else "Verificar",
                        " ": "", 
                        "cod conta contabil": conta, 
                        "descrição": descricao,
                        "  ": "", 
                        "valor": limpar_valor(valor_str),
                        "AUDITORIA_BOLETO": nome_no_boleto,
                        "AUDITORIA_GABARITO": nome_gabarito,
                        "STATUS": status_obs
                    })
                
                i += 1 

    return dados_extraidos

# --- EXECUÇÃO ---
print("--> Gerando Excel com DATAS EM TEXTO (DD/MM/AAAA)...")
caminho_gabarito = os.path.join(PASTA_PDFS, ARQUIVO_GABARITO_PDF)
caminho_json = os.path.join(PASTA_PDFS, ARQUIVO_CACHE_JSON)

db_unidades = carregar_gabarito_inteligente(caminho_gabarito, caminho_json)

todos_dados = []
arquivos_pdf = [f for f in os.listdir(PASTA_PDFS) if f.lower().endswith(".pdf")]

for arquivo in arquivos_pdf:
    if arquivo == ARQUIVO_GABARITO_PDF: continue
    print(f"Lendo: {arquivo}...")
    try:
        caminho = os.path.join(PASTA_PDFS, arquivo)
        dados = extrair_dados_final(caminho, db_unidades)
        todos_dados.extend(dados)
    except Exception as e:
        print(f"   [ERRO] {arquivo}: {e}")

if todos_dados:
    df = pd.DataFrame(todos_dados)
    
    # 1. Converte para data real primeiro (para garantir que é data)
    df['vencimento'] = pd.to_datetime(df['vencimento'], format='%d/%m/%Y', errors='coerce')
    
    # 2. CONVERTE PARA TEXTO FORMATADO (AQUI É O SEGREDO)
    # Isso força o Excel a mostrar "20/01/2026" como se fosse uma palavra, sem tentar formatar
    df['vencimento'] = df['vencimento'].dt.strftime('%d/%m/%Y').fillna("Verificar")
    
    df = df[df['valor'] > 0]
    
    # Ordenação das colunas
    colunas_ordem = ["condominio", "bloco", "unidade", "vencimento", " ", "cod conta contabil", "descrição", "  ", "valor", "AUDITORIA_BOLETO", "AUDITORIA_GABARITO", "STATUS"]
    df = df[colunas_ordem]

    with pd.ExcelWriter(NOME_ARQUIVO_SAIDA, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Importacao')
        workbook  = writer.book
        worksheet = writer.sheets['Importacao']
        
        fmt_num = workbook.add_format({'num_format': '#,##0.00'})
        
        # Ajuste de larguras
        worksheet.set_column('A:C', 10)
        worksheet.set_column('D:D', 18) # Data (agora é texto, não precisa de formato especial)
        worksheet.set_column('E:E', 2)
        worksheet.set_column('F:F', 15)
        worksheet.set_column('G:G', 40)
        worksheet.set_column('H:H', 2)
        worksheet.set_column('I:I', 15, fmt_num)
        worksheet.set_column('J:L', 25)
        
    print(f"\n[SUCESSO] Arquivo pronto: {NOME_ARQUIVO_SAIDA}")
else:
    print("\n[ERRO] Nenhum dado extraído.")