import io
import datetime
import unicodedata
import re
from flask import Flask, request, send_file, render_template
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Alignment, Border, Side, Font
from openpyxl.utils.dataframe import dataframe_to_rows

# --- IMPORTAÇÕES DE COMPATIBILIDADE ---
try:
    import xlrd
    from xlrd import xldate
except ImportError:
    xlrd = None

try:
    import pandas as pd
    import lxml
except ImportError:
    pd = None

app = Flask(__name__, template_folder='.', static_folder='.', static_url_path='')

# --- 1. FUNÇÕES UTILITÁRIAS (TEXTO E DATA) ---

def normalizar(texto):
    """Remove acentos e deixa maiúsculo (ex: 'LANÇAMENTO' -> 'LANCAMENTO')"""
    if texto is None: return ""
    texto = str(texto).strip()
    texto = unicodedata.normalize("NFD", texto)
    return ''.join(c for c in texto if unicodedata.category(c) != "Mn").upper()

def separar_duas_linhas(texto):
    """
    Separa o texto da célula que contém duas informações.
    Ex: "Nome da Pessoa\nCód: 123" -> Retorna ("Nome da Pessoa", "Cód: 123")
    """
    if not texto: return "", ""
    txt = str(texto).strip()
    
    if "\n" in txt:
        partes = txt.split("\n")
        return partes[0].strip(), partes[1].strip()
    
    # Fallback: Se não tiver quebra de linha, retorna tudo na primeira parte
    return txt, ""

def limpar_valor_monetario(valor):
    """Converte 'R$ 1.234,56' para float 1234.56"""
    if isinstance(valor, (int, float)):
        return float(valor)
    
    if isinstance(valor, str):
        # Remove R$, espaços e pontos de milhar
        limpo = valor.replace("R$", "").replace(" ", "").replace(".", "")
        # Troca vírgula decimal por ponto
        limpo = limpo.replace(",", ".")
        try:
            return float(limpo)
        except ValueError:
            return 0.0
    return 0.0

def converter_data_robusta(valor):
    """Tenta converter string ou data para objeto datetime"""
    if not valor: return None
    
    if isinstance(valor, (datetime.datetime, datetime.date)):
        return valor
        
    # Lista de formatos possíveis
    formatos = [
        "%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", 
        "%Y/%m/%d", "%d/%m/%y", "%m-%Y", "%m/%Y"
    ]
    
    valor_str = str(valor).strip().split(' ')[0] # Remove hora se houver
    
    for fmt in formatos:
        try:
            return datetime.datetime.strptime(valor_str, fmt)
        except ValueError:
            continue
            
    return None

def mes_abreviado(data_obj):
    if not data_obj: return ""
    meses = {1:"jan", 2:"fev", 3:"mar", 4:"abr", 5:"mai", 6:"jun",
             7:"jul", 8:"ago", 9:"set", 10:"out", 11:"nov", 12:"dez"}
    return f"{meses.get(data_obj.month, '')}-{str(data_obj.year)[-2:]}"

# --- 2. FUNÇÕES DE CONVERSÃO DE ARQUIVO ---

def is_html_file(file_stream):
    """Detecta se é um 'fake XLS' (HTML)"""
    pos = file_stream.tell()
    file_stream.seek(0)
    try:
        head = file_stream.read(2048)
        if b'<html' in head or b'<table' in head or b'\xef\xbb\xbf' in head:
            return True
    except: pass
    finally:
        file_stream.seek(pos)
    return False

def convert_html_to_xlsx(file_stream):
    """
    Converte HTML para XLSX focando em achar o cabeçalho 'LANÇAMENTO'.
    """
    if pd is None: raise EnvironmentError("Instale pandas e lxml")
    
    file_stream.seek(0)
    # header=None obriga o pandas a ler TUDO como dados, sem adivinhar cabeçalho
    try:
        dfs = pd.read_html(file_stream.read().decode('utf-8', errors='ignore'), header=None)
    except ValueError: raise ValueError("Nenhuma tabela encontrada no HTML")
    
    if not dfs: raise ValueError("Arquivo HTML vazio")
    
    # Pega a tabela com mais colunas
    df = max(dfs, key=lambda x: x.shape[1])
    
    # --- BUSCA A LINHA DE CABEÇALHO REAL ---
    header_index = -1
    
    # Varre as primeiras 15 linhas procurando "LANÇAMENTO" e "VALOR"
    for i, row in df.head(15).iterrows():
        linha_txt = " ".join([normalizar(x) for x in row.values])
        
        # Palavras-chave baseadas na sua imagem
        if "LANCAMENTO" in linha_txt and "VALOR" in linha_txt:
            header_index = i
            break
            
    if header_index >= 0:
        # Define a linha encontrada como cabeçalho
        df.columns = df.iloc[header_index]
        # Pega os dados apenas DAÍ para baixo
        df = df[header_index+1:].reset_index(drop=True)
    else:
        # Se não achou, assume a primeira linha (melhor que nada)
        df.columns = df.iloc[0]
        df = df[1:].reset_index(drop=True)

    # Exporta para XLSX limpo em memória
    output = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    
    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)
        
    wb.save(output)
    output.seek(0)
    return output

def convert_xls_to_xlsx(file_stream):
    """Converte XLS antigo (binário) para XLSX"""
    if xlrd is None: raise EnvironmentError("Instale xlrd")
    file_stream.seek(0)
    book = xlrd.open_workbook(file_contents=file_stream.read(), formatting_info=False)
    sheet = book.sheet_by_index(0)
    
    wb = Workbook()
    ws = wb.active
    
    for r in range(sheet.nrows):
        row_vals = []
        for c in range(sheet.ncols):
            val = sheet.cell_value(r, c)
            ctype = sheet.cell_type(r, c)
            if ctype == xlrd.XL_CELL_DATE:
                try: val = xldate.xldate_as_datetime(val, book.datemode)
                except: pass
            row_vals.append(val)
        ws.append(row_vals)
        
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out

# --- 3. LÓGICA DE PROCESSAMENTO (MAPEAR E FILTRAR) ---

def processar_dados(file_stream, mes, ano):
    wb_in = load_workbook(file_stream)
    ws_in = wb_in.active
    
    # 1. Mapeamento de Colunas (Onde está o que?)
    # Baseado na sua imagem: LANÇAMENTO, CLASSIFICAÇÃO, RUBRICA, TIPO DOC, COMPETÊNCIA, DATA PAG, VALOR, SITUAÇÃO
    mapa = {}
    
    # Procura o cabeçalho na linha 1
    for cell in ws_in[1]:
        if not cell.value: continue
        nome = normalizar(cell.value)
        mapa[nome] = cell.column # Guarda o índice da coluna (1, 2, 3...)

    # Tenta encontrar índices essenciais. Se não achar pelo nome, tenta posição fixa baseada na imagem
    # A=1, B=2, C=3, D=4, E=5, F=6, G=7, H=8
    idx_lanc = mapa.get("LANCAMENTO") or 1
    idx_rubrica = mapa.get("RUBRICA") or 3
    idx_tipo = mapa.get("TIPO DOCUMENTO") or mapa.get("TIPO DOCUM") or 4
    idx_comp = mapa.get("COMPETENCIA") or mapa.get("COMPETENC") or 5
    idx_data = mapa.get("DATA PAGAMENTO") or mapa.get("DATA PAGAM") or 6
    idx_valor = mapa.get("VALOR") or 7
    idx_sit = mapa.get("SITUACAO") or 8

    # Cria planilha de saída
    wb_out = Workbook()
    ws_out = wb_out.active
    
    # Novo Cabeçalho
    headers = ["CÓDIGO", "FORNECEDOR", "RUBRICA", "DOCUMENTO", "COMPETÊNCIA", "DATA", "SITUAÇÃO", "VALOR"]
    ws_out.append(headers)
    
    # Itera sobre os dados (pulando a linha 1 do cabeçalho)
    for row in ws_in.iter_rows(min_row=2, values_only=True):
        if not any(row): continue
        
        # Filtro de Data (Coluna COMPETÊNCIA - E)
        try:
            raw_comp = row[idx_comp - 1]
            data_comp = converter_data_robusta(raw_comp)
        except IndexError: continue
            
        # Se a data for válida e bater com o filtro
        if data_comp and data_comp.month == mes and data_comp.year == ano:
            
            # --- Extração e Tratamento ---
            
            # 1. Coluna LANÇAMENTO (Nome e Código juntos)
            raw_lanc = row[idx_lanc - 1] if idx_lanc <= len(row) else ""
            fornecedor, codigo_txt = separar_duas_linhas(raw_lanc)
            # Limpa o texto "Cod.: " se existir
            codigo = codigo_txt.replace("Cod.:", "").replace("Cod:", "").strip()
            
            # 2. Coluna TIPO DOCUMENTO (Tipo e Doc juntos)
            raw_tipo = row[idx_tipo - 1] if idx_tipo <= len(row) else ""
            tipo_doc, num_doc = separar_duas_linhas(raw_tipo)
            # Limpa o texto "Doc.: "
            documento = num_doc.replace("Doc.:", "").replace("Doc:", "").strip()
            
            # 3. Outros campos diretos
            rubrica = row[idx_rubrica - 1] if idx_rubrica <= len(row) else ""
            situacao = row[idx_sit - 1] if idx_sit <= len(row) else ""
            
            # 4. Data Pagamento
            raw_data = row[idx_data - 1] if idx_data <= len(row) else ""
            data_pag = converter_data_robusta(raw_data)
            
            # 5. Valor
            raw_valor = row[idx_valor - 1] if idx_valor <= len(row) else 0
            valor_float = limpar_valor_monetario(raw_valor)
            
            # --- Monta a linha final ---
            nova_linha = [
                codigo,         # CÓDIGO
                fornecedor,     # FORNECEDOR
                rubrica,        # RUBRICA
                documento,      # DOCUMENTO
                data_comp,      # COMPETÊNCIA (obj data)
                data_pag,       # DATA (obj data)
                situacao,       # SITUAÇÃO
                valor_float     # VALOR (float)
            ]
            
            ws_out.append(nova_linha)

    # --- Estilização Final ---
    formatar_saida(ws_out)
    
    out = io.BytesIO()
    wb_out.save(out)
    out.seek(0)
    return out

def formatar_saida(ws):
    borda = Border(left=Side(style='thin'), right=Side(style='thin'), 
                   top=Side(style='thin'), bottom=Side(style='thin'))
    
    for row in ws.iter_rows():
        for cell in row:
            cell.border = borda
            
            # Cabeçalho
            if cell.row == 1:
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal='center')
            else:
                # Alinhamentos
                if cell.column in [1, 5, 6]: # Cod, Datas
                    cell.alignment = Alignment(horizontal='center')
                elif cell.column == 8: # Valor
                    cell.alignment = Alignment(horizontal='right')
                else:
                    cell.alignment = Alignment(horizontal='left')
                
                # Formata Datas (Colunas E=5, F=6)
                if cell.column == 5: # Competência (mmm-aa)
                    if isinstance(cell.value, (datetime.date, datetime.datetime)):
                        cell.value = mes_abreviado(cell.value)
                elif cell.column == 6: # Data Pagamento (dd/mm/aaaa)
                    if isinstance(cell.value, (datetime.date, datetime.datetime)):
                        cell.number_format = 'dd/mm/yyyy'
                        
                # Formata Moeda (Coluna H=8)
                if cell.column == 8:
                    cell.number_format = '#,##0.00'

# --- ROTAS FLASK ---

@app.route("/")
def index():
    return render_template("site.html")

@app.route("/processar", methods=["POST"])
def processar():
    if 'file' not in request.files: return "Sem arquivo", 400
    file = request.files['file']
    if not file.filename: return "Arquivo vazio", 400
    
    try:
        mes = int(request.form.get('mes'))
        ano = int(request.form.get('ano'))
        stream = file.stream
        nome = file.filename.lower()
        
        # 1. Conversão
        if is_html_file(stream):
            print("LOG: Convertendo HTML...")
            stream = convert_html_to_xlsx(stream)
        elif nome.endswith('.xls'):
            print("LOG: Convertendo XLS binário...")
            stream = convert_xls_to_xlsx(stream)
            
        # 2. Processamento
        output = processar_dados(stream, mes, ano)
        
        filename_out = f"processado_{mes}_{ano}.xlsx"
        return send_file(
            output, 
            as_attachment=True, 
            download_name=filename_out,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        
    except Exception as e:
        print(f"ERRO: {e}")
        return f"Erro no processamento: {str(e)}", 500

if __name__ == "__main__":
    app.run(debug=True)
