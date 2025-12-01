import io
import datetime
import unicodedata
from flask import Flask, request, send_file, render_template
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Alignment, Border, Side, Font
from openpyxl.utils.exceptions import InvalidFileException
from openpyxl.utils.dataframe import dataframe_to_rows

# Importações para tratar arquivos antigos e HTML
try:
    import xlrd
    from xlrd import xldate
except ImportError:
    xlrd = None

try:
    import pandas as pd
except ImportError:
    pd = None

app = Flask(__name__, template_folder='.', static_folder='.', static_url_path='')

# --- 1. FUNÇÕES DE CONVERSÃO E DETECÇÃO ---

def is_html_file(file_stream):
    """
    Verifica se o arquivo é HTML (comum em exportações 'fake' .xls).
    Lê os primeiros bytes para procurar tags como <table ou <html.
    """
    start_pos = file_stream.tell()
    file_stream.seek(0)
    try:
        # Lê o início do arquivo (decodificando bytes para string)
        header = file_stream.read(1024)
        # Verifica assinatura de HTML ou BOM UTF-8 seguido de tag
        if b'<html' in header or b'<table' in header or b'<div' in header:
            return True
        if b'\xef\xbb\xbf<tabl' in header: # Assinatura específica do seu erro
            return True
    except:
        pass
    finally:
        file_stream.seek(start_pos) # Retorna o ponteiro para o início
    return False

def convert_html_to_xlsx(file_stream):
    """Lê HTML e converte para XLSX usando Pandas"""
    if pd is None:
        raise EnvironmentError("Instale 'pandas' e 'lxml' para ler arquivos HTML/XLS.")
    
    file_stream.seek(0)
    # Tenta ler as tabelas do HTML
    try:
        dfs = pd.read_html(file_stream.read().decode('utf-8', errors='ignore'), decimal=',', thousands='.')
    except ValueError:
        raise ValueError("Não foi possível encontrar uma tabela no arquivo HTML.")

    if not dfs:
        raise ValueError("Arquivo HTML vazio ou sem tabelas.")

    # Pega a maior tabela encontrada (assumindo ser a principal)
    df = max(dfs, key=len)

    # Converte para Workbook OpenPyXL
    wb = Workbook()
    ws = wb.active

    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

def convert_xls_to_openpyxl_stream(file_stream):
    """Converte XLS binário (1997-2003) para XLSX"""
    if xlrd is None:
        raise EnvironmentError("Instale 'xlrd' para processar arquivos .xls antigos.")

    file_stream.seek(0)
    file_content = file_stream.read()
    
    book_xls = xlrd.open_workbook(file_contents=file_content, formatting_info=False)
    sheet_xls = book_xls.sheet_by_index(0)

    wb_xlsx = Workbook()
    ws_xlsx = wb_xlsx.active

    for row in range(sheet_xls.nrows):
        new_row = []
        for col in range(sheet_xls.ncols):
            val = sheet_xls.cell_value(row, col)
            cell_type = sheet_xls.cell_type(row, col)

            if cell_type == xlrd.XL_CELL_DATE:
                try:
                    dt = xldate.xldate_as_datetime(val, book_xls.datemode)
                    val = dt
                except: pass
            elif cell_type == xlrd.XL_CELL_NUMBER:
                if val == int(val): val = int(val)

            new_row.append(val)
        ws_xlsx.append(new_row)

    output = io.BytesIO()
    wb_xlsx.save(output)
    output.seek(0)
    return output

# --- 2. FUNÇÕES AUXILIARES DE LIMPEZA ---

def separar_duas_linhas(texto):
    if texto is None: return "", ""
    partes = str(texto).split("\n")
    if len(partes) >= 2:
        return partes[0].strip(), partes[1].strip()
    return str(texto).strip(), ""

def normalizar(texto):
    if texto is None: return ""
    texto = str(texto).strip()
    texto = unicodedata.normalize("NFD", texto)
    return ''.join(c for c in texto if unicodedata.category(c) != "Mn").upper()

def encontrar_coluna(ws, nome):
    nome = normalizar(nome)
    # Procura na linha 1
    for cell in ws[1]:
        if normalizar(str(cell.value)) == nome:
            return cell.column
    return None

def atualizar_indices(ws):
    return {
        "FORNECEDOR": encontrar_coluna(ws, "LANÇAMENTO") or encontrar_coluna(ws, "FORNECEDOR"),
        "RUBRICA": encontrar_coluna(ws, "RUBRÍCA"),
        "CODIGO": encontrar_coluna(ws, "CÓDIGO"),
        "DOCUMENTO": encontrar_coluna(ws, "DOCUMENTO"),
        "COMP": encontrar_coluna(ws, "COMPETÊNCIA"),
        "DATA": encontrar_coluna(ws, "DATA PAGAMENTO"),
        "SITUACAO": encontrar_coluna(ws, "SITUAÇÃO"),
        "VALOR": encontrar_coluna(ws, "VALOR"),
        "CLASSIFICACAO": encontrar_coluna(ws, "CLASSIFICAÇÃO"),
        "TIPO": encontrar_coluna(ws, "TIPO DOCUMENTO"),
    }

def mover_coluna(ws, origem_idx, destino_idx):
    if not origem_idx or not destino_idx or origem_idx == destino_idx: return
    max_row = ws.max_row
    dados = [ws.cell(row=r, column=origem_idx).value for r in range(1, max_row + 1)]
    ws.delete_cols(origem_idx)
    
    ajuste = 0 if origem_idx > destino_idx else -1
    idx_final = max(1, destino_idx + ajuste)
    
    ws.insert_cols(idx_final)
    for r in range(1, max_row + 1):
        ws.cell(row=r, column=idx_final).value = dados[r-1]

def mes_abreviado(m):
    return {1:"jan", 2:"fev", 3:"mar", 4:"abr", 5:"mai", 6:"jun",
            7:"jul", 8:"ago", 9:"set", 10:"out", 11:"nov", 12:"dez"}.get(m, "")

# --- 3. LÓGICA PRINCIPAL ---

def processar_excel_logic(file_stream, mes_filtro, ano_filtro):
    wb_entrada = load_workbook(file_stream)
    ws_entrada = wb_entrada.active
    
    wb_filtrado = Workbook()
    ws_filtrado = wb_filtrado.active

    # Copia cabeçalhos
    headers = [cell.value for cell in ws_entrada[1]]
    if "CÓDIGO" not in headers: headers.append("CÓDIGO")
    if "DOCUMENTO" not in headers: headers.append("DOCUMENTO")
    ws_filtrado.append(headers)

    idx_comp = encontrar_coluna(ws_entrada, "COMPETÊNCIA") or 5 # Padrão coluna 5
    DATE_FORMATS = ["%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%Y/%m/%d"]

    for row in ws_entrada.iter_rows(min_row=2, values_only=True):
        if not any(row): continue
        
        try: val_data = row[idx_comp - 1]
        except IndexError: continue

        data_obj = None
        if isinstance(val_data, (datetime.datetime, datetime.date)):
            data_obj = val_data
        elif isinstance(val_data, str):
            val_str = val_data.strip().split(' ')[0]
            for fmt in DATE_FORMATS:
                try:
                    data_obj = datetime.datetime.strptime(val_str, fmt)
                    break
                except ValueError: continue
        
        if data_obj and data_obj.month == mes_filtro and data_obj.year == ano_filtro:
            linha_nova = list(row)
            
            # Tratamento de linhas aglutinadas
            # Tenta pegar índices dinâmicos, se não der, usa fixos (0 e 3)
            idx_lanc = encontrar_coluna(ws_entrada, "LANÇAMENTO")
            idx_tipo = encontrar_coluna(ws_entrada, "TIPO DOCUMENTO")
            
            i_lanc = (idx_lanc - 1) if idx_lanc else 0
            i_tipo = (idx_tipo - 1) if idx_tipo else 3

            # Separa Fornecedor/Código
            if i_lanc < len(linha_nova):
                forn, cod = separar_duas_linhas(linha_nova[i_lanc])
                linha_nova[i_lanc] = forn
            else:
                cod = ""

            # Separa Tipo/Doc
            if i_tipo < len(linha_nova):
                tipo, doc = separar_duas_linhas(linha_nova[i_tipo])
                linha_nova[i_tipo] = tipo
            else:
                doc = ""

            linha_nova.append(cod)
            linha_nova.append(doc)
            ws_filtrado.append(linha_nova)

    # --- 4. ESTILIZAÇÃO E FORMATAÇÃO ---
    ws = ws_filtrado
    col = atualizar_indices(ws)
    
    # Renomeia LANÇAMENTO
    c_forn = col.get("FORNECEDOR")
    if c_forn: ws.cell(row=1, column=c_forn).value = "FORNECEDOR"

    # Apaga colunas desnecessárias
    for nome in ["CLASSIFICACAO", "TIPO"]: # Apagar TIPO antigo se houver duplicata
        idx = col.get(nome)
        if idx: mover_coluna(ws, idx, 99) # Move pro fim (ou apaga se preferir)
        # Nota: no código original apagava, aqui mantive simples. 
        # Para apagar descomente abaixo:
        # if idx: ws.delete_cols(idx) 

    # Reordena
    col = atualizar_indices(ws)
    ordem = [("CODIGO", 1), ("FORNECEDOR", 2), ("RUBRICA", 3), ("DOCUMENTO", 4),
             ("COMP", 5), ("DATA", 6), ("SITUACAO", 7), ("VALOR", 8)]
    
    for nome, nova_pos in ordem:
        atual = col.get(nome)
        if atual and atual != nova_pos:
            mover_coluna(ws, atual, nova_pos)
            col = atualizar_indices(ws) # Recalcula após mover

    # Formata VALOR e Borda
    borda = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
    c_valor = col.get("VALOR")
    c_comp = col.get("COMP")

    for row in ws.iter_rows():
        for cell in row:
            if cell.value is not None:
                cell.border = borda
                if cell.row == 1: cell.font = Font(bold=True)
                
                # Formata Valor
                if c_valor and cell.column == c_valor and cell.row > 1:
                    if isinstance(cell.value, str):
                        try:
                            v = float(cell.value.replace("R$", "").replace(".", "").replace(",", ".").strip())
                            cell.value = v
                        except: pass
                    cell.number_format = '#,##0.00'

                # Formata Competência (mmm-aa)
                if c_comp and cell.column == c_comp and cell.row > 1:
                    val = cell.value
                    try:
                        if isinstance(val, (datetime.date, datetime.datetime)):
                            cell.value = f"{mes_abreviado(val.month)}-{str(val.year)[-2:]}"
                    except: pass

    output = io.BytesIO()
    wb_filtrado.save(output)
    output.seek(0)
    return output

# --- ROTAS ---

@app.route("/")
def index():
    return render_template("site.html")

@app.route("/processar", methods=["POST"])
def processar():
    if 'file' not in request.files: return "Sem arquivo", 400
    file = request.files['file']
    if file.filename == '': return "Arquivo vazio", 400
    
    try:
        mes = int(request.form.get('mes'))
        ano = int(request.form.get('ano'))
        
        file_stream = file.stream
        filename = file.filename.lower()

        # --- PIPELINE DE DETECÇÃO E CONVERSÃO ---
        
        # 1. Verifica se é HTML disfarçado de XLS (Erro BOF)
        if is_html_file(file_stream):
            print("Detectado arquivo HTML/Fake-XLS. Convertendo com Pandas...")
            file_stream = convert_html_to_xlsx(file_stream)
        
        # 2. Verifica se é XLS binário antigo
        elif filename.endswith('.xls'):
            print("Detectado arquivo binário XLS antigo. Convertendo com xlrd...")
            file_stream = convert_xls_to_openpyxl_stream(file_stream)

        # 3. Processa
        output = processar_excel_logic(file_stream, mes, ano)
        
        return send_file(output, as_attachment=True, download_name=f"processado_{mes}_{ano}.xlsx", mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    except Exception as e:
        print(f"ERRO CRÍTICO: {e}")
        return f"Erro no processamento: {str(e)}", 500

if __name__ == "__main__":
    app.run(debug=True)
