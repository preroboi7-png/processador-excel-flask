import io
import datetime
import unicodedata
from flask import Flask, request, send_file, render_template
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Alignment, Border, Side, Font
from openpyxl.utils.dataframe import dataframe_to_rows

# Bibliotecas necessárias para a "Simulação de Ctrl+C / Ctrl+V"
try:
    import pandas as pd
    import lxml
except ImportError:
    pd = None

try:
    import xlrd
    from xlrd import xldate
except ImportError:
    xlrd = None

app = Flask(__name__, template_folder='.', static_folder='.', static_url_path='')

# ==============================================================================
# 1. SIMULAÇÃO DO PROCESSO HUMANO (CTRL+C / CTRL+V)
# ==============================================================================

def simular_ctrl_c_ctrl_v(file_stream, filename):
    """
    Simula a ação de abrir o arquivo (seja HTML ou XLS), selecionar tudo,
    copiar e colar em uma planilha Excel 2007-365 limpa.
    """
    file_stream.seek(0)
    
    # Prepara o Workbook "Limpo" (Onde daremos o Ctrl+V)
    wb_limpo = Workbook()
    ws_limpo = wb_limpo.active
    
    df = None

    # TENTATIVA 1: Tratar como HTML (O "Fake XLS")
    # Isso equivale a abrir o HTML no navegador/Excel e copiar a tabela
    try:
        if pd is None: raise EnvironmentError("Pandas não instalado")
        
        # header=None é o segredo: ele pega TUDO como dados brutos (igual Ctrl+A)
        # não tenta adivinhar cabeçalho
        dfs = pd.read_html(
            file_stream.read().decode('utf-8', errors='ignore'), 
            header=None, 
            decimal=',', 
            thousands='.'
        )
        
        if dfs:
            # Pega a tabela com mais dados (geralmente a principal)
            df = max(dfs, key=lambda x: x.size)
            print("LOG: Tabela HTML copiada com sucesso.")
            
    except Exception as e:
        # TENTATIVA 2: Se falhar (talvez seja um XLS binário real), usa xlrd
        if filename.endswith('.xls') and xlrd:
            try:
                print("LOG: Tentando ler como binário XLS...")
                file_stream.seek(0)
                book = xlrd.open_workbook(file_contents=file_stream.read(), formatting_info=False)
                sheet = book.sheet_by_index(0)
                
                # Converte sheet xlrd para lista de listas (simula copia)
                dados = []
                for r in range(sheet.nrows):
                    linha = []
                    for c in range(sheet.ncols):
                        val = sheet.cell_value(r, c)
                        # Trata data do excel antigo
                        if sheet.cell_type(r, c) == xlrd.XL_CELL_DATE:
                            try: val = xldate.xldate_as_datetime(val, book.datemode)
                            except: pass
                        linha.append(val)
                    dados.append(linha)
                
                # Joga direto no Excel limpo
                for row in dados:
                    ws_limpo.append(row)
                
                return wb_limpo # Retorna o Excel "Colado"
                
            except Exception as erro_xls:
                print(f"Erro na leitura binária: {erro_xls}")

    # Se conseguiu ler como HTML (DataFrame), agora fazemos o "Ctrl+V" no Excel
    if df is not None:
        # dataframe_to_rows joga linha por linha no Excel
        # header=False e index=False garantem que só colamos os dados puros
        for row in dataframe_to_rows(df, index=False, header=False):
            # Limpeza básica de valores NaN (vazios)
            row_limpa = ["" if pd.isna(x) else x for x in row]
            ws_limpo.append(row_limpa)
    else:
        raise ValueError("Não foi possível copiar os dados do arquivo (formato não reconhecido).")

    return wb_limpo

# ==============================================================================
# 2. FUNÇÕES DE LIMPEZA E TEXTO
# ==============================================================================

def normalizar(texto):
    if texto is None: return ""
    return ''.join(c for c in unicodedata.normalize("NFD", str(texto).strip()) 
                   if unicodedata.category(c) != "Mn").upper()

def separar_duas_linhas(texto):
    """Separa texto com quebra de linha (Ex: Nome \n Código)"""
    if not texto: return "", ""
    txt = str(texto).strip()
    if "\n" in txt:
        partes = txt.split("\n")
        return partes[0].strip(), partes[1].strip()
    return txt, ""

def limpar_valor_monetario(valor):
    """Transforma R$ texto em float"""
    if isinstance(valor, (int, float)): return float(valor)
    if isinstance(valor, str):
        v = valor.replace("R$", "").replace(" ", "").replace(".", "").replace(",", ".")
        try: return float(v)
        except: return 0.0
    return 0.0

def converter_data(valor):
    """Tenta converter string para data"""
    if not valor: return None
    if isinstance(valor, (datetime.datetime, datetime.date)): return valor
    
    formatos = ["%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%Y/%m/%d"]
    v_str = str(valor).strip().split(' ')[0]
    
    for fmt in formatos:
        try: return datetime.datetime.strptime(v_str, fmt)
        except: continue
    return None

def mes_abreviado(data):
    meses = {1:"jan", 2:"fev", 3:"mar", 4:"abr", 5:"mai", 6:"jun",
             7:"jul", 8:"ago", 9:"set", 10:"out", 11:"nov", 12:"dez"}
    return f"{meses.get(data.month, '')}-{str(data.year)[-2:]}"

# ==============================================================================
# 3. LÓGICA DE FILTRO (Usando o arquivo "Colado")
# ==============================================================================

def encontrar_cabecalho(ws):
    """Procura em qual linha está o cabeçalho real (LANÇAMENTO/COMPETENCIA)"""
    for r in range(1, 20): # Olha as primeiras 20 linhas
        row_vals = [normalizar(c.value) for c in ws[r]]
        txt_row = " ".join(row_vals)
        # Se achar LANCAMENTO e VALOR na mesma linha, é o cabeçalho
        if "LANCAMENTO" in txt_row and "VALOR" in txt_row:
            return r
    return 1

def processar_arquivo_limpo(wb_entrada, mes, ano):
    ws_entrada = wb_entrada.active
    
    # 1. Achar onde começa a tabela
    linha_header = encontrar_cabecalho(ws_entrada)
    
    # 2. Mapear colunas
    mapa = {}
    for cell in ws_entrada[linha_header]:
        if cell.value:
            mapa[normalizar(cell.value)] = cell.column

    # Índices (tenta pegar pelo nome, se falhar tenta posição relativa fixa)
    idx_lanc = mapa.get("LANCAMENTO") or 1
    idx_rubrica = mapa.get("RUBRICA") or 3
    idx_tipo = mapa.get("TIPO DOCUMENTO") or mapa.get("TIPO DOCUM") or 4
    idx_comp = mapa.get("COMPETENCIA") or mapa.get("COMPETENC") or 5
    idx_data = mapa.get("DATA PAGAMENTO") or mapa.get("DATA PAGAM") or 6
    idx_valor = mapa.get("VALOR") or 7
    idx_sit = mapa.get("SITUACAO") or 8

    # 3. Criar arquivo final
    wb_final = Workbook()
    ws_final = wb_final.active
    ws_final.append(["CÓDIGO", "FORNECEDOR", "RUBRICA", "DOCUMENTO", "COMPETÊNCIA", "DATA", "SITUAÇÃO", "VALOR"])

    # 4. Iterar e Filtrar
    for row in ws_entrada.iter_rows(min_row=linha_header + 1, values_only=True):
        if not any(row): continue # Pula linha vazia
        
        # Pega a data de competência
        try:
            raw_comp = row[idx_comp - 1]
            dt_comp = converter_data(raw_comp)
        except: continue

        # Verifica filtro
        if dt_comp and dt_comp.month == mes and dt_comp.year == ano:
            
            # Tratamento dos dados (Separar nomes, limpar códigos)
            raw_lanc = row[idx_lanc - 1] if idx_lanc <= len(row) else ""
            forn, cod = separar_duas_linhas(raw_lanc)
            cod = cod.replace("Cod.:", "").replace("Cod:", "").strip()

            raw_tipo = row[idx_tipo - 1] if idx_tipo <= len(row) else ""
            tipo, doc = separar_duas_linhas(raw_tipo)
            doc = doc.replace("Doc.:", "").replace("Doc:", "").strip()

            raw_valor = row[idx_valor - 1] if idx_valor <= len(row) else 0
            val_float = limpar_valor_monetario(raw_valor)

            rubrica = row[idx_rubrica - 1] if idx_rubrica <= len(row) else ""
            situacao = row[idx_sit - 1] if idx_sit <= len(row) else ""
            dt_pag = converter_data(row[idx_data - 1]) if idx_data <= len(row) else None

            ws_final.append([cod, forn, rubrica, doc, dt_comp, dt_pag, situacao, val_float])

    # 5. Formatação
    thin = Side(style="thin")
    borda = Border(left=thin, right=thin, top=thin, bottom=thin)
    
    for row in ws_final.iter_rows():
        for cell in row:
            cell.border = borda
            if cell.row == 1:
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal='center')
            else:
                if cell.column in [1, 5, 6]: cell.alignment = Alignment(horizontal='center')
                elif cell.column == 8: # Valor
                    cell.alignment = Alignment(horizontal='right')
                    cell.number_format = '#,##0.00'
                
                # Formata Data Visual
                if cell.column == 5 and isinstance(cell.value, datetime.datetime):
                     cell.value = mes_abreviado(cell.value)
                if cell.column == 6 and isinstance(cell.value, datetime.datetime):
                     cell.number_format = 'dd/mm/yyyy'

    # Salvar em memória
    out = io.BytesIO()
    wb_final.save(out)
    out.seek(0)
    return out

# ==============================================================================
# ROTAS
# ==============================================================================

@app.route("/")
def index():
    return render_template("site.html")

@app.route("/processar", methods=["POST"])
def processar():
    if 'file' not in request.files: return "Sem arquivo", 400
    file = request.files['file']
    if not file.filename: return "Vazio", 400

    try:
        mes = int(request.form.get('mes'))
        ano = int(request.form.get('ano'))
        
        # PASSO 1: Simular Ctrl+C (do arquivo original) -> Ctrl+V (num Excel novo)
        print("LOG: Iniciando 'Copia e Cola' virtual...")
        wb_limpo = simular_ctrl_c_ctrl_v(file.stream, file.filename.lower())
        
        # PASSO 2: Usar o Excel novo para processar
        print("LOG: Processando dados limpos...")
        output = processar_arquivo_limpo(wb_limpo, mes, ano)
        
        return send_file(output, as_attachment=True, download_name=f"processado_{mes}_{ano}.xlsx", mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    except Exception as e:
        print(f"ERRO: {e}")
        return f"Erro: {str(e)}", 500

if __name__ == "__main__":
    app.run(debug=True)
