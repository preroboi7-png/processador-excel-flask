import io
import datetime
import unicodedata
from flask import Flask, request, send_file, render_template
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Alignment, Border, Side
from openpyxl.utils.exceptions import InvalidFileException

# Configura o Flask para procurar templates (HTML) e static (CSS) na pasta atual
# Nota: Para rodar no Render, esta pasta deve ser ajustada para 'SEPARADOR/site_3'
app = Flask(__name__, template_folder='.', static_folder='.', static_url_path='')

# --- Funções Auxiliares (Lógica Original) ---

def separar_duas_linhas(texto):
    if texto is None:
        return "", ""
    partes = str(texto).split("\n")
    if len(partes) >= 2:
        return partes[0], partes[1]
    return str(texto), ""

def normalizar(texto):
    if texto is None:
        return ""
    texto = str(texto).strip()
    texto = unicodedata.normalize("NFD", texto)
    return ''.join(c for c in texto if unicodedata.category(c) != "Mn").upper()

def encontrar_coluna(ws, nome):
    nome = normalizar(nome)
    for col in range(1, ws.max_column + 1):
        valor = ws.cell(row=1, column=col).value
        if normalizar(str(valor)) == nome:
            return col
    return None

def atualizar_indices(ws):
    # O OpenPyXL usa índice 1 para a primeira coluna
    return {
        "FORNECEDOR": encontrar_coluna(ws, "LANÇAMENTO"),
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

def mover_coluna(ws, origem, destino):
    if origem is None or destino is None or origem == destino:
        return
    max_row = ws.max_row
    max_col = ws.max_column
    if origem > max_col: return
    
    # 1. Armazena os dados da coluna de origem
    dados = [ws.cell(row=r, column=origem).value for r in range(1, max_row + 1)]
    
    if destino > max_col + 1: destino = max_col + 1
    
    # 2. Desloca as colunas para abrir espaço
    if destino > origem:
        # Move colunas para a esquerda (apagar a original)
        for c in range(origem + 1, destino + 1):
            for r in range(1, max_row + 1):
                ws.cell(row=r, column=c - 1).value = ws.cell(row=r, column=c).value
    elif destino < origem:
        # Move colunas para a direita (abrir espaço)
        for c in range(origem - 1, destino - 1, -1):
            for r in range(1, max_row + 1):
                ws.cell(row=r, column=c + 1).value = ws.cell(row=r, column=c).value

    # 3. Insere os dados na coluna de destino
    for r in range(1, max_row + 1):
        ws.cell(row=r, column=destino).value = dados[r - 1]

def apagar(ws, coluna):
    if coluna:
        ws.delete_cols(coluna, 1)

def mes_abreviado(m):
    return {
        1: "jan", 2: "fev", 3: "mar", 4: "abr", 5: "mai", 6: "jun",
        7: "jul", 8: "ago", 9: "set", 10: "out", 11: "nov", 12: "dez"
    }.get(m, "")

# --- Lógica de Processamento Unificada ---

def processar_excel_logic(file_stream, mes_filtro, ano_filtro):
    # 1. Carregar e Filtrar (Lógica do app.py original)
    try:
        wb_entrada = load_workbook(file_stream)
    except InvalidFileException:
        raise ValueError("Arquivo inválido.")
    
    ws_entrada = wb_entrada.active
    wb_filtrado = Workbook()
    ws_filtrado = wb_filtrado.active # Este será o nosso 'ws' daqui pra frente

    # Formatos de data a serem testados, do mais completo ao mais simples
    DATE_FORMATS = [
        "%d/%m/%Y %H:%M:%S", # DD/MM/AAAA HH:MM:SS (completo)
        "%d/%m/%Y",          # DD/MM/AAAA (data curta)
        "%Y-%m-%d %H:%M:%S", # YYYY-MM-DD HH:MM:SS (Formato ISO Excel)
        "%Y-%m-%d",          # YYYY-MM-DD (Formato ISO curto)
        "%m/%d/%Y",          # MM/DD/AAAA (Formato US)
        "%d-%m-%Y",          # DD-MM-AAAA
    ]

    for i, row in enumerate(ws_entrada.iter_rows(values_only=True)):
        if i == 0:
            ws_filtrado.append(list(row) + ["CÓDIGO", "DOCUMENTO"])
            continue

        data_texto = row[5] # Coluna F (Índice 5, pois a contagem inicia em 0)
        # Atenção: a coluna "DATA PAGAMENTO" no cabeçalho é a 6ª coluna, índice 5.
        # Seu código original usava row[4] (Coluna E - COMPETÊNCIA). Se você quer
        # filtrar por DATA PAGAMENTO, o índice correto é 5.
        # Mantendo o row[4] conforme seu código, mas com AVISO!
        data_texto = row[4] # Corrigindo para Coluna E: COMPETÊNCIA (Índice 4)
        # Se for para filtrar por DATA PAGAMENTO, use row[5]

        if not data_texto: continue

        data = None
        
        # Rotina de conversão de data robusta
        try:
            if isinstance(data_texto, (datetime.datetime, datetime.date)):
                data = data_texto
            elif isinstance(data_texto, str):
                texto_limpo = data_texto.strip()
                
                # Tenta formatar a string com todos os formatos
                for fmt in DATE_FORMATS:
                    try:
                        # Tenta usar a string completa, se falhar, tenta apenas a primeira parte
                        try:
                            data = datetime.datetime.strptime(texto_limpo, fmt)
                        except ValueError:
                            # Tenta pegar apenas a parte da data antes do primeiro espaço
                            if " " in texto_limpo:
                                data = datetime.datetime.strptime(texto_limpo.split(' ')[0], fmt.split(' ')[0])
                            
                        if data:
                            break # Sai do loop de formatos se a conversão for bem-sucedida
                    except:
                        continue # Tenta o próximo formato na lista
        except:
            # Se a leitura falhar completamente (fora da string/datetime check), ignora a linha
            continue

        # Verifica se a data convertida corresponde ao filtro
        if data and data.month == mes_filtro and data.year == ano_filtro:
            linha = list(row)
            # Separa colunas
            a1, a2 = separar_duas_linhas(row[0])
            d1, d2 = separar_duas_linhas(row[3])
            
            linha[0] = a1
            linha[3] = d1
            linha.append(a2)
            linha.append(d2)
            ws_filtrado.append(linha)

    # 2. Formatação e Estilo (Lógica do app2.py original)
    ws = ws_filtrado
    col = atualizar_indices(ws)

    if col["FORNECEDOR"]:
        ws.cell(row=1, column=col["FORNECEDOR"]).value = "FORNECEDOR"

    for deletar in sorted([col.get("CLASSIFICACAO"), col.get("TIPO")], reverse=True):
        if deletar: apagar(ws, deletar)

    col = atualizar_indices(ws)
    comp_col = col.get("COMP")
    if comp_col:
        for r in range(2, ws.max_row + 1):
            valor = ws.cell(row=r, column=comp_col).value
            if not valor: continue
            try:
                if isinstance(valor, (datetime.datetime, datetime.date)):
                    m, a = valor.month, valor.year
                else:
                    txt = str(valor).replace(" 00:00:00", "")
                    partes = txt.split("-")
                    if len(partes) == 3:
                        dt = datetime.datetime.strptime(txt, "%Y-%m-%d")
                        m, a = dt.month, dt.year
                    elif len(partes) == 2:
                        # Tenta MM-AAAA ou MM-AA
                        p_mes = int(partes[0])
                        p_ano = int(partes[1])
                        # Se o ano tem 2 dígitos, assume o século 21
                        if len(str(p_ano)) == 2:
                            p_ano += 2000
                        m, a = p_mes, p_ano
                    else:
                        continue
                ws.cell(row=r, column=comp_col).value = f"{mes_abreviado(m)}-{str(a)[-2:]}"
            except: pass

    # Reordenação
    ordem = [("CODIGO", 1), ("FORNECEDOR", 2), ("RUBRICA", 3), ("DOCUMENTO", 4),
              ("COMP", 5), ("DATA", 6), ("SITUACAO", 7), ("VALOR", 8)]
    
    for nome, nova_pos in ordem:
        col = atualizar_indices(ws)
        atual = col.get(nome)
        if atual and atual != nova_pos:
            mover_coluna(ws, atual, nova_pos)

    # Estilos
    ws.auto_filter.ref = ws.dimensions
    borda = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
    
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is not None: cell.border = borda

    col = atualizar_indices(ws)
    col_valor = col.get("VALOR")
    if col_valor:
        for r in range(2, ws.max_row + 1):
            valor = ws.cell(row=r, column=col_valor).value
            if not valor: continue
            try:
                # Trata strings como "R$ 1.234,56"
                if isinstance(valor, str):
                    v_str = str(valor).replace("R$", "").replace(" ", "").replace(".", "")
                    v_num = float(v_str.replace(",", "."))
                else:
                    v_num = float(valor)
                    
                ws.cell(row=r, column=col_valor).value = v_num
                ws.cell(row=r, column=col_valor).number_format = '#,##0.00'
            except: pass

    alinhamentos = {1:"left", 2:"left", 3:"left", 4:"left", 5:"right", 6:"right", 7:"left", 8:"right"}
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=8):
        for cell in row:
            if cell.value is not None and cell.column in alinhamentos:
                cell.alignment = Alignment(horizontal=alinhamentos[cell.column])

    output = io.BytesIO()
    wb_filtrado.save(output)
    output.seek(0)
    return output

# --- Rotas ---

@app.route("/")
def index():
    # Este arquivo HTML deve estar na mesma pasta do app.py
    return render_template("site.html")

@app.route("/processar", methods=["POST"])
def processar():
    if 'file' not in request.files: return "Sem arquivo", 400
    file = request.files['file']
    if file.filename == '': return "Arquivo vazio", 400
    
    try:
        mes = int(request.form.get('mes'))
        ano = int(request.form.get('ano'))
        output = processar_excel_logic(file, mes, ano)
        return send_file(output, as_attachment=True, download_name=f"processado_{mes}_{ano}.xlsx", mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    except Exception as e:
        # Se ocorrer um erro durante o processamento, retorna a mensagem de erro.
        print(f"Erro no processamento: {e}")
        return f"Erro no processamento: {str(e)}", 500

if __name__ == "__main__":
    app.run(debug=True)
