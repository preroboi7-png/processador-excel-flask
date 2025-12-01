import io
import datetime
import unicodedata
from flask import Flask, request, send_file, render_template
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Alignment, Border, Side, Font
from openpyxl.utils.exceptions import InvalidFileException

# Tenta importar a biblioteca xlrd para compatibilidade com arquivos .xls (Excel 97-2003)
# ATENÇÃO: Se for rodar em um ambiente de produção como Render ou Heroku, 
# certifique-se de que 'xlrd' está incluído no seu 'requirements.txt'.
try:
    import xlrd
except ImportError:
    xlrd = None
    # AVISO: A biblioteca 'xlrd' não está instalada. Arquivos .xls (Excel 97-2003) NÃO serão suportados.

# Configura o Flask para procurar templates (HTML) e static (CSS) na pasta atual
app = Flask(__name__, template_folder='.', static_folder='.', static_url_path='')

# --- Funções Auxiliares (Lógica de Limpeza e Preparação) ---

def separar_duas_linhas(texto):
    """Separa um texto que contém quebras de linha em duas partes."""
    if texto is None:
        return "", ""
    partes = str(texto).split("\n")
    if len(partes) >= 2:
        return partes[0].strip(), partes[1].strip()
    return str(texto).strip(), ""

def normalizar(texto):
    """Remove acentos e coloca em caixa alta para padronizar a busca de cabeçalhos."""
    if texto is None:
        return ""
    texto = str(texto).strip()
    texto = unicodedata.normalize("NFD", texto)
    return ''.join(c for c in texto if unicodedata.category(c) != "Mn").upper()

def encontrar_coluna(ws, nome):
    """Encontra o índice da coluna dado o nome do cabeçalho na primeira linha."""
    nome = normalizar(nome)
    for col in range(1, ws.max_column + 1):
        valor = ws.cell(row=1, column=col).value
        if normalizar(str(valor)) == nome:
            return col
    return None

def atualizar_indices(ws):
    """Mapeia os nomes das colunas aos seus respectivos índices atuais."""
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
    """Move o conteúdo completo de uma coluna para outra posição."""
    if origem is None or destino is None or origem == destino:
        return
    max_row = ws.max_row
    max_col = ws.max_column
    if origem > max_col: return
    
    # 1. Armazena os dados da coluna de origem
    dados = [ws.cell(row=r, column=origem).value for r in range(1, max_row + 1)]
    
    # Ajusta o destino caso seja maior que o número atual de colunas
    if destino > max_col + 1: destino = max_col + 1
    
    # 2. Desloca as colunas vizinhas para abrir/fechar espaço
    if destino > origem:
        # Move colunas para a esquerda (sobrepondo a coluna de origem, que já foi salva)
        for c in range(origem + 1, destino + 1):
            for r in range(1, max_row + 1):
                ws.cell(row=r, column=c - 1).value = ws.cell(row=r, column=c).value
    elif destino < origem:
        # Move colunas para a direita (abrindo espaço no destino)
        for c in range(origem - 1, destino - 1, -1):
            for r in range(1, max_row + 1):
                ws.cell(row=r, column=c + 1).value = ws.cell(row=r, column=c).value

    # 3. Insere os dados na coluna de destino
    for r in range(1, max_row + 1):
        ws.cell(row=r, column=destino).value = dados[r - 1]

def apagar(ws, coluna):
    """Apaga uma coluna inteira pelo índice."""
    if coluna:
        ws.delete_cols(coluna, 1)

def mes_abreviado(m):
    """Retorna o nome abreviado do mês."""
    return {
        1: "jan", 2: "fev", 3: "mar", 4: "abr", 5: "mai", 6: "jun",
        7: "jul", 8: "ago", 9: "set", 10: "out", 11: "nov", 12: "dez"
    }.get(m, "")

# --- Função de Conversão XLS (xlrd) para XLSX (openpyxl) ---

def convert_xls_to_openpyxl_stream(file_stream):
    """
    Lê o conteúdo de um arquivo .xls (Excel 97-2003) usando xlrd e o reconstrói 
    em um novo Workbook do openpyxl (.xlsx), retornando como um stream em memória.
    """
    if xlrd is None:
        raise EnvironmentError("A biblioteca 'xlrd' é necessária para processar arquivos .xls. Por favor, adicione 'xlrd' ao seu requirements.txt.")
    
    # xlrd.open_workbook precisa do conteúdo do arquivo em bytes
    file_stream.seek(0)
    data = file_stream.read()
    
    # 1. Abre o arquivo .xls com xlrd
    try:
        book_xls = xlrd.open_workbook(file_contents=data, encoding_override="utf-8")
    except Exception as e:
        raise ValueError(f"Erro ao ler arquivo XLS: {e}")

    sheet_xls = book_xls.sheet_by_index(0) # Pega a primeira planilha

    # 2. Cria um novo Workbook OpenPyXL (target .xlsx format)
    wb_xlsx = Workbook()
    ws_xlsx = wb_xlsx.active

    # 3. Copia os dados do XLS para o XLSX, tratando as datas
    for r in range(sheet_xls.nrows):
        row_values = []
        for c in range(sheet_xls.ncols):
            cell = sheet_xls.cell(r, c)
            value = cell.value
            
            # Converte datas/horas do xlrd para Python datetime
            if cell.ctype == xlrd.XL_CELL_DATE:
                try:
                    dt_tuple = xlrd.xldate_as_tuple(value, book_xls.datemode)
                    # Cria um objeto datetime ou date
                    if dt_tuple[3:] == (0, 0, 0): 
                        value = datetime.date(*dt_tuple[:3])
                    else:
                        value = datetime.datetime(*dt_tuple)
                except:
                    # Em caso de erro, mantém o valor original
                    value = str(cell.value) 

            row_values.append(value)
        
        ws_xlsx.append(row_values)

    # 4. Salva o novo Workbook em um stream de bytes
    output_xlsx_stream = io.BytesIO()
    wb_xlsx.save(output_xlsx_stream)
    output_xlsx_stream.seek(0)
    
    return output_xlsx_stream

# --- Lógica de Processamento e Filtro ---

def processar_excel_logic(file_stream, mes_filtro, ano_filtro):
    # 1. Carregar e Filtrar
    try:
        # file_stream é garantidamente um stream XLSX neste ponto
        wb_entrada = load_workbook(file_stream)
    except InvalidFileException:
        raise ValueError("Arquivo inválido (Não é um formato Excel válido).")
    
    ws_entrada = wb_entrada.active
    wb_filtrado = Workbook()
    ws_filtrado = wb_filtrado.active

    # Formatos de data para conversão robusta (se o openpyxl não tiver lido como data)
    DATE_FORMATS = [
        "%d/%m/%Y %H:%M:%S",
        "%d/%m/%Y",          
        "%Y-%m-%d %H:%M:%S", 
        "%Y-%m-%d",          
        "%m/%d/%Y",          
        "%d-%m-%Y",          
    ]

    for i, row in enumerate(ws_entrada.iter_rows(values_only=True)):
        if i == 0:
            # Adiciona os novos cabeçalhos que serão criados pela separação de linhas
            ws_filtrado.append(list(row) + ["CÓDIGO", "DOCUMENTO"])
            continue

        # COLUNA E: COMPETÊNCIA (Índice 4, pois a contagem inicia em 0)
        data_texto = row[4] 
        
        if not data_texto: continue

        data = None
        
        # Rotina de conversão de data robusta
        try:
            if isinstance(data_texto, (datetime.datetime, datetime.date)):
                data = data_texto
            elif isinstance(data_texto, str):
                texto_limpo = data_texto.strip()
                
                for fmt in DATE_FORMATS:
                    try:
                        # Tenta formatar a string
                        try:
                            data = datetime.datetime.strptime(texto_limpo, fmt)
                        except ValueError:
                            # Se falhar, tenta apenas a parte da data
                            if " " in texto_limpo:
                                data = datetime.datetime.strptime(texto_limpo.split(' ')[0], fmt.split(' ')[0])
                            
                        if data: break 
                    except:
                        continue
        except:
            continue

        # Verifica se a data convertida corresponde ao filtro
        if data and data.month == mes_filtro and data.year == ano_filtro:
            linha = list(row)
            
            # Coluna A (LANÇAMENTO): divide em FORNECEDOR e CÓDIGO
            fornecedor, codigo = separar_duas_linhas(row[0])
            # Coluna D (TIPO DOCUMENTO): divide em TIPO DOC e DOCUMENTO
            tipo_doc, documento = separar_duas_linhas(row[3])
            
            linha[0] = fornecedor
            linha[3] = tipo_doc
            
            # Adiciona as novas colunas no final
            linha.append(codigo)
            linha.append(documento)
            
            ws_filtrado.append(linha)

    # 2. Formatação e Estilo
    ws = ws_filtrado
    col = atualizar_indices(ws)

    # Renomeia LANÇAMENTO para FORNECEDOR
    if col.get("FORNECEDOR"):
        ws.cell(row=1, column=col["FORNECEDOR"]).value = "FORNECEDOR"

    # Apaga CLASSIFICAÇÃO e TIPO DOCUMENTO (já separados)
    for deletar in sorted([col.get("CLASSIFICACAO"), col.get("TIPO")], reverse=True):
        if deletar: apagar(ws, deletar)

    col = atualizar_indices(ws)
    comp_col = col.get("COMP")
    
    # Formata a coluna COMPETÊNCIA (COMP) para "mmm-aa"
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
                    if len(partes) == 3: # Formato YYYY-MM-DD
                        dt = datetime.datetime.strptime(txt, "%Y-%m-%d")
                        m, a = dt.month, dt.year
                    elif len(partes) == 2: # Formato MM-AAAA ou MM-AA
                        p_mes = int(partes[0])
                        p_ano = int(partes[1])
                        if len(str(p_ano)) == 2:
                            p_ano += 2000
                        m, a = p_mes, p_ano
                    else:
                        continue
                ws.cell(row=r, column=comp_col).value = f"{mes_abreviado(m)}-{str(a)[-2:]}"
            except: pass

    # Reordenação das colunas finais
    ordem = [("CODIGO", 1), ("FORNECEDOR", 2), ("RUBRICA", 3), ("DOCUMENTO", 4),
              ("COMP", 5), ("DATA", 6), ("SITUACAO", 7), ("VALOR", 8)]
    
    for nome, nova_pos in ordem:
        col = atualizar_indices(ws)
        atual = col.get(nome)
        if atual and atual != nova_pos:
            mover_coluna(ws, atual, nova_pos)

    # Estilos
    # Habilita o filtro automático
    ws.auto_filter.ref = ws.dimensions
    
    # Define a borda fina
    borda = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
    
    # Aplica borda em todas as células
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is not None: 
                cell.border = borda
                # Define o título em negrito
                if cell.row == 1:
                    cell.font = Font(bold=True)

    col = atualizar_indices(ws)
    col_valor = col.get("VALOR")
    
    # Formatação da coluna VALOR
    if col_valor:
        for r in range(2, ws.max_row + 1):
            valor = ws.cell(row=r, column=col_valor).value
            if not valor: continue
            try:
                # Trata strings como "R$ 1.234,56" ou similar
                if isinstance(valor, str):
                    v_str = str(valor).replace("R$", "").replace(" ", "").replace(".", "")
                    v_num = float(v_str.replace(",", "."))
                else:
                    v_num = float(valor)
                    
                ws.cell(row=r, column=col_valor).value = v_num
                # Formato monetário brasileiro
                ws.cell(row=r, column=col_valor).number_format = '#,##0.00' 
            except: pass

    # Alinhamentos
    alinhamentos = {1:"left", 2:"left", 3:"left", 4:"left", 5:"center", 6:"center", 7:"left", 8:"right"}
    
    # Aplica alinhamento nas colunas de 1 a 8
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=8):
        for cell in row:
            if cell.value is not None and cell.column in alinhamentos:
                cell.alignment = Alignment(horizontal=alinhamentos[cell.column])

    # Salva o arquivo final em memória
    output = io.BytesIO()
    wb_filtrado.save(output)
    output.seek(0)
    return output

# --- Rotas do Flask ---

@app.route("/")
def index():
    # Retorna o template HTML para o usuário interagir
    return render_template("site.html")

@app.route("/processar", methods=["POST"])
def processar():
    if 'file' not in request.files: return "Sem arquivo", 400
    file = request.files['file']
    if file.filename == '': return "Arquivo vazio", 400
    
    # 1. Checa o tipo de arquivo e converte se for .xls
    filename = file.filename.lower()
    file_to_process = file.stream # Stream inicial
    
    if filename.endswith('.xls'): 
        # Detecta e converte XLS para XLSX em memória
        try:
            print(f"Detectado arquivo XLS: {filename}. Convertendo para XLSX...")
            file_to_process = convert_xls_to_openpyxl_stream(file_to_process)
        except (EnvironmentError, ValueError) as e:
            return f"Erro na conversão de XLS: {str(e)}. Certifique-se de que 'xlrd' está instalado.", 500
        except Exception as e:
            return f"Erro inesperado durante a conversão do arquivo XLS: {str(e)}", 500

    # 2. Processa o arquivo (garantidamente XLSX neste ponto)
    try:
        mes = int(request.form.get('mes'))
        ano = int(request.form.get('ano'))
        
        # O processamento da lógica de filtro e formatação
        output = processar_excel_logic(file_to_process, mes, ano)
        
        # Envia o novo arquivo XLSX (Excel 2007-365) para download
        return send_file(output, as_attachment=True, download_name=f"processado_{mes}_{ano}.xlsx", mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    except Exception as e:
        print(f"Erro no processamento: {e}")
        return f"Erro no processamento: {str(e)}", 500

if __name__ == "__main__":
    # Rodar o servidor Flask em modo de depuração
    app.run(debug=True)
