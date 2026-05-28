import fitz
import re
import os
import mysql.connector

# Conf Database
DB_CONFIG = {
    "host": "localhost",
    "user": "root",
    "password": "El@mysql.32",
    "database": "qconcursos"
}

IMAGES_DIR = r"C:\Users\elman\GoogleDrive\QuestoesMED\imagens\QuestoesResidencia"
os.makedirs(IMAGES_DIR, exist_ok=True)

GABARITO_FILE = r"C:\Users\elman\GoogleDrive\QuestoesMED\provas\revalida\revalida-2025.1-gabarito.pdf"
gabarito_map = {}
if os.path.exists(GABARITO_FILE):
    gabarito_doc = fitz.open(GABARITO_FILE)
    tokens = []
    for page in gabarito_doc:
        for line in page.get_text().split('\n'):
            l = line.strip()
            if re.match(r'^\d+$', l) or l in ['A', 'B', 'C', 'D', 'Anulada']:
                tokens.append(l)
    
    for i in range(len(tokens)-1):
        if re.match(r'^\d+$', tokens[i]) and tokens[i+1] in ['A', 'B', 'C', 'D', 'Anulada']:
            q = int(tokens[i])
            ans = tokens[i+1]
            gabarito_map[q] = ans
            
    print(f"[LOG] {len(gabarito_map)} gabaritos carregados com sucesso do PDF.")
else:
    print(f"[AVISO] Arquivo de gabarito não encontrado em: {GABARITO_FILE}")
    print("As questões serão inseridas sem gabarito inicial.")

PDF_FILE = r"C:\Users\elman\GoogleDrive\QuestoesMED\provas\revalida\revalida-2025.1.pdf"
if not os.path.exists(PDF_FILE):
    print(f"[ERRO] Arquivo PDF não encontrado: {PDF_FILE}")
    exit(1)
    
doc = fitz.open(PDF_FILE)

questions = {}
current_q = None
current_part = None
expected_alt = "A"

def is_header_footer(line):
    l = line.strip()
    # Ajuste estas strings ignoradas conforme o cabecalho/rodape do PDF do Revalida se necessário
    if l in ["CADERNO", "01", "Outubro | 25", "AMPLIADA", "LEIA COM ATENÇÃO AS INSTRUÇÕES ABAIXO"] or re.match(r'^\d+$', l):
        return True
    return False

def append_text(q_dict, part, text, is_new_paragraph):
    if part == "enunciado":
        current_text = q_dict["enunciado"]
    else:
        current_text = q_dict["alts"][part]
        
    if current_text:
        if is_new_paragraph:
            if not current_text.endswith('\n'):
                current_text += '\n'
            current_text += text
        else:
            if not current_text.endswith(' ') and not current_text.endswith('\n'):
                current_text += ' '
            current_text += text
    else:
        current_text = text
        
    if part == "enunciado":
        q_dict["enunciado"] = current_text
    else:
        q_dict["alts"][part] = current_text

def extract_table_markdown(page, bbox):
    words = page.get_text('words', clip=bbox)
    if not words: return ''
    
    words.sort(key=lambda w: (w[1], w[0]))
    
    rows = []
    current_row = []
    current_y = words[0][1]
    
    for w in words:
        if abs(w[1] - current_y) > 3.0:
            rows.append(current_row)
            current_row = [w]
            current_y = w[1]
        else:
            current_row.append(w)
    if current_row:
        rows.append(current_row)
        
    x_ranges = []
    for row in rows:
        for w in row:
            x_ranges.append((w[0], w[2]))
            
    x_ranges.sort(key=lambda x: x[0])
    merged_ranges = []
    if x_ranges:
        curr = list(x_ranges[0])
        for r in x_ranges[1:]:
            if r[0] <= curr[1] + 15:
                curr[1] = max(curr[1], r[1])
            else:
                merged_ranges.append(curr)
                curr = list(r)
        merged_ranges.append(curr)
        
    md_rows = []
    for row in rows:
        cols = [''] * len(merged_ranges)
        for w in row:
            wx_center = (w[0] + w[2]) / 2
            for i, r in enumerate(merged_ranges):
                if r[0] - 5 <= wx_center <= r[1] + 5:
                    cols[i] = cols[i] + ' ' + w[4] if cols[i] else w[4]
                    break
        md_rows.append('| ' + ' | '.join(cols) + ' |')
    
    if len(md_rows) > 1:
        md_rows.insert(1, '|' + '|'.join(['---'] * len(merged_ranges)) + '|')
        
    return '\n'.join(md_rows)

for page_idx in range(len(doc)):
    page = doc[page_idx]
    
    tables = page.find_tables()
    page_tables = []
    for t in tables.tables:
        md = extract_table_markdown(page, t.bbox)
        if md and len(md.strip().split('\n')) >= 3:
            page_tables.append({"bbox": t.bbox, "md": md, "processed": False})
        
    blocks = page.get_text("dict")["blocks"]
    
    for block in blocks:
        if block["type"] == 0:  # texto
            lines = []
            prev_y1 = None
            for line_dict in block["lines"]:
                line_text = ""
                for span in line_dict["spans"]:
                    line_text += span["text"]
                
                y0 = line_dict["bbox"][1]
                y1 = line_dict["bbox"][3]
                
                is_new_paragraph = False
                if prev_y1 is not None:
                    if (y0 - prev_y1) > 3.2:
                        is_new_paragraph = True
                else:
                    is_new_paragraph = True
                is_revalida_alt = False
                if len(line_dict["spans"]) >= 2:
                    s0 = line_dict["spans"][0]["text"].strip()
                    if s0 in ["A", "B", "C", "D"]:
                        is_revalida_alt = True
                
                lines.append({
                    "text": line_text,
                    "is_new_paragraph": is_new_paragraph,
                    "bbox": line_dict["bbox"],
                    "is_revalida_alt": is_revalida_alt,
                    "revalida_letra": s0 if is_revalida_alt else None
                })
                prev_y1 = y1
                
            for line_obj in lines:
                line_text = line_obj["text"].strip()
                is_new_paragraph = line_obj["is_new_paragraph"]
                line_bbox = line_obj["bbox"]
                
                is_in_table = False
                cy = (line_bbox[1] + line_bbox[3]) / 2
                cx = (line_bbox[0] + line_bbox[2]) / 2
                for t_info in page_tables:
                    tb = t_info["bbox"]
                    if tb[0] <= cx <= tb[2] and tb[1] <= cy <= tb[3]:
                        is_in_table = True
                        if not t_info["processed"]:
                            if current_q is not None and current_q <= 100:
                                current_part_safe = current_part if current_part else "enunciado"
                                append_text(questions[current_q], current_part_safe, "\n\n" + t_info["md"] + "\n\n", is_new_paragraph=True)
                            t_info["processed"] = True
                        break
                
                if is_in_table:
                    continue
                
                if not line_text:
                    continue
                if is_header_footer(line_text):
                    continue
                
                m_q = re.match(r'^QUESTÃO\s+(\d+)', line_text)
                if m_q:
                    q_num = int(m_q.group(1))
                    
                    # Evitar pegar questionario final (se houver)
                    if current_q == 100 and q_num < 100:
                        current_q = 999 
                        continue
                    if q_num > 100 or current_q == 999:
                        continue
                    
                    if 1 <= q_num <= 100:
                        current_q = q_num
                        questions[current_q] = {
                            "enunciado": "",
                            "alts": {"A": "", "B": "", "C": "", "D": ""},
                            "images": []
                        }
                        current_part = "enunciado"
                        expected_alt = "A"
                        
                        rest = line_text[m_q.end():].strip()
                        if rest:
                            append_text(questions[current_q], "enunciado", rest, is_new_paragraph=False)
                        continue
                
                if current_q is not None and current_q <= 100:
                    m_alt = re.match(r'^\(([A-D])\)\s*(.*)', line_text)
                    if m_alt:
                        letra = m_alt.group(1)
                        current_part = letra
                        append_text(questions[current_q], letra, m_alt.group(2).strip(), is_new_paragraph=False)
                    elif line_text.startswith('(A)') or line_text.startswith('(B)') or line_text.startswith('(C)') or line_text.startswith('(D)'):
                        letra = line_text[1]
                        current_part = letra
                        append_text(questions[current_q], letra, line_text[3:].strip(), is_new_paragraph=False)
                    elif line_obj.get("is_revalida_alt") and line_obj.get("revalida_letra") == expected_alt:
                        letra = line_obj["revalida_letra"]
                        current_part = letra
                        rest = line_text[2:].strip()
                        append_text(questions[current_q], letra, rest, is_new_paragraph=False)
                        next_alt_map = {"A": "B", "B": "C", "C": "D", "D": None}
                        expected_alt = next_alt_map[letra]
                    else:
                        append_text(questions[current_q], current_part, line_text, is_new_paragraph)

        elif block["type"] == 1: # imagem
            if current_q is not None and current_q <= 100:
                img_idx = len(questions[current_q]["images"])
                questions[current_q]["images"].append(block)
                placeholder = f"[IMAGEM_PLACEHOLDER_{img_idx}]"
                if current_part:
                    append_text(questions[current_q], current_part, placeholder, is_new_paragraph=True)

print(f"Total questoes encontradas: {len(questions)}")

# Conectar e INSERIR no bando
conn = mysql.connector.connect(**DB_CONFIG)
cursor = conn.cursor()

# Codigo base distinto do ENAMED (600.000.001) para evitar conflitos no banco
codigo_base = 600000101
sucesso_count = 0

for q_num in range(1, 101):
    q_data = questions.get(q_num)
    if not q_data:
        continue
    
    codigo = codigo_base + (q_num - 1)
    
    # save images
    qtde_imagens = len(q_data["images"])
    tem_imagem = 1 if qtde_imagens > 0 else 0
    
    for i, img_block in enumerate(q_data["images"]):
        img_bytes = img_block["image"]
        if i == 0:
            filename = f"{codigo}.jpeg"
        else:
            filename = f"{codigo}_{i+1}.jpeg"
        img_path = os.path.join(IMAGES_DIR, filename)
        with open(img_path, "wb") as f_img:
            f_img.write(img_bytes)
            
        img_tag = f'<img src="C:\\Users\\elman\\GoogleDrive\\QuestoesMED\\imagens\\QuestoesResidencia\\{filename}">'
        placeholder = f"[IMAGEM_PLACEHOLDER_{i}]"
        
        q_data["enunciado"] = q_data["enunciado"].replace(placeholder, img_tag)
        q_data["alts"]["A"] = q_data["alts"]["A"].replace(placeholder, img_tag)
        q_data["alts"]["B"] = q_data["alts"]["B"].replace(placeholder, img_tag)
        q_data["alts"]["C"] = q_data["alts"]["C"].replace(placeholder, img_tag)
        q_data["alts"]["D"] = q_data["alts"]["D"].replace(placeholder, img_tag)
            
    # get gabarito
    gabarito = gabarito_map.get(q_num)
    gabarito_texto = ""
    if gabarito and gabarito in ["A", "B", "C", "D"]:
        gabarito_texto = q_data["alts"][gabarito].strip()
        
    enunciado = q_data["enunciado"].strip()
    altA = q_data["alts"]["A"].strip()
    altB = q_data["alts"]["B"].strip()
    altC = q_data["alts"]["C"].strip()
    altD = q_data["alts"]["D"].strip()
    
    sql = """
    INSERT INTO questaoresidencia (
        codigo, numero, enunciado, alternativaA, alternativaB, alternativaC, alternativaD,
        instituicao, prova, ano, gabarito, gabarito_texto, tem_imagem, qtde_imagens
    ) VALUES (
        %s, %s, %s, %s, %s, %s, %s,
        'INEP', 'Revalida 2025.1', 2025, %s, %s, %s, %s
    )
    ON DUPLICATE KEY UPDATE
        enunciado = VALUES(enunciado),
        alternativaA = VALUES(alternativaA),
        alternativaB = VALUES(alternativaB),
        alternativaC = VALUES(alternativaC),
        alternativaD = VALUES(alternativaD),
        gabarito_texto = VALUES(gabarito_texto);
    """
    cursor.execute(sql, (
        codigo, q_num, enunciado, altA, altB, altC, altD,
        gabarito, gabarito_texto, tem_imagem, qtde_imagens
    ))
    sucesso_count += 1

conn.commit()
cursor.close()
conn.close()

print(f"Insercao concluida com sucesso! Total: {sucesso_count}")
