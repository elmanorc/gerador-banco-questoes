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

GABARITO_FILE = r"c:\Users\elman\git\gerador-banco-questoes\provas\enamed\gabarito-enamed-2025.txt"
gabarito_map = {}
with open(GABARITO_FILE, 'r', encoding='utf-8') as f:
    for line in f:
        m = re.match(r'^(\d+)\s*-\s*([A-Z]|Anulada)', line.strip())
        if m:
            gabarito_map[int(m.group(1))] = m.group(2)

PDF_FILE = r"c:\Users\elman\git\gerador-banco-questoes\provas\enamed\2025_caderno_ampliado_preliminar.pdf"
doc = fitz.open(PDF_FILE)

questions = {}
current_q = None
current_part = None

def is_header_footer(line):
    l = line.strip()
    # Ignorar cabecalho, rodape, assinatura padrao e numeracao solta de pagina
    if l in ["CADERNO", "01", "Outubro | 25", "AMPLIADA", "LEIA COM ATENÇÃO AS INSTRUÇÕES ABAIXO"] or re.match(r'^\d+$', l):
        return True
    return False

for page_idx in range(len(doc)):
    page = doc[page_idx]
    blocks = page.get_text("dict")["blocks"]
    
    for block in blocks:
        if block["type"] == 0:  # texto
            lines = []
            for line_dict in block["lines"]:
                line_text = ""
                for span in line_dict["spans"]:
                    line_text += span["text"]
                lines.append(line_text)
                
            for line in lines:
                line_text = line.strip()
                if not line_text:
                    continue
                if is_header_footer(line_text):
                    continue
                
                m_q = re.match(r'^QUESTÃO\s+(\d+)', line_text)
                if m_q:
                    q_num = int(m_q.group(1))
                    
                    # Evitar pegar questionario final
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
                        
                        rest = line_text[m_q.end():].strip()
                        if rest:
                            questions[current_q]["enunciado"] += rest + "\n"
                        continue
                
                if current_q is not None and current_q <= 100:
                    m_alt = re.match(r'^\(([A-D])\)\s*(.*)', line_text)
                    if m_alt:
                        letra = m_alt.group(1)
                        current_part = letra
                        questions[current_q]["alts"][letra] += m_alt.group(2) + "\n"
                    elif line_text.startswith('(A)') or line_text.startswith('(B)') or line_text.startswith('(C)') or line_text.startswith('(D)'):
                        letra = line_text[1]
                        current_part = letra
                        questions[current_q]["alts"][letra] += line_text[3:].strip() + "\n"
                    else:
                        if current_part == "enunciado":
                            questions[current_q]["enunciado"] += line_text + "\n"
                        elif current_part in ["A", "B", "C", "D"]:
                            questions[current_q]["alts"][current_part] += line_text + "\n"

        elif block["type"] == 1: # imagem
            if current_q is not None and current_q <= 100:
                questions[current_q]["images"].append(block)

print(f"Total questoes encontradas: {len(questions)}")

# Conectar e INSERIR no bando
conn = mysql.connector.connect(**DB_CONFIG)
cursor = conn.cursor()

codigo_base = 500000001
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
    INSERT IGNORE INTO questaoresidencia (
        codigo, numero, enunciado, alternativaA, alternativaB, alternativaC, alternativaD,
        instituicao, prova, ano, gabarito, gabarito_texto, tem_imagem, qtde_imagens
    ) VALUES (
        %s, %s, %s, %s, %s, %s, %s,
        'INEP', 'ENAMED', 2025, %s, %s, %s, %s
    )
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
