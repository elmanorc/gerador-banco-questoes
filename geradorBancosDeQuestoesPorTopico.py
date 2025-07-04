import sys
import random

def info(type, value, tb):
    import traceback
    import pdb
    traceback.print_exception(type, value, tb)
    print("\nEntrando no modo de depuração interativo (pdb) devido a uma exceção não tratada:")
    pdb.post_mortem(tb)

sys.excepthook = info

import mysql.connector
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.section import WD_SECTION
from markdown2 import markdown
from bs4 import BeautifulSoup
from bs4 import Comment
import os
import re
from docx.shared import Inches
from docx.shared import RGBColor
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

# Configurações do banco
DB_CONFIG = {
    "host": "localhost",
    "user": "root",
    "password": "El@mysql.32",
    "database": "qconcursos"
}

def get_connection():
    print("[LOG] Abrindo conexão com o banco de dados...")
    return mysql.connector.connect(**DB_CONFIG)

def get_subtopics(conn, id_topico):
    """Recupera todos os sub-tópicos recursivamente."""
    cursor = conn.cursor(dictionary=True)
    cursor.execute("SELECT id, nome FROM topico WHERE id_pai = %s", (id_topico,))
    subtopics = cursor.fetchall()
    all_subtopics = []
    for sub in subtopics:
        all_subtopics.append(sub)
        all_subtopics.extend(get_subtopics(conn, sub['id']))
    return all_subtopics

def get_topic_tree_recursive(conn, id_topico):
    print(f"[LOG] Buscando árvore de tópicos recursivamente para id_topico={id_topico}")
    cursor = conn.cursor(dictionary=True)
    cursor.execute("SELECT id, nome FROM topico WHERE id = %s", (id_topico,))
    root = cursor.fetchone()
    cursor.execute("SELECT id, nome FROM topico WHERE id_pai = %s", (id_topico,))
    children = cursor.fetchall()
    root['children'] = [get_topic_tree_recursive(conn, child['id']) for child in children]
    return root

def get_all_topic_ids(topic_tree):
    """Retorna uma lista de todos os ids de tópicos na árvore."""
    ids = [topic_tree['id']]
    for child in topic_tree.get('children', []):
        ids.extend(get_all_topic_ids(child))
    return ids

def get_questions_for_topics(conn, topic_ids):
    print(f"[LOG] Buscando questões para tópicos: {topic_ids}")
    format_strings = ','.join(['%s'] * len(topic_ids))
    query = f"""
        SELECT q.*, cq.id_topico
        FROM questaoresidencia q
        JOIN classificacao_questao cq ON q.questao_id = cq.id_questao
        WHERE cq.id_topico IN ({format_strings})
        ORDER BY cq.id_topico, q.questao_id
    """
    cursor = conn.cursor(dictionary=True)
    cursor.execute(query, tuple(topic_ids))
    return cursor.fetchall()

def html_to_text(html):
    soup = BeautifulSoup(html, "html.parser")
    return soup.get_text(separator="\n")

def markdown_to_text(md):
    html = markdown(md)
    text = html_to_text(html)
    # Remove múltiplas linhas em branco seguidas
    text = re.sub(r'\n{3,}', '\n\n', text)
    return text.strip()

def add_toc(paragraph):
    """Adiciona um campo de TOC (sumário) no docx."""
    run = paragraph.add_run()
    fldChar = OxmlElement('w:fldChar')
    fldChar.set(qn('w:fldCharType'), 'begin')
    instrText = OxmlElement('w:instrText')
    instrText.text = 'TOC \\o "1-3" \\h \\z \\u'
    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'separate')
    fldChar3 = OxmlElement('w:fldChar')
    fldChar3.set(qn('w:fldCharType'), 'end')
    r_element = run._r
    r_element.append(fldChar)
    r_element.append(instrText)
    r_element.append(fldChar2)
    r_element.append(fldChar3)

def count_questions_in_subtree(topic_tree, questions_by_topic):
    """Conta o total de questões neste tópico e em todos os seus sub-tópicos."""
    total = len(questions_by_topic.get(topic_tree['id'], []))
    for child in topic_tree.get('children', []):
        total += count_questions_in_subtree(child, questions_by_topic)
    return total

def get_breadcrumb(topic_tree, numbering, parent_names=None):
    """Gera o breadcrumb do tópico atual, ex: 1. Obesidade > 1.1 Diagnóstico > 1.1.1 Avaliação Clínica"""
    if parent_names is None:
        parent_names = []
    breadcrumb_parts = []
    for i, (num, name) in enumerate(zip(numbering, parent_names + [topic_tree['nome']])):
        sub_numbering = '.'.join(str(n) for n in numbering[:i+1])
        breadcrumb_parts.append(f"{sub_numbering}. {name}")
    return ' > '.join(breadcrumb_parts)

def get_breadcrumb_from_db(conn, id_topico):
    nomes = []
    cursor = conn.cursor(dictionary=True)
    while id_topico is not None:
        cursor.execute("SELECT id, nome, id_pai FROM topico WHERE id = %s", (id_topico,))
        row = cursor.fetchone()
        if not row:
            break
        nomes.append(row['nome'])
        id_topico = row['id_pai']
    nomes.reverse()
    return ' > '.join(nomes)

def add_topic_sections_recursive(document, topic_tree, questions_by_topic, level=1, numbering=None, parent_names=None, questao_num=1, breadcrumb_raiz=None):
    print(f"[LOG] Adicionando seção para tópico: {topic_tree['nome']} (ID: {topic_tree['id']})")
    if numbering is None:
        numbering = [1]
    else:
        numbering = numbering.copy()
    if parent_names is None:
        parent_names = []
    numbering_str = '.'.join(str(n) for n in numbering) + '.'
    total_questoes = count_questions_in_subtree(topic_tree, questions_by_topic)
   
    heading_text = f"{numbering_str} {topic_tree['nome']} ({total_questoes} {'questões' if total_questoes != 1 else 'questão'})"

    # Cria nova seção para este tópico (exceto para o primeiro)
    if numbering != [1]:
        document.add_section(WD_SECTION.NEW_PAGE)
    # Adiciona breadcrumb no cabeçalho da seção
    section = document.sections[-1]
    section.header.is_linked_to_previous = False
    section.footer.is_linked_to_previous = True
    header = section.header
    for p in header.paragraphs:
        p.clear()
    if numbering == [1] and breadcrumb_raiz:
        p = header.paragraphs[0]
        p.clear()
        run = p.add_run(breadcrumb_raiz)
        run.bold = True
        p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    else:
        breadcrumb = get_breadcrumb(topic_tree, numbering, parent_names)
        p = header.paragraphs[0]
        p.clear()
        run = p.add_run(breadcrumb)
        run.bold = True
        p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    document.add_heading(heading_text, level=level)
    document.add_paragraph("")
    # Adiciona questões deste tópico (sem heading)
    for q in questions_by_topic.get(topic_tree['id'], []):
        print(f"[LOG] Adicionando questão {q.get('codigo', '?')} ao tópico {topic_tree['nome']}")
        # Determina o nível de dificuldade textual
        dificuldade_val = q.get('dificuldade', 0)
        try:
            dificuldade_val = int(dificuldade_val)
        except Exception:
            dificuldade_val = 0
        if dificuldade_val in [1, 2]:
            nivel_dificuldade = 'FÁCIL'
        elif dificuldade_val == 3:
            nivel_dificuldade = 'MÉDIO'
        elif dificuldade_val in [4, 5]:
            nivel_dificuldade = 'DIFÍCIL'
        else:
            nivel_dificuldade = ''
        # Monta o cabeçalho no padrão solicitado
        cabecalho = (
            f"{questao_num}. (QR.{q['codigo']}, {q['ano']}, {q.get('instituicao', '')}"
            f"{' - ' + q.get('orgao', '') if q.get('orgao') else ''}. Dificuldade: {nivel_dificuldade}). "
        )
        # Cria o parágrafo e adiciona o cabeçalho em negrito
        p = document.add_paragraph()
        p.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY;
        run = p.add_run(clean_xml_illegal_chars(cabecalho))
        run.bold = True
        # Adiciona o enunciado (texto puro) na mesma linha
        enunciado_texto = extrair_texto_sem_imagens(q['enunciado'])
        p.add_run(clean_xml_illegal_chars(enunciado_texto))
        # Adiciona as imagens do enunciado (abaixo do texto)
        add_imagens_enunciado(document, q['enunciado'], q['codigo'], r"C:\\Users\\elman\\OneDrive\\Imagens\\QuestoesResidencia")
        for alt in ['A', 'B', 'C', 'D', 'E']:
            alt_text = q.get(f'alternativa{alt}')
            if alt_text:
                safe_text = clean_xml_illegal_chars(f"{alt}) {alt_text}")
                document.add_paragraph(safe_text)
        document.add_paragraph("")
        p = document.add_paragraph()
        run = p.add_run("------  COMENTÁRIO  ------")
        run.bold = True
        run.font.color.rgb = RGBColor(0x1E, 0x90, 0xFF)
        p = document.add_paragraph()
        gabarito_texto_limpo = clean_xml_illegal_chars(q['gabarito_texto'])
        run = p.add_run(f"Gabarito: {q['gabarito']} - {gabarito_texto_limpo}")
        run.bold = True

        if q.get('comentario'):
            add_comentario_with_images(document, q['comentario'], q['codigo'], r"C:\\Users\\elman\\OneDrive\\Imagens\\QuestoesResidencia_comentarios")
        document.add_paragraph("")  # Espaço
        questao_num += 1
    # Adiciona filhos recursivamente
    for idx, child in enumerate(topic_tree.get('children', []), 1):
        print(f"[LOG] Descendo para sub-tópico: {child['nome']} (ID: {child['id']})")
        questao_num = add_topic_sections_recursive(
            document,
            child,
            questions_by_topic,
            level=min(level+1, 9),
            numbering=numbering + [idx],
            parent_names=parent_names + [topic_tree['nome']],
            questao_num=questao_num,
            breadcrumb_raiz=breadcrumb_raiz
        )
    return questao_num

# Função para adicionar rodapé customizado
def add_footer_with_text_and_page_number(document):
    section = document.sections[0]
    footer = section.footer
    # Limpa o rodapé existente
    for p in footer.paragraphs:
        p.clear()
    # Primeiro parágrafo: texto centralizado
    p_center = footer.add_paragraph()
    p_center.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    p_center.add_run("Questões MED - 2025")
    # Segundo parágrafo: numeração de página à direita
    p_right = footer.add_paragraph()
    p_right.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
    run_right = p_right.add_run()
    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(qn('w:fldCharType'), 'begin')
    instrText = OxmlElement('w:instrText')
    instrText.text = 'PAGE'
    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'separate')
    fldChar3 = OxmlElement('w:fldChar')
    fldChar3.set(qn('w:fldCharType'), 'end')
    run_right._r.append(fldChar1)
    run_right._r.append(instrText)
    run_right._r.append(fldChar2)
    run_right._r.append(fldChar3)
    p_right.add_run(" de ")
    run_total = p_right.add_run()
    fldChar1t = OxmlElement('w:fldChar')
    fldChar1t.set(qn('w:fldCharType'), 'begin')
    instrTextt = OxmlElement('w:instrText')
    instrTextt.text = 'NUMPAGES'
    fldChar2t = OxmlElement('w:fldChar')
    fldChar2t.set(qn('w:fldCharType'), 'separate')
    fldChar3t = OxmlElement('w:fldChar')
    fldChar3t.set(qn('w:fldCharType'), 'end')
    run_total._r.append(fldChar1t)
    run_total._r.append(instrTextt)
    run_total._r.append(fldChar2t)
    run_total._r.append(fldChar3t)

def add_enunciado_with_images(document, enunciado_html, codigo_questao, imagens_dir):
    soup = BeautifulSoup(enunciado_html, "html.parser")
    img_count = [1]
    buffer = []

    def flush_buffer():
        if buffer:
            text = ''.join(buffer).replace('\xa0', ' ')
            text = re.sub(r'\n{3,}', '\n\n', text)
            document.add_paragraph(clean_xml_illegal_chars(text))
            buffer.clear()

    def process_element(elem):
        if isinstance(elem, Comment):
            return  # Ignora comentários
        if isinstance(elem, str):
            text = elem.replace('\xa0', ' ')
            if text:
                buffer.append(text)
        elif elem.name == "img":
            flush_buffer()
            src = elem.get("src", "")
            ext = os.path.splitext(src)[1].split("?")[0]  # pega extensão, ignora query string
            if not ext:
                ext = ".jpeg"  # fallback
            if img_count[0] == 1:
                img_filename = f"{codigo_questao}.jpeg"
            else:
                img_filename = f"{codigo_questao}_{img_count[0]}.jpeg"
            img_path = os.path.join(imagens_dir, img_filename)
            if os.path.exists(img_path):
                max_width = get_max_image_width(document)
                document.add_picture(img_path, width=max_width)
            else:
                document.add_paragraph(f"[Imagem não encontrada: {img_filename}]")
            img_count[0] += 1
        elif elem.name in ["br"]:
            flush_buffer()
        elif elem.name in ["div", "p"]:
            flush_buffer()
            for child in elem.children:
                process_element(child)
            flush_buffer()
        elif elem.name == "span":
            for child in elem.children:
                process_element(child)
        else:
            for child in elem.children:
                process_element(child)

    for elem in soup.contents:
        process_element(elem)
    flush_buffer()

def add_comentario_with_images(document, comentario_md, codigo_questao, imagens_dir):
    html = markdown(comentario_md)
    soup = BeautifulSoup(html, "html.parser")
    img_count = [1]
    buffer = []

    def flush_buffer():
        if buffer:
            text = ''.join(buffer).replace('\xa0', ' ')
            text = re.sub(r'\n{3,}', '\n\n', text)
            text = re.sub(r'(\n\s*\n)+', '\n\n', text)
            document.add_paragraph(clean_xml_illegal_chars(text.strip()))
            buffer.clear()

    def process_element(elem):
        if isinstance(elem, Comment):
            return
        if isinstance(elem, str):
            text = elem.replace('\xa0', ' ')
            if text:
                buffer.append(text)
        elif elem.name == "img":
            flush_buffer()
            src = elem.get("src", "")
            ext = os.path.splitext(src)[1].split("?")[0]
            if not ext:
                ext = ".jpeg"
            if img_count[0] == 1:
                img_filename = f"{codigo_questao}{ext}"
            else:
                img_filename = f"{codigo_questao}_{img_count[0]}{ext}"
            img_path = os.path.join(imagens_dir, img_filename)
            if os.path.exists(img_path):
                max_width = get_max_image_width(document)
                document.add_picture(img_path, width=max_width)
            else:
                document.add_paragraph(f"[Imagem não encontrada: {img_filename}]")
            img_count[0] += 1
        elif elem.name in ["br"]:
            flush_buffer()
        elif elem.name in ["div", "p"]:
            flush_buffer()
            for child in elem.children:
                process_element(child)
            flush_buffer()
        elif elem.name in ["ul", "ol"]:
            for child in elem.children:
                process_element(child)
        elif elem.name == "li":
            item_text = []
            for child in elem.children:
                if isinstance(child, str):
                    item_text.append(child.replace('\xa0', ' '))
                else:
                    # Recursivamente pega o texto dos filhos
                    sub_buffer = []
                    def collect_text(e):
                        if isinstance(e, str):
                            sub_buffer.append(e.replace('\xa0', ' '))
                        else:
                            for sub_e in e.children:
                                collect_text(sub_e)
                    collect_text(child)
                    item_text.append(''.join(sub_buffer))
            document.add_paragraph(clean_xml_illegal_chars("• " + ''.join(item_text).strip()))
        elif elem.name == "span":
            for child in elem.children:
                process_element(child)
        else:
            for child in elem.children:
                process_element(child)

    for elem in soup.contents:
        process_element(elem)
    flush_buffer()

def get_max_image_width(document):
    section = document.sections[-1]
    page_width = section.page_width
    left_margin = section.left_margin
    right_margin = section.right_margin
    return page_width - left_margin - right_margin

# Função para extrair apenas o texto do enunciado, sem imagens
def extrair_texto_sem_imagens(enunciado_html):
    soup = BeautifulSoup(enunciado_html, "html.parser")
    for img in soup.find_all('img'):
        img.decompose()
    return soup.get_text(separator=" ").replace('\xa0', ' ').strip()

# Função para adicionar apenas as imagens do enunciado
def add_imagens_enunciado(document, enunciado_html, codigo_questao, imagens_dir):
    soup = BeautifulSoup(enunciado_html, "html.parser")
    img_count = 1
    for img in soup.find_all('img'):
        if img_count == 1:
            img_filename = f"{codigo_questao}.jpeg"
        else:
            img_filename = f"{codigo_questao}_{img_count}.jpeg"
        img_path = os.path.join(imagens_dir, img_filename)
        if os.path.exists(img_path):
            max_width = get_max_image_width(document)
            document.add_picture(img_path, width=max_width)
        else:
            document.add_paragraph(f"[Imagem não encontrada: {img_filename}]")
        img_count += 1

def clean_xml_illegal_chars(text):
    # Remove caracteres de controle e inválidos para XML (exceto \t, \n, \r)
    # Inclui \ufffe, \uffff, e outros fora do intervalo permitido
    illegal_unichrs = [
        (0x00, 0x08), (0x0B, 0x0C), (0x0E, 0x1F),
        (0x7F, 0x84), (0x86, 0x9F),
        (0xFDD0, 0xFDDF), (0xFFFE, 0xFFFF)
    ]
    re_illegal = u'|'.join('%s-%s' % (chr(low), chr(high)) for (low, high) in illegal_unichrs)
    re_illegal = '[%s]' % re_illegal
    text = re.sub(re_illegal, '', text)
    # Remove qualquer outro caractere de controle ASCII, exceto \t, \n, \r
    text = re.sub(r'[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]', '', text)
    return text

# Função para adicionar hyperlink em python-docx
def add_hyperlink(paragraph, url, text):
    part = paragraph.part
    r_id = part.relate_to(url, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink', is_external=True)
    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('r:id'), r_id)
    new_run = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    new_run.append(rPr)
    t = OxmlElement('w:t')
    t.text = text
    new_run.append(t)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)
    return paragraph

# Função para gerar o banco de questões de um tópico específico
def gerar_banco_questoes_para_topico(conn, id_topico, output_filename, amostra=1.0):
    print(f"[LOG] Iniciando geração do banco de questões para o tópico {id_topico} em {output_filename}")
    topic_tree = get_topic_tree_recursive(conn, id_topico)
    print(f"[LOG] Árvore de tópicos para {id_topico} recuperada.")
    all_topics = get_all_topic_ids(topic_tree)
    print(f"[LOG] IDs de tópicos: {all_topics}")
    questions = get_questions_for_topics(conn, all_topics)
    print(f"[LOG] Total de questões recuperadas: {len(questions)}")
    questions_by_topic = {tid: [] for tid in all_topics}
    for q in questions:
        questions_by_topic[q['id_topico']].append(q)
    # Amostragem por tópico
    for tid in questions_by_topic:
        questoes = questions_by_topic[tid]
        if 0 < amostra < 1 and len(questoes) > 0:
            n = max(1, int(len(questoes) * amostra))
            questions_by_topic[tid] = random.sample(questoes, n)
    document = Document()
    breadcrumb_raiz = get_breadcrumb_from_db(conn, id_topico)
    print(f"[LOG] Breadcrumb raiz: {breadcrumb_raiz}")
    heading = document.add_heading(f"Banco de Questões de {breadcrumb_raiz}", 0)
    heading.runs[0].font.size = Pt(20)
    document.add_paragraph("")
    document.add_paragraph("Sumário:")
    toc_paragraph = document.add_paragraph()
    add_toc(toc_paragraph)
    print(f"[LOG] Adicionando seções de tópicos recursivamente...")
    add_topic_sections_recursive(document, topic_tree, questions_by_topic, level=1, numbering=[1], questao_num=1, breadcrumb_raiz=breadcrumb_raiz)
    add_footer_with_text_and_page_number(document)
    document.save(output_filename)
    print(f"[LOG] Arquivo {output_filename} gerado com sucesso.")

# Função para gerar nome de arquivo conforme a numeração e nome do tópico
def gerar_nome_arquivo(numero, nome_topico):
    nome_limpo = ''.join(c for c in nome_topico if c.isalnum() or c in ' _-').strip().replace(' ', '_')
    return f"{numero}.{nome_limpo[:30]}.docx"

def main(id_topico_raiz, output_filename, amostra=1.0):
    print(f"[LOG] Iniciando main com id_topico_raiz={id_topico_raiz}, output_filename={output_filename}, amostra={amostra}")
    conn = get_connection()
    print("[LOG] Conexão com o banco estabelecida.")
    topic_tree = get_topic_tree_recursive(conn, id_topico_raiz)
    print("[LOG] Árvore de tópicos recuperada.")
    all_topics = get_all_topic_ids(topic_tree)
    print(f"[LOG] IDs de todos os tópicos coletados: {all_topics}")
    questions = get_questions_for_topics(conn, all_topics)
    print(f"[LOG] Total de questões recuperadas: {len(questions)}")
    questions_by_topic = {tid: [] for tid in all_topics}
    for q in questions:
        questions_by_topic[q['id_topico']].append(q)
    # Amostragem por tópico
    for tid in questions_by_topic:
        questoes = questions_by_topic[tid]
        if 0 < amostra < 1 and len(questoes) > 0:
            n = max(1, int(len(questoes) * amostra))
            questions_by_topic[tid] = random.sample(questoes, n)
    document = Document()
    style = document.styles['Normal']
    font = style.font
    font.name = 'Calibri'
    font.size = Pt(12)
    paragraph_format = style.paragraph_format
    paragraph_format.space_after = Pt(3)
    paragraph_format.space_before = Pt(0)
    paragraph_format.line_spacing = 1
    section_capa = document.sections[0]
    section_capa.header.is_linked_to_previous = False
    header_capa = section_capa.first_page_header
    for p in header_capa.paragraphs:
        p.clear()
    img_path = os.path.join(os.path.dirname(__file__), 'logotipo_frase.png')
    p = header_capa.paragraphs[0]
    p.clear()
    if os.path.exists(img_path):
        print(f"[LOG] Adicionando imagem de capa: {img_path}")
        run = p.add_run()
        run.add_picture(img_path)
    else:
        print(f"[LOG] Imagem de capa não encontrada: {img_path}")
    p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    header_normal = section_capa.header
    for p in header_normal.paragraphs:
        p.clear()
    capa_title = document.add_paragraph()
    capa_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = capa_title.add_run(f"Banco de Questões de {get_breadcrumb_from_db(conn, id_topico_raiz)}")
    run.bold = True
    run.font.size = Pt(28)
    document.add_section(WD_SECTION.NEW_PAGE)
    section = document.sections[1]
    section.header.is_linked_to_previous = False
    header = section.header
    for p in header.paragraphs:
        p.clear()
    p = header.paragraphs[0]
    p.clear()
    run = p.add_run(get_breadcrumb_from_db(conn, id_topico_raiz))
    run.bold = True
    p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    toc_paragraph = document.add_paragraph()
    add_toc(toc_paragraph)
    cursor = conn.cursor(dictionary=True)
    cursor.execute("SELECT nome FROM topico WHERE id = %s", (id_topico_raiz,))
    row_raiz = cursor.fetchone()
    nome_raiz = row_raiz['nome'] if row_raiz else 'topico_raiz'
    filename_raiz = gerar_nome_arquivo("1", nome_raiz)
    print(f"[LOG] Gerando banco de questões para o tópico raiz: {nome_raiz} ({id_topico_raiz}) em {filename_raiz}")
    gerar_banco_questoes_para_topico(conn, id_topico_raiz, filename_raiz, amostra=amostra)
    cursor.execute("SELECT id, nome FROM topico WHERE id_pai = %s", (id_topico_raiz,))
    filhos = cursor.fetchall()
    filhos_info = []
    for idx, filho in enumerate(filhos, 1):
        id_filho = filho['id']
        nome_filho = filho['nome']
        numero = f"1.{idx}"
        filename = gerar_nome_arquivo(numero, nome_filho)
        print(f"[LOG] Gerando banco de questões para o filho: {nome_filho} ({id_filho}) em {filename}")
        gerar_banco_questoes_para_topico(conn, id_filho, filename, amostra=amostra)
        filhos_info.append({'id': id_filho, 'nome': nome_filho, 'filename': filename})
    def gerar_sumario_docx_custom(conn, id_topico_raiz, filhos_info, output_filename):
        from docx import Document
        from docx.shared import Pt, RGBColor
        from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
        import os
        print(f"[LOG] Gerando arquivo de sumário: {output_filename}")
        document = Document()
        section = document.sections[0]
        header = section.header
        for p in header.paragraphs:
            p.clear()
        img_path = os.path.join(os.path.dirname(__file__), 'logotipo_frase.png')
        p = header.paragraphs[0]
        p.clear()
        if os.path.exists(img_path):
            run = p.add_run()
            run.add_picture(img_path)
        p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        document.add_paragraph("")
        # Buscar nome do tópico raiz
        cursor = conn.cursor(dictionary=True)
        cursor.execute("SELECT nome FROM topico WHERE id = %s", (id_topico_raiz,))
        row_raiz = cursor.fetchone()
        nome_raiz = row_raiz['nome'] if row_raiz else 'topico_raiz'
        # Calcular total geral de questões (após amostragem) dos filhos
        total_geral = 0
        totais_filhos = []
        for filho in filhos_info:
            topic_tree = get_topic_tree_recursive(conn, filho['id'])
            all_topics = get_all_topic_ids(topic_tree)
            questions = get_questions_for_topics(conn, all_topics)
            questions_by_topic = {tid: [] for tid in all_topics}
            for q in questions:
                questions_by_topic[q['id_topico']].append(q)
            amostra_local = 1.0
            if 'amostra' in globals():
                amostra_local = globals()['amostra']
            for tid in questions_by_topic:
                questoes = questions_by_topic[tid]
                if 0 < amostra_local < 1 and len(questoes) > 0:
                    n = max(1, int(len(questoes) * amostra_local))
                    questions_by_topic[tid] = random.sample(questoes, n)
            total_questoes = sum(len(questions_by_topic[tid]) for tid in questions_by_topic)
            totais_filhos.append(total_questoes)
            total_geral += total_questoes
        # Adiciona o título principal
        titulo = f"Banco de Questões de {nome_raiz} ({total_geral} {'questões' if total_geral != 1 else 'questão'})"
        titulo_paragraph = document.add_paragraph()
        titulo_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        run = titulo_paragraph.add_run(titulo)
        run.bold = True
        run.font.size = Pt(20)
        document.add_paragraph("")
        # Lista os tópicos filhos
        for idx, filho in enumerate(filhos_info, 1):
            nome = filho['nome']
            filename = filho['filename']
            total_questoes = totais_filhos[idx-1]
            texto_link = f"{nome} ({total_questoes} {'questões' if total_questoes != 1 else 'questão'})"
            p = document.add_paragraph()
            p.add_run(f"{idx}. ")
            add_hyperlink(p, filename, texto_link)
            for run in p.runs:
                if run.text == texto_link:
                    run.font.color.rgb = RGBColor(0x00, 0x00, 0xFF)
        document.save(output_filename)
        print(f"[LOG] Arquivo de sumário {output_filename} gerado com sucesso.")
    gerar_sumario_docx_custom(conn, id_topico_raiz, filhos_info, "sumario_bancos_de_questoes.docx")
    document.save(output_filename)
    print(f"[LOG] Arquivo {output_filename} gerado com sucesso.")

if __name__ == "__main__":
    cod = int(input("Codigo do topico:"))
    try:
        amostra = float(input("Porcentagem de questões a incluir (0-100, Enter para tudo):") or 100) / 100.0
    except Exception:
        amostra = 1.0
    main(id_topico_raiz=cod, output_filename="questoes_por_topico.docx", amostra=amostra)
