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
from docx.image.exceptions import UnrecognizedImageError
import mimetypes

# Configurações do banco
DB_CONFIG = {
    "host": "localhost",
    "user": "root",
    "password": "El@mysql.32",
    "database": "qconcursos"
}

def verificar_e_adicionar_imagem(document, img_path, max_width=None):
    """
    Função auxiliar para verificar e adicionar imagem de forma segura.
    Retorna True se a imagem foi adicionada com sucesso, False caso contrário.
    """
    try:
        # Verificar se o arquivo existe
        if not os.path.exists(img_path):
            print(f"[AVISO] Arquivo de imagem não encontrado: {img_path}")
            return False
        
        # Verificar se é um arquivo válido
        if not os.path.isfile(img_path):
            print(f"[AVISO] Caminho não é um arquivo válido: {img_path}")
            return False
        
        # Verificar tamanho do arquivo
        file_size = os.path.getsize(img_path)
        if file_size == 0:
            print(f"[AVISO] Arquivo de imagem vazio: {img_path}")
            return False
        
        # Verificar formato da imagem
        mime_type, _ = mimetypes.guess_type(img_path)
        if mime_type and not mime_type.startswith('image/'):
            print(f"[AVISO] Arquivo não parece ser uma imagem válida: {img_path} (tipo: {mime_type})")
            return False
        
        # Tentar adicionar a imagem
        if max_width:
            document.add_picture(img_path, width=max_width)
        else:
            document.add_picture(img_path)
        
        print(f"[LOG] Imagem adicionada com sucesso: {img_path}")
        return True
        
    except UnrecognizedImageError as e:
        print(f"[ERRO] Formato de imagem não reconhecido: {img_path}")
        print(f"[ERRO] Detalhes: {str(e)}")
        return False
    except Exception as e:
        print(f"[ERRO] Erro ao adicionar imagem {img_path}: {str(e)}")
        return False

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
          AND q.alternativaE is null 
          AND q.comentario IS NOT NULL
          AND CHAR_LENGTH(q.comentario) >= 500
          AND ano>=2020
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
    
    # Verificar se o tópico tem questões antes de processá-lo
    total_questoes = count_questions_in_subtree(topic_tree, questions_by_topic)
    if total_questoes == 0:
        print(f"[LOG] Pulando tópico {topic_tree['nome']} - sem questões")
        # Processar apenas os filhos que têm questões
        for idx, child in enumerate(topic_tree.get('children', []), 1):
            print(f"[LOG] Verificando filho: {child['nome']} (ID: {child['id']})")
            questao_num = add_topic_sections_recursive(
                document,
                child,
                questions_by_topic,
                level=min(level+1, 9),
                numbering=numbering + [idx] if numbering else [1, idx],
                parent_names=parent_names + [topic_tree['nome']] if parent_names else [topic_tree['nome']],
                questao_num=questao_num,
                breadcrumb_raiz=breadcrumb_raiz
            )
        return questao_num
    
    if numbering is None:
        numbering = [1]
    else:
        numbering = numbering.copy()
    if parent_names is None:
        parent_names = []
    numbering_str = '.'.join(str(n) for n in numbering) + '.'
   
    # Calcular questões diretamente associadas ao tópico pai
    questoes_diretas = questions_by_topic.get(topic_tree['id'], [])
    total_questoes_filhos = total_questoes - len(questoes_diretas)
    
    heading_text = f"{numbering_str} {topic_tree['nome']} ({total_questoes} {'questões' if total_questoes != 1 else 'questão'})"

    # Cria nova seção apenas para tópicos de nível 1 e 2
    if numbering != [1] and level <= 2:
        document.add_section(WD_SECTION.NEW_PAGE)
    # Adiciona breadcrumb no cabeçalho da seção apenas para tópicos de nível 1 e 2
    if level <= 2:
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
    
    # Adiciona questões diretamente associadas ao tópico pai
    for q in questoes_diretas:
        print(f"[LOG] Adicionando questão {q.get('codigo', '?')} diretamente ao tópico {topic_tree['nome']}")
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
        add_imagens_enunciado(document, q['enunciado'], q['codigo'], r"C:\Users\elman\OneDrive\Imagens\QuestoesResidencia")
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
            add_comentario_with_images(document, q['comentario'], q['codigo'], r"C:\Users\elman\OneDrive\Imagens\QuestoesResidencia_comentarios")
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
            max_width = get_max_image_width(document)
            if not verificar_e_adicionar_imagem(document, img_path, max_width):
                document.add_paragraph(f"[Imagem não encontrada ou inválida: {img_filename}]")
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
    # Reduz múltiplas linhas em branco para apenas uma (\n\n), mantendo parágrafos separados
    comentario_md = re.sub(r'\n{3,}', '\n\n', comentario_md)
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
            max_width = get_max_image_width(document)
            if not verificar_e_adicionar_imagem(document, img_path, max_width):
                document.add_paragraph(f"[Imagem não encontrada ou inválida: {img_filename}]")
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
        max_width = get_max_image_width(document)
        if not verificar_e_adicionar_imagem(document, img_path, max_width):
            document.add_paragraph(f"[Imagem não encontrada ou inválida: {img_filename}]")
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
            # Estratégia para reduzir total de questões
            if len(questoes) == 1:
                # Tópicos com apenas 1 questão: 50% de chance de serem incluídos
                if random.random() < 0.5:
                    n = 1
                else:
                    n = 0
            else:
                # Tópicos com mais questões: aplicar amostra normal
                n = max(1, int(len(questoes) * amostra))
            
            if n > 0:
                questions_by_topic[tid] = random.sample(questoes, n)
            else:
                # Remove tópicos sem questões selecionadas
                questions_by_topic[tid] = []
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

AREAS = {
    "CIRURGIA": {"codigos": [33], "proporcao": 0.15},
    "CLÍNICA MÉDICA": {"codigos": [100], "proporcao": 0.25},
    "GINECOLOGIA": {"codigos": [183], "proporcao": 0.06},
    "OBSTETRÍCIA": {"codigos": [218], "proporcao": 0.09},
    "PEDIATRIA": {"codigos": [48], "proporcao": 0.15},
    "Medicina da Família e Comunidade": {"codigos": [180], "proporcao": 0.10},
    "Saúde Mental": {"codigos": [68], "proporcao": 0.10},
    "Saúde Coletiva": {"codigos": [30, 53, 1593, 1735], "proporcao": 0.10},
}

def get_questoes_por_area(conn, codigos_area, n_questoes):
    # Busca todos os tópicos descendentes dos códigos principais
    todos_topicos = []
    for codigo in codigos_area:
        topic_tree = get_topic_tree_recursive(conn, codigo)
        todos_topicos.extend(get_all_topic_ids(topic_tree))
    
    # Buscar todas as questões da área
    cursor = conn.cursor(dictionary=True)
    format_strings = ','.join(['%s'] * len(todos_topicos))
    query = f'''
        SELECT q.questao_id
        FROM classificacao_questao cq
        JOIN questaoresidencia q ON cq.id_questao = q.questao_id
        WHERE cq.id_topico IN ({format_strings})
          AND q.alternativaE IS NULL
          AND q.comentario IS NOT NULL
          AND CHAR_LENGTH(q.comentario) >= 500
          AND q.ano >= 2020
    '''
    cursor.execute(query, tuple(todos_topicos))
    questoes = cursor.fetchall()
    
    # Aplicar amostragem para atingir n_questoes
    if len(questoes) > n_questoes:
        # Calcular amostra necessária
        amostra = n_questoes / len(questoes)
        n_selecionadas = int(len(questoes) * amostra)
        questoes_selecionadas = random.sample(questoes, n_selecionadas)
    else:
        # Se há menos questões que o solicitado, usar todas
        questoes_selecionadas = questoes
    
    # Retornar apenas os IDs das questões selecionadas
    return [q['questao_id'] for q in questoes_selecionadas]

def gerar_banco_proporcional(conn, N):
    """
    Gera um banco de questões com N questões totais, respeitando as proporções das áreas principais.
    """
    print(f"[LOG] Gerando banco de questões com {N} questões totais...")
    
    # Calcular número de questões por área
    questoes_por_area = {}
    for area, info in AREAS.items():
        n_area = int(N * info["proporcao"])
        questoes_por_area[area] = n_area
        print(f"[LOG] {area}: {n_area} questões")
    
    # Coletar todas as questões selecionadas
    todas_questoes_ids = []
    area_info = {}
    
    for area, info in AREAS.items():
        print(f"[LOG] Selecionando questões para {area}...")
        questoes_area = get_questoes_por_area(conn, info["codigos"], questoes_por_area[area])
        todas_questoes_ids.extend(questoes_area)
        area_info[area] = {
            "codigos": info["codigos"],
            "proporcao": info["proporcao"],
            "questoes_selecionadas": len(questoes_area),
            "questoes_ids": questoes_area
        }
        print(f"[LOG] {area}: {len(questoes_area)} questões selecionadas")
    
    # Buscar dados completos das questões selecionadas
    if not todas_questoes_ids:
        print("[ERRO] Nenhuma questão foi selecionada!")
        return
    
    cursor = conn.cursor(dictionary=True)
    format_strings = ','.join(['%s'] * len(todas_questoes_ids))
    query = f"""
        SELECT q.*, cq.id_topico
        FROM questaoresidencia q
        JOIN classificacao_questao cq ON q.questao_id = cq.id_questao
        WHERE q.questao_id IN ({format_strings})
        ORDER BY cq.id_topico, q.questao_id
    """
    cursor.execute(query, tuple(todas_questoes_ids))
    questoes_completas = cursor.fetchall()
    
    print(f"[LOG] Total de questões recuperadas: {len(questoes_completas)}")
    
    # Organizar questões por tópico
    questions_by_topic = {}
    for q in questoes_completas:
        tid = q['id_topico']
        if tid not in questions_by_topic:
            questions_by_topic[tid] = []
        questions_by_topic[tid].append(q)
    
    # Recuperar estrutura hierárquica completa dos tópicos selecionados
    print("[LOG] Recuperando estrutura hierárquica dos tópicos...")
    topicos_selecionados = set(questions_by_topic.keys())
    
    # Para cada área, recuperar a árvore hierárquica completa
    topic_trees = []
    for area, info in AREAS.items():
        for codigo in info["codigos"]:
            topic_tree = get_topic_tree_recursive(conn, codigo)
            topic_trees.append(topic_tree)
    
    # Função para coletar todos os tópicos da árvore que têm questões
    def collect_topic_with_questions(topic_tree, questions_by_topic):
        result = []
        if topic_tree['id'] in questions_by_topic:
            result.append(topic_tree)
        for child in topic_tree.get('children', []):
            result.extend(collect_topic_with_questions(child, questions_by_topic))
        return result
    
    # Coletar tópicos com questões em ordem hierárquica
    topicos_hierarquicos = []
    for topic_tree in topic_trees:
        topicos_hierarquicos.extend(collect_topic_with_questions(topic_tree, questions_by_topic))
    
    print(f"[LOG] Tópicos organizados hierarquicamente: {len(topicos_hierarquicos)}")
    
    # Reorganizar questões para limitar sumário ao terceiro nível
    print("[LOG] Reorganizando questões para limitar sumário ao terceiro nível...")
    
    # Função para encontrar o tópico pai de nível 3
    def find_level3_parent(topic_tree, target_id, current_path=None):
        if current_path is None:
            current_path = []
        
        current_path.append(topic_tree['id'])
        
        if topic_tree['id'] == target_id:
            # Encontrar o tópico de nível 3 no caminho
            if len(current_path) >= 3:
                return current_path[2]  # Índice 2 = nível 3 (0-based)
            elif len(current_path) >= 2:
                return current_path[1]  # Se não tem nível 3, usa nível 2
            else:
                return current_path[0]  # Se não tem nível 2, usa nível 1
        
        for child in topic_tree.get('children', []):
            result = find_level3_parent(child, target_id, current_path.copy())
            if result:
                return result
        
        return None
    
    # Reorganizar questões por tópicos de nível 3 ou menor
    questions_by_level3_topic = {}
    topic_level3_info = {}
    
    for topic in topicos_hierarquicos:
        tid = topic['id']
        questoes_topic = questions_by_topic.get(tid, [])
        
        if not questoes_topic:
            continue
        
        # Encontrar o tópico pai de nível 3
        level3_parent_id = None
        for topic_tree in topic_trees:
            level3_parent_id = find_level3_parent(topic_tree, tid)
            if level3_parent_id:
                break
        
        if not level3_parent_id:
            level3_parent_id = tid  # Se não encontrar, usa o próprio tópico
        
        # Buscar informações do tópico pai de nível 3
        cursor.execute("SELECT nome FROM topico WHERE id = %s", (level3_parent_id,))
        row = cursor.fetchone()
        level3_parent_name = row['nome'] if row else f"Tópico {level3_parent_id}"
        
        # Agrupar questões sob o tópico de nível 3
        if level3_parent_id not in questions_by_level3_topic:
            questions_by_level3_topic[level3_parent_id] = []
            topic_level3_info[level3_parent_id] = {
                'nome': level3_parent_name,
                'id': level3_parent_id
            }
        
        questions_by_level3_topic[level3_parent_id].extend(questoes_topic)
    
    # Determinar nível hierárquico e ordem hierárquica para cada tópico de nível 3
    topic_level3_with_level = []
    for level3_id, questoes in questions_by_level3_topic.items():
        if not questoes:
            continue
        
        # Encontrar o nível hierárquico e a posição hierárquica do tópico
        def get_topic_level_and_position(topic_tree, target_id, current_level=1, current_path=None):
            if current_path is None:
                current_path = []
            
            current_path.append(topic_tree['id'])
            
            if topic_tree['id'] == target_id:
                return current_level, current_path
            
            for child in topic_tree.get('children', []):
                result = get_topic_level_and_position(child, target_id, current_level + 1, current_path.copy())
                if result:
                    return result
            
            return None
        
        level = 1
        hierarchical_path = []
        for topic_tree in topic_trees:
            result = get_topic_level_and_position(topic_tree, level3_id)
            if result:
                level, hierarchical_path = result
                break
        
        topic_level3_with_level.append({
            'id': level3_id,
            'nome': topic_level3_info[level3_id]['nome'],
            'questoes': questoes,
            'level': level,
            'hierarchical_path': hierarchical_path
        })
    
    # Ordenar por caminho hierárquico (mantém a ordem natural da árvore)
    def sort_by_hierarchical_path(topic_info):
        path = topic_info['hierarchical_path']
        # Converte o caminho em uma string ordenável
        return '_'.join(str(id) for id in path)
    
    topic_level3_with_level.sort(key=sort_by_hierarchical_path)
    
    # Gerar numeração hierárquica correta (mesma lógica da opção 1)
    def generate_hierarchical_numbering(topic_level3_with_level):
        """Gera numeração hierárquica sequencial simples"""
        numbering_map = {}
        
        # Contadores para cada nível
        current_main = 0
        current_sub = 0
        current_subsub = 0
        
        for topic_info in topic_level3_with_level:
            level = topic_info['level']
            
            if level == 1:
                # Novo tópico principal
                current_main += 1
                current_sub = 0
                current_subsub = 0
                numbering = [current_main]
            elif level == 2:
                # Subtópico nível 2
                current_sub += 1
                current_subsub = 0
                numbering = [current_main, current_sub]
            elif level == 3:
                # Subtópico nível 3
                current_subsub += 1
                numbering = [current_main, current_sub, current_subsub]
            
            numbering_str = '.'.join(str(n) for n in numbering) + '.'
            numbering_map[topic_info['id']] = numbering_str
        
        return numbering_map
    
    # Gerar numeração hierárquica
    numbering_map = generate_hierarchical_numbering(topic_level3_with_level)
    
    # Adicionar numeração para cada tópico
    for topic_info in topic_level3_with_level:
        topic_info['numbering'] = numbering_map.get(topic_info['id'], "1.")
    
    print(f"[LOG] Tópicos reorganizados para sumário limitado: {len(topic_level3_with_level)}")
    
    # Gerar documento
    document = Document()
    style = document.styles['Normal']
    font = style.font
    font.name = 'Calibri'
    font.size = Pt(12)
    paragraph_format = style.paragraph_format
    paragraph_format.space_after = Pt(3)
    paragraph_format.space_before = Pt(0)
    paragraph_format.line_spacing = 1
    
    # Configurar cabeçalho da capa
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
        try:
            run.add_picture(img_path)
            print(f"[LOG] Imagem de capa adicionada com sucesso: {img_path}")
        except UnrecognizedImageError as e:
            print(f"[ERRO] Formato de imagem de capa não reconhecido: {img_path}")
            print(f"[ERRO] Detalhes: {str(e)}")
        except Exception as e:
            print(f"[ERRO] Erro ao adicionar imagem de capa {img_path}: {str(e)}")
    else:
        print(f"[LOG] Imagem de capa não encontrada: {img_path}")
    p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    
    # Título principal
    capa_title = document.add_paragraph()
    capa_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = capa_title.add_run(f"Banco de Questões - {N} Questões")
    run.bold = True
    run.font.size = Pt(20)
    document.add_paragraph("")
    document.add_paragraph("Sumário:")
    toc_paragraph = document.add_paragraph()
    add_toc(toc_paragraph)
    
    # Adicionar questões organizadas por tópicos de nível 3
    questao_num = 1
    for idx, topic_info in enumerate(topic_level3_with_level, 1):
        tid = topic_info['id']
        nome_topico = topic_info['nome']
        questoes_topic = topic_info['questoes']
        level = topic_info['level']
        hierarchical_path = topic_info['hierarchical_path']
        numbering = topic_info['numbering']

        # Tópico principal em maiúsculas
        if level == 1:
            nome_topico = nome_topico.upper()

        # Calcular total de questões incluindo filhos
        def calculate_total_questions_for_topic(topic_id, topic_level3_with_level):
            total = len(questoes_topic)  # Questões diretamente associadas
            
            # Buscar questões dos filhos
            for child_topic in topic_level3_with_level:
                if child_topic['id'] != topic_id:  # Não contar o próprio tópico
                    # Verificar se este tópico filho pertence ao tópico pai
                    if topic_id in child_topic['hierarchical_path']:
                        total += len(child_topic['questoes'])
            
            return total
        
        total_questoes = calculate_total_questions_for_topic(tid, topic_level3_with_level)
        heading_text = f"{numbering} {nome_topico} ({total_questoes} {'questões' if total_questoes != 1 else 'questão'})"
        document.add_section(WD_SECTION.NEW_PAGE)
        section = document.sections[-1]
        section.header.is_linked_to_previous = False
        section.footer.is_linked_to_previous = True
        header = section.header
        for p in header.paragraphs:
            p.clear()
        
        # Adicionar breadcrumb no cabeçalho para tópicos de nível 1, 2 e 3
        if level <= 3:
            # Gerar breadcrumb baseado na numeração do documento
            breadcrumb_parts = []
            
            # Para cada nível no caminho hierárquico, buscar o tópico correspondente
            for i, path_id in enumerate(hierarchical_path):
                # Buscar nome do tópico no caminho
                cursor.execute("SELECT nome FROM topico WHERE id = %s", (path_id,))
                row = cursor.fetchone()
                if row:
                    nome_caminho = row['nome']
                    
                    # Encontrar a numeração correspondente para este tópico
                    # Procurar nos tópicos já processados para encontrar a numeração correta
                    numero_correspondente = None
                    for topic_processed in topic_level3_with_level[:idx]:
                        if topic_processed['id'] == path_id:
                            numero_correspondente = topic_processed['numbering']
                            break
                    
                    # Se não encontrou, usar a numeração atual para o último nível
                    if numero_correspondente is None and i == len(hierarchical_path) - 1:
                        numero_correspondente = numbering
                    
                    if numero_correspondente:
                        breadcrumb_parts.append(f"{numero_correspondente} {nome_caminho}")
            
            breadcrumb = ' > '.join(breadcrumb_parts)
            p = header.paragraphs[0]
            p.clear()
            run = p.add_run(breadcrumb)
            run.bold = True
            p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        
        document.add_heading(heading_text, level=level)
        document.add_paragraph("")
        
        # Adicionar questões do tópico
        for q in questoes_topic:
            print(f"[LOG] Adicionando questão {q.get('codigo', '?')} ao tópico {nome_topico}")
            
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
            p.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
            run = p.add_run(clean_xml_illegal_chars(cabecalho))
            run.bold = True
            
            # Adiciona o enunciado (texto puro) na mesma linha
            enunciado_texto = extrair_texto_sem_imagens(q['enunciado'])
            p.add_run(clean_xml_illegal_chars(enunciado_texto))
            
            # Adiciona as imagens do enunciado (abaixo do texto)
            add_imagens_enunciado(document, q['enunciado'], q['codigo'], r"C:\Users\elman\OneDrive\Imagens\QuestoesResidencia")
            
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
                add_comentario_with_images(document, q['comentario'], q['codigo'], r"C:\Users\elman\OneDrive\Imagens\QuestoesResidencia_comentarios")
            
            questao_num += 1
    
    # Adicionar rodapé
    add_footer_with_text_and_page_number(document)
    
    # Salvar documento
    from datetime import datetime
    data_atual = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_filename = f"banco_questoes_{N}_{data_atual}.docx"
    
    document.save(output_filename)
    print(f"[LOG] Arquivo {output_filename} gerado com sucesso.")
    print(f"[LOG] Total de questões no banco: {len(questoes_completas)}")
    
    return output_filename

def gerar_banco_hierarquico_proporcional(conn, N):
    """
    Gera um banco de questões com N questões totais, respeitando as proporções das áreas principais
    e usando a estrutura hierárquica completa da Opção 1.
    """
    print(f"[LOG] Gerando banco de questões hierárquico com {N} questões totais...")
    
    # Calcular número de questões por área
    questoes_por_area = {}
    for area, info in AREAS.items():
        n_area = int(N * info["proporcao"])
        questoes_por_area[area] = n_area
        print(f"[LOG] {area}: {n_area} questões")
    
    # Coletar todas as questões selecionadas
    todas_questoes_ids = []
    area_info = {}
    
    for area, info in AREAS.items():
        print(f"[LOG] Selecionando questões para {area}...")
        questoes_area = get_questoes_por_area(conn, info["codigos"], questoes_por_area[area])
        todas_questoes_ids.extend(questoes_area)
        area_info[area] = {
            "codigos": info["codigos"],
            "proporcao": info["proporcao"],
            "questoes_selecionadas": len(questoes_area),
            "questoes_ids": questoes_area
        }
        print(f"[LOG] {area}: {len(questoes_area)} questões selecionadas")
    
    # Buscar dados completos das questões selecionadas
    if not todas_questoes_ids:
        print("[ERRO] Nenhuma questão foi selecionada!")
        return
    
    cursor = conn.cursor(dictionary=True)
    format_strings = ','.join(['%s'] * len(todas_questoes_ids))
    query = f"""
        SELECT q.*, cq.id_topico
        FROM questaoresidencia q
        JOIN classificacao_questao cq ON q.questao_id = cq.id_questao
        WHERE q.questao_id IN ({format_strings})
        ORDER BY cq.id_topico, q.questao_id
    """
    cursor.execute(query, tuple(todas_questoes_ids))
    questoes_completas = cursor.fetchall()
    
    print(f"[LOG] Total de questões recuperadas: {len(questoes_completas)}")
    
    # Organizar questões por tópico
    questions_by_topic = {}
    for q in questoes_completas:
        tid = q['id_topico']
        if tid not in questions_by_topic:
            questions_by_topic[tid] = []
        questions_by_topic[tid].append(q)
    
    # Recuperar estrutura hierárquica completa dos tópicos selecionados
    print("[LOG] Recuperando estrutura hierárquica dos tópicos...")
    
    # Para cada área, recuperar a árvore hierárquica completa
    topic_trees = []
    area_topic_mapping = {}  # Mapeia códigos de área para suas árvores
    
    for area, info in AREAS.items():
        area_trees = []
        for codigo in info["codigos"]:
            topic_tree = get_topic_tree_recursive(conn, codigo)
            area_trees.append(topic_tree)
            topic_trees.append(topic_tree)
        area_topic_mapping[area] = area_trees
    
    # Função para coletar todos os tópicos da árvore que têm questões
    def collect_topic_with_questions(topic_tree, questions_by_topic):
        result = []
        if topic_tree['id'] in questions_by_topic:
            result.append(topic_tree)
        for child in topic_tree.get('children', []):
            result.extend(collect_topic_with_questions(child, questions_by_topic))
        return result
    
    # Coletar tópicos com questões em ordem hierárquica por área
    topicos_por_area = {}
    for area, area_trees in area_topic_mapping.items():
        topicos_area = []
        for topic_tree in area_trees:
            topicos_area.extend(collect_topic_with_questions(topic_tree, questions_by_topic))
        topicos_por_area[area] = topicos_area
    
    print(f"[LOG] Tópicos organizados por área: {len(topicos_por_area)}")
    
    # Gerar documento
    document = Document()
    style = document.styles['Normal']
    font = style.font
    font.name = 'Calibri'
    font.size = Pt(12)
    paragraph_format = style.paragraph_format
    paragraph_format.space_after = Pt(3)
    paragraph_format.space_before = Pt(0)
    paragraph_format.line_spacing = 1
    
    # Configurar cabeçalho da capa
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
        try:
            run.add_picture(img_path)
            print(f"[LOG] Imagem de capa adicionada com sucesso: {img_path}")
        except UnrecognizedImageError as e:
            print(f"[ERRO] Formato de imagem de capa não reconhecido: {img_path}")
            print(f"[ERRO] Detalhes: {str(e)}")
        except Exception as e:
            print(f"[ERRO] Erro ao adicionar imagem de capa {img_path}: {str(e)}")
    else:
        print(f"[LOG] Imagem de capa não encontrada: {img_path}")
    p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    
    # Título principal
    capa_title = document.add_paragraph()
    capa_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = capa_title.add_run(f"Banco de Questões - {N} Questões")
    run.bold = True
    run.font.size = Pt(28)
    
    # Adicionar seção de distribuição por área
    document.add_section(WD_SECTION.NEW_PAGE)
    section = document.sections[1]
    section.header.is_linked_to_previous = False
    header = section.header
    for p in header.paragraphs:
        p.clear()
    
    # Listar distribuição por área
    for area, info in area_info.items():
        p = document.add_paragraph()
        p.add_run(f"{area}: ").bold = True
        p.add_run(f"{info['questoes_selecionadas']} questões ({info['proporcao']*100:.0f}%)")
    
    document.add_paragraph("")
    document.add_paragraph("Sumário:")
    toc_paragraph = document.add_paragraph()
    add_toc(toc_paragraph)
    
    # Adicionar questões organizadas por área e estrutura hierárquica
    questao_num = 1
    area_number = 1
    
    for area, info in AREAS.items():
        print(f"[LOG] Processando área: {area}")
        
        # Verificar se a área tem questões
        if area not in topicos_por_area or not topicos_por_area[area]:
            print(f"[LOG] Área {area} não tem questões, pulando...")
            continue
        
        # Para áreas com múltiplos códigos (como Saúde Coletiva), criar subseções
        if len(info["codigos"]) > 1:
            # Área com múltiplos tópicos pai
            print(f"[LOG] Área {area} tem múltiplos tópicos pai: {info['codigos']}")
            
            # Criar seção principal da área
            document.add_section(WD_SECTION.NEW_PAGE)
            section = document.sections[-1]
            section.header.is_linked_to_previous = False
            section.footer.is_linked_to_previous = True
            header = section.header
            for p in header.paragraphs:
                p.clear()
            
            # Breadcrumb para área principal
            p = header.paragraphs[0]
            p.clear()
            run = p.add_run(f"{area_number}. {area}")
            run.bold = True
            p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            
            # Título da área principal
            total_questoes_area = sum(len(questions_by_topic.get(topic['id'], [])) for topic in topicos_por_area[area])
            heading_text = f"{area_number}. {area} ({total_questoes_area} {'questões' if total_questoes_area != 1 else 'questão'})"
            document.add_heading(heading_text, level=1)
            document.add_paragraph("")
            
            # Processar cada tópico pai da área
            for idx, codigo in enumerate(info["codigos"], 1):
                # Buscar árvore do tópico pai
                topic_tree = None
                for tree in area_topic_mapping[area]:
                    if tree['id'] == codigo:
                        topic_tree = tree
                        break
                
                if not topic_tree:
                    print(f"[LOG] Árvore não encontrada para código {codigo}")
                    continue
                
                # Buscar nome do tópico pai
                cursor.execute("SELECT nome FROM topico WHERE id = %s", (codigo,))
                row = cursor.fetchone()
                nome_topico_pai = row['nome'] if row else f"Tópico {codigo}"
                
                print(f"[LOG] Processando tópico pai: {nome_topico_pai} (código {codigo})")
                
                # Processar estrutura hierárquica completa do tópico pai sem criar título duplicado
                # A função add_topic_sections_recursive já criará o título correto
                questao_num = add_topic_sections_recursive(
                    document,
                    topic_tree,
                    questions_by_topic,
                    level=2,
                    numbering=[area_number, idx],
                    parent_names=[area],
                    questao_num=questao_num,
                    breadcrumb_raiz=f"{area_number}. {area}"
                )
            
            area_number += 1
            
        else:
            # Área com um único tópico pai
            codigo = info["codigos"][0]
            topic_tree = area_topic_mapping[area][0]
            
            print(f"[LOG] Processando área {area} com tópico pai: {topic_tree['nome']} (código {codigo})")
            
            # Processar estrutura hierárquica completa sem criar título duplicado
            # A função add_topic_sections_recursive já criará o título correto
            questao_num = add_topic_sections_recursive(
                document,
                topic_tree,
                questions_by_topic,
                level=1,
                numbering=[area_number],
                parent_names=[],
                questao_num=questao_num,
                breadcrumb_raiz=f"{area_number}. {area}"
            )
            
            area_number += 1
    
    # Adicionar rodapé
    add_footer_with_text_and_page_number(document)
    
    # Salvar documento
    from datetime import datetime
    data_atual = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_filename = f"banco_questoes_hierarquico_{N}_{data_atual}.docx"
    
    document.save(output_filename)
    print(f"[LOG] Arquivo {output_filename} gerado com sucesso.")
    print(f"[LOG] Total de questões no banco: {len(questoes_completas)}")
    
    return output_filename

def main(id_topico_raiz, output_filename, amostra=1.0):
    print(f"[LOG] Iniciando main com id_topico_raiz={id_topico_raiz}, output_filename={output_filename}, amostra={amostra}")
    conn = get_connection()
    print("[LOG] Conexão com o banco estabelecida.")
    # Recupera o nome do tópico pai para criar o diretório
    cursor = conn.cursor(dictionary=True)
    cursor.execute("SELECT nome FROM topico WHERE id = %s", (id_topico_raiz,))
    row_raiz = cursor.fetchone()
    nome_raiz = row_raiz['nome'] if row_raiz else 'topico_raiz'
    # Limpa o nome para ser um nome de pasta seguro
    nome_raiz_dir = ''.join(c for c in nome_raiz if c.isalnum() or c in ' _-').strip().replace(' ', '_')
    # Cria o diretório se não existir
    if not os.path.exists(nome_raiz_dir):
        os.makedirs(nome_raiz_dir)
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
            # Estratégia para reduzir total de questões
            if len(questoes) == 1:
                # Tópicos com apenas 1 questão: 50% de chance de serem incluídos
                if random.random() < 0.5:
                    n = 1
                else:
                    n = 0
            else:
                # Tópicos com mais questões: aplicar amostra normal
                n = max(1, int(len(questoes) * amostra))
            
            if n > 0:
                questions_by_topic[tid] = random.sample(questoes, n)
            else:
                # Remove tópicos sem questões selecionadas
                questions_by_topic[tid] = []
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
    # --- Geração dos arquivos dos filhos dentro do diretório ---
    cursor.execute("SELECT id, nome FROM topico WHERE id_pai = %s", (id_topico_raiz,))
    filhos = cursor.fetchall()
    filhos_info = []
    for idx, filho in enumerate(filhos, 1):
        id_filho = filho['id']
        nome_filho = filho['nome']
        numero = f"1.{idx}"
        filename = gerar_nome_arquivo(numero, nome_filho)
        caminho_completo = os.path.join(nome_raiz_dir, filename)
        print(f"[LOG] Gerando banco de questões para o filho: {nome_filho} ({id_filho}) em {caminho_completo}")
        gerar_banco_questoes_para_topico(conn, id_filho, caminho_completo, amostra=amostra)
        filhos_info.append({'id': id_filho, 'nome': nome_filho, 'filename': caminho_completo})
    # --- Geração do sumário dentro do diretório ---
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
            try:
                run.add_picture(img_path)
                print(f"[LOG] Imagem de sumário adicionada com sucesso: {img_path}")
            except UnrecognizedImageError as e:
                print(f"[ERRO] Formato de imagem de sumário não reconhecido: {img_path}")
                print(f"[ERRO] Detalhes: {str(e)}")
            except Exception as e:
                print(f"[ERRO] Erro ao adicionar imagem de sumário {img_path}: {str(e)}")
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
                    # Estratégia para reduzir total de questões
                    if len(questoes) == 1:
                        # Tópicos com apenas 1 questão: 50% de chance de serem incluídos
                        if random.random() < 0.5:
                            n = 1
                        else:
                            n = 0
                    elif len(questoes) <= 3:
                        # Tópicos com 2-3 questões: aplicar amostra reduzida
                        n = max(0, int(len(questoes) * amostra_local * 0.7))  # 30% menos
                    else:
                        # Tópicos com mais questões: aplicar amostra normal
                        n = max(1, int(len(questoes) * amostra_local))
                    
                    if n > 0:
                        questions_by_topic[tid] = random.sample(questoes, n)
                    else:
                        # Remove tópicos sem questões selecionadas
                        questions_by_topic[tid] = []
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
            filename = os.path.basename(filho['filename'])
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
    sumario_path = os.path.join(nome_raiz_dir, "sumario_bancos_de_questoes.docx")
    gerar_sumario_docx_custom(conn, id_topico_raiz, filhos_info, sumario_path)
    # Salva o arquivo principal também dentro do diretório
    #output_filename = os.path.join(nome_raiz_dir, os.path.basename(output_filename))
    #document.save(output_filename)
    #print(f"[LOG] Arquivo {output_filename} gerado com sucesso.")

if __name__ == "__main__":
    print("=== GERADOR DE BANCOS DE QUESTÕES ===")
    print("1. Gerar banco por tópico específico")
    print("2. Gerar banco proporcional por área")
    print("3. Gerar banco hierárquico proporcional por área")
    
    opcao = input("Escolha a opção (1, 2 ou 3): ").strip()
    
    if opcao == "1":
        # Modo original - por tópico específico
        cod = int(input("Código do tópico: "))
        try:
            amostra = float(input("Porcentagem de questões a incluir (0-100, Enter para tudo):") or 100) / 100.0
        except Exception:
            amostra = 1.0
        main(id_topico_raiz=cod, output_filename="questoes_por_topico.docx", amostra=amostra)
    
    elif opcao == "2":
        # Novo modo - proporcional por área
        try:
            N = int(input("Número total de questões do banco (ex: 1000, 2000, 3000): "))
            if N <= 0:
                print("Erro: N deve ser um número positivo!")
                exit(1)
        except ValueError:
            print("Erro: N deve ser um número inteiro!")
            exit(1)
        
        conn = get_connection()
        print("[LOG] Conexão com o banco estabelecida.")
        gerar_banco_proporcional(conn, N)
        conn.close()
    
    elif opcao == "3":
        # Novo modo - hierárquico proporcional por área
        try:
            N = int(input("Número total de questões do banco (ex: 1000, 2000, 3000): "))
            if N <= 0:
                print("Erro: N deve ser um número positivo!")
                exit(1)
        except ValueError:
            print("Erro: N deve ser um número inteiro!")
            exit(1)
        
        conn = get_connection()
        print("[LOG] Conexão com o banco estabelecida.")
        gerar_banco_hierarquico_proporcional(conn, N)
        conn.close()
    
    else:
        print("Opção inválida!")
