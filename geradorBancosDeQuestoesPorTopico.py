import sys
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
from datetime import datetime
from PIL import Image

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
    Preserva o tamanho original se for menor que max_width.
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
        
        # Calcular tamanho adequado da imagem
        width_to_use = None
        if max_width:
            try:
                with Image.open(img_path) as img:
                    original_width_px = img.width
                    original_height_px = img.height
                    
                    # Converter largura original de pixels para inches (assumindo 96 DPI)
                    original_width_inches = Inches(original_width_px / 96.0)
                    
                    # Usar o menor valor entre largura original e largura máxima
                    width_to_use = min(original_width_inches, max_width)
                    
                    status = "ORIGINAL" if width_to_use == original_width_inches else "REDUZIDA"
                    print(f"[LOG] Imagem {os.path.basename(img_path)}: {original_width_px}x{original_height_px}px ({status})")
                    print(f"[LOG] Largura: original={original_width_inches:.2f}in, máx={max_width:.2f}in, usada={width_to_use:.2f}in")
                    
            except Exception as e:
                print(f"[AVISO] Não foi possível obter dimensões da imagem {img_path}: {str(e)}")
                width_to_use = max_width  # Fallback para largura máxima
        
        # Tentar adicionar a imagem
        if width_to_use:
            document.add_picture(img_path, width=width_to_use)
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

def get_topic_tree_recursive(conn, id_topico, current_level=1, max_level=4):
    print(f"[LOG] Buscando árvore de tópicos recursivamente para id_topico={id_topico} (nível {current_level})")
    cursor = conn.cursor(dictionary=True)
    cursor.execute("SELECT id, nome FROM topico WHERE id = %s", (id_topico,))
    root = cursor.fetchone()
    
    if not root:
        return None
    
    # Adicionar campo 'children' e 'nivel' se não existir
    root['children'] = []
    root['nivel'] = current_level
    
    if current_level >= max_level:
        print(f"[LOG] Limite de profundidade atingido (nível {current_level}) para tópico {root['nome']}")
        return root
    
    cursor.execute("SELECT id, nome FROM topico WHERE id_pai = %s", (id_topico,))
    children = cursor.fetchall()
    for child in children:
        child_tree = get_topic_tree_recursive(conn, child['id'], current_level + 1, max_level)
        if child_tree:
            root['children'].append(child_tree)
    
    return root

def get_all_topic_ids(topic_tree):
    """Retorna uma lista de todos os ids de tópicos na árvore."""
    ids = [topic_tree['id']]
    for child in topic_tree.get('children', []):
        ids.extend(get_all_topic_ids(child))
    return ids

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

def add_topic_sections_recursive(document, topic_tree, questions_by_topic, level=1, numbering=None, parent_names=None, questao_num=1, breadcrumb_raiz=None, permitir_repeticao=True, questoes_adicionadas=None, total_questoes_banco=1000):
    print(f"[LOG] Adicionando seção para tópico: {topic_tree['nome']} (ID: {topic_tree['id']})")
    
    # Usar o nível da árvore se disponível, senão usar o parâmetro level
    current_level = topic_tree.get('nivel', level)
    
    # Inicializar conjunto de questões adicionadas se não fornecido
    if questoes_adicionadas is None:
        questoes_adicionadas = set()
    
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
                level=min(current_level+1, 9),
                numbering=numbering + [idx] if numbering else [1, idx],
                parent_names=parent_names + [topic_tree['nome']] if parent_names else [topic_tree['nome']],
                questao_num=questao_num,
                breadcrumb_raiz=breadcrumb_raiz,
                permitir_repeticao=permitir_repeticao,
                questoes_adicionadas=questoes_adicionadas,
                total_questoes_banco=total_questoes_banco
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

    # Variável para controlar se é o primeiro tópico de nível 1
    is_first_level1 = (current_level == 1 and numbering == [1])
    
    # Lógica de criação de seções baseada no número total de questões
    # Se <= 500 questões: apenas tópicos de nível 1 têm quebras de página
    # Se > 500 questões: tópicos de níveis 1, 2 e 3 têm quebras de página
    needs_new_section = False
    
    if total_questoes_banco <= 500:
        # Para bancos pequenos (<= 500): apenas nível 1 com quebra de página
        if current_level == 1 and not is_first_level1:
            needs_new_section = True
            print(f"[LOG] Banco pequeno ({total_questoes_banco} questões): quebra apenas nível 1")
    else:
        # Para bancos grandes (>= 500): níveis 1, 2 e 3 com quebra de página
        if current_level == 1 and not is_first_level1:
            # Criar nova seção para tópicos de nível 1 a partir do segundo
            needs_new_section = True
        elif current_level in [2, 3]:
            # Sempre criar nova seção para tópicos de nível 2 e 3
            needs_new_section = True
            print(f"[LOG] Banco grande ({total_questoes_banco} questões): quebra níveis 1-3")
    
    if needs_new_section:
        document.add_section(WD_SECTION.NEW_PAGE)
        print(f"[LOG] Nova seção criada para tópico nível {current_level}: {topic_tree['nome']}")
    
    # Adiciona breadcrumb no cabeçalho baseado no número de questões
    # Se <= 500: apenas nível 1 | Se >= 500: níveis 1, 2 e 3
    max_breadcrumb_level = 1 if total_questoes_banco <= 500 else 3
    if current_level <= max_breadcrumb_level:
        section = document.sections[-1]
        section.header.is_linked_to_previous = False
        section.footer.is_linked_to_previous = True
        header = section.header
        for p in header.paragraphs:
            p.clear()
        
        # Gerar breadcrumb numerado para níveis 1, 2 e 3
        breadcrumb_parts = []
        
        # Construir lista com numerações e nomes dos ancestrais + tópico atual
        all_names = parent_names + [topic_tree['nome']]
        
        for i, name in enumerate(all_names):
            # Criar numeração parcial (ex: "1", "1.2", "1.2.3")
            partial_numbering = '.'.join(str(n) for n in numbering[:i+1])
            breadcrumb_parts.append(f"{partial_numbering}. {name}")
        
        breadcrumb_text = ' > '.join(breadcrumb_parts)
        print(f"[LOG] Breadcrumb criado para nível {current_level}: {breadcrumb_text}")
        
        p = header.paragraphs[0]
        p.clear()
        run = p.add_run(breadcrumb_text)
        run.bold = True
        p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    document.add_heading(heading_text, level=current_level)
    document.add_paragraph("")
    
    # Adiciona questões diretamente associadas ao tópico pai
    for q in questoes_diretas:
        # Verificar se a questão já foi adicionada (se não permitir repetição)
        if not permitir_repeticao and q['questao_id'] in questoes_adicionadas:
            print(f"[LOG] Pulando questão {q.get('codigo', '?')} - já adicionada anteriormente")
            continue
            
        print(f"[LOG] Adicionando questão {q.get('codigo', '?')} diretamente ao tópico {topic_tree['nome']}")
        
        # Adicionar questão ao conjunto de questões já adicionadas
        if not permitir_repeticao:
            questoes_adicionadas.add(q['questao_id'])
        
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
            f"{questao_num}. ({q['questao_id']}, {q['ano']}, {q.get('instituicao', '')}"
            f". Dificuldade: {nivel_dificuldade}). "
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
            level=min(current_level+1, 9),
            numbering=numbering + [idx],
            parent_names=parent_names + [topic_tree['nome']],
            questao_num=questao_num,
            breadcrumb_raiz=breadcrumb_raiz,
            permitir_repeticao=permitir_repeticao,
            questoes_adicionadas=questoes_adicionadas,
            total_questoes_banco=total_questoes_banco
        )
    
    return questao_num

# Função para adicionar rodapé customizado em todas as seções
def add_footer_with_text_and_page_number(document):
    # Aplicar rodapé a todas as seções
    for section in document.sections:
        section.footer.is_linked_to_previous = False
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
    """
    Calcula a largura máxima disponível para imagens na página.
    Esta largura é usada como LIMITE MÁXIMO, não como largura forçada.
    """
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

def configurar_metadados_documento(document, total_questoes):
    """
    Configura os metadados do documento DOCX.
    
    Args:
        document: Objeto Document do python-docx
        total_questoes: Número total de questões no banco
    """
    print("[LOG] Configurando metadados do documento...")
    
    # === PROPRIEDADES PRINCIPAIS (CORE PROPERTIES) ===
    
    # 👤 Autor do documento
    document.core_properties.author = "Questões MED"
    
    # 📝 Título do documento
    document.core_properties.title = f"Guia Gratuito para o ENAMED 2025 - {total_questoes} Questões Recentes e Comentadas"
    
    # 📚 Assunto/Tema
    document.core_properties.subject = "Questões de Residência Médica para o ENAMED 2025"
    
    # 🏷️ Palavras-chave (separadas por vírgula)
    document.core_properties.keywords = "medicina, residência médica, questões, prova, banco questões, enamed"
    
    # 👔 Gerente/Responsável
    document.core_properties.manager = "Professor Elmano Cavalcanti"
    
    # 📂 Categoria
    document.core_properties.category = "Educação Médica"
    
    # 💭 Comentários/Descrição
    data_geracao = datetime.now().strftime('%d/%m/%Y às %H:%M')
    document.core_properties.comments = (
        f"Banco de questões médicas de preparação para o ENAMED "
        f"Contém {total_questoes} questões recentes (2020-2025), comentadas e organizadas hierarquicamente pelos temas cobrados na prova."
    )
    
    # 📊 Último editor
    document.core_properties.last_modified_by = "Sistema Gerador de B Questões MED"
    
    # 🔄 Número da revisão
    document.core_properties.revision = 1
    
    # ⏰ Data de criação
    document.core_properties.created = datetime.now()
    
    # 📅 Data de modificação
    document.core_properties.modified = datetime.now()
    
    print(f"[LOG] Metadados configurados:")
    print(f"  - Autor: {document.core_properties.author}")
    print(f"  - Título: {document.core_properties.title}")
    print(f"  - Categoria: {document.core_properties.category}")
    print(f"  - Palavras-chave: {document.core_properties.keywords}")
    print(f"  - Data criação: {document.core_properties.created.strftime('%d/%m/%Y %H:%M')}")

def gerar_banco_estratificacao_deterministica(conn, total_questoes=1000, permitir_repeticao=True):
    """
    Gera um banco de questões usando consulta SQL específica com N questões
    e organizando hierarquicamente com profundidade máxima de nível 4.
    """
    print(f"[LOG] Gerando banco de questões com consulta SQL específica - {total_questoes} questões...")
    
    # Informar comportamento de seções baseado no número de questões
    if total_questoes <= 500:
        print(f"[LOG] Banco COMPACTO ({total_questoes} questões): quebras de página apenas para tópicos de NÍVEL 1")
    else:
        print(f"[LOG] Banco EXPANDIDO ({total_questoes} questões): quebras de página para tópicos de NÍVEIS 1, 2 e 3")
    
    # Executar a consulta SQL fornecida para obter as questões selecionadas
    cursor = conn.cursor(dictionary=True)
    
    # Calcular cotas por área baseado no total N (ordem: 1.Cirurgia, 2.Clínica Médica, 3.Pediatria, 4.Ginecologia, 5.Obstetrícia, 6.Med.Preventiva)
    cotas = {
        'Cirurgia': round(total_questoes * 0.2),
        'Clínica Médica': round(total_questoes * 0.2),
        'Pediatria': round(total_questoes * 0.2),
        'Ginecologia': round(total_questoes * 0.1),
        'Obstetrícia': round(total_questoes * 0.1),
        'Medicina Preventiva': round(total_questoes * 0.2)
    }
    
    print(f"[LOG] Cotas calculadas para {total_questoes} questões: {cotas}")
    print(f"[LOG] Usando consulta com tópicos raiz específicos e ordenação SHA2 determinística")
    
    query_questoes = f"""
    WITH cotas AS (
        SELECT 33   AS topico_id_raiz, 'Cirurgia'            AS area, ROUND({total_questoes} * 0.20) AS qtd
        UNION ALL
        SELECT 100  AS topico_id_raiz, 'Clínica Médica'      AS area, ROUND({total_questoes} * 0.20)
        UNION ALL
        SELECT 48   AS topico_id_raiz, 'Pediatria'           AS area, ROUND({total_questoes} * 0.20)
        UNION ALL
        SELECT 183  AS topico_id_raiz, 'Ginecologia'         AS area, ROUND({total_questoes} * 0.10)
        UNION ALL
        SELECT 218  AS topico_id_raiz, 'Obstetrícia'         AS area, ROUND({total_questoes} * 0.10)
        UNION ALL
        SELECT 29   AS topico_id_raiz, 'Medicina Preventiva' AS area, ROUND({total_questoes} * 0.20)
    ),
    ordenadas AS (
        SELECT 
            q.*,
            ROW_NUMBER() OVER (
                PARTITION BY c.area
                ORDER BY SHA2(CONCAT(q.questao_id, 'SEMENTE_FIXA'), 256)
            ) AS ordem,
            c.qtd
        FROM questaoresidencia q
        JOIN cotas c ON q.area = c.area
        WHERE q.alternativaE IS NULL
          AND q.comentario IS NOT NULL
          AND CHAR_LENGTH(q.comentario) >= 500
          AND q.ano >= 2020
    )
    SELECT 
        o.*
    FROM ordenadas o
    WHERE o.ordem <= o.qtd
    ORDER BY o.area, o.ordem
    """
    
    print("[LOG] Executando consulta SQL para selecionar questões...")
    cursor.execute(query_questoes)
    questoes_selecionadas = cursor.fetchall()
    
    print(f"[LOG] Total de questões selecionadas: {len(questoes_selecionadas)}")
    
    # Mostrar distribuição por área das questões selecionadas
    distribuicao_selecionadas = {}
    for q in questoes_selecionadas:
        area = q['area']
        distribuicao_selecionadas[area] = distribuicao_selecionadas.get(area, 0) + 1
    
    print("[LOG] Distribuição por área das questões selecionadas:")
    for area, count in distribuicao_selecionadas.items():
        print(f"  - {area}: {count} questões")
    
    # Mapear áreas para tópicos raiz conforme definido na consulta (ordem: 1.Cirurgia, 2.Clínica Médica, 3.Pediatria, 4.Ginecologia, 5.Obstetrícia, 6.Med.Preventiva)
    area_para_topico_raiz = {
        'Cirurgia': 33,
        'Clínica Médica': 100,
        'Pediatria': 48,
        'Ginecologia': 183,
        'Obstetrícia': 218,
        'Medicina Preventiva': 29
    }
    
    print(f"[LOG] Mapeamento área -> tópico raiz: {area_para_topico_raiz}")
    
    # Associar cada questão ao seu tópico raiz baseado na área
    questoes_sem_topico = 0
    for q in questoes_selecionadas:
        area = q['area']
        topico_raiz = area_para_topico_raiz.get(area)
        if topico_raiz:
            q['id_topico'] = topico_raiz
        else:
            print(f"[ERRO] Área '{area}' não mapeada para tópico raiz")
            q['id_topico'] = None
            questoes_sem_topico += 1
    
    if questoes_sem_topico == 0:
        print(f"[LOG] Todas as questões associadas aos tópicos raiz por área")
    else:
        print(f"[ERRO] {questoes_sem_topico} questões não puderam ser associadas a tópicos")
    
    # Obter questões com classificações mais específicas para melhor organização
    questao_ids = [q['questao_id'] for q in questoes_selecionadas]
    format_strings = ','.join(['%s'] * len(questao_ids))
    
    query_topicos_especificos = f"""
    SELECT DISTINCT cq.id_topico, cq.id_questao
    FROM classificacao_questao cq
    WHERE cq.id_questao IN ({format_strings})
    ORDER BY cq.id_questao, cq.id_topico
    """
    
    cursor.execute(query_topicos_especificos, tuple(questao_ids))
    classificacoes_especificas = cursor.fetchall()
    
    print(f"[LOG] Classificações específicas encontradas: {len(classificacoes_especificas)}")
    
    # Criar mapeamento de questão -> tópicos específicos para melhor organização
    questao_topicos_especificos = {}
    for classificacao in classificacoes_especificas:
        questao_id = classificacao['id_questao']
        topico_id = classificacao['id_topico']
        if questao_id not in questao_topicos_especificos:
            questao_topicos_especificos[questao_id] = []
        questao_topicos_especificos[questao_id].append(topico_id)
    
    # Usar tópico mais específico se disponível, senão manter tópico raiz
    for q in questoes_selecionadas:
        topicos_especificos = questao_topicos_especificos.get(q['questao_id'], [])
        if topicos_especificos:
            # Usar o primeiro tópico específico encontrado para melhor organização
            q['id_topico'] = topicos_especificos[0]
        # Se não houver tópico específico, mantém o tópico raiz já definido
    
    # Como usamos INNER JOIN, todas as questões têm tópico associado
    questoes_com_topico = questoes_selecionadas
    print(f"[LOG] Questões com tópico associado: {len(questoes_com_topico)}")
    
    # Verificar se obtivemos exatamente o número esperado
    if len(questoes_com_topico) < total_questoes:
        diferenca = total_questoes - len(questoes_com_topico)
        print(f"[AVISO] Obtidas apenas {len(questoes_com_topico)} questões de {total_questoes} solicitadas.")
        print(f"[AVISO] Diferença: {diferenca} questões. Isso pode indicar que não há questões suficientes")
        print(f"[AVISO] no banco que atendam aos critérios (comentário ≥500 chars, ano ≥2020, etc.)")
    
    # Mostrar distribuição final por área
    distribuicao_final = {}
    for q in questoes_com_topico:
        area = q['area']
        distribuicao_final[area] = distribuicao_final.get(area, 0) + 1
    
    print("[LOG] Distribuição final por área:")
    for area, count in distribuicao_final.items():
        cota_esperada = cotas.get(area, 0)
        status = "✅" if count == cota_esperada else f"❌ (esperado: {cota_esperada})"
        print(f"  - {area}: {count} questões {status}")
    
    if len(questoes_com_topico) == total_questoes:
        print(f"✅ [SUCESSO] Exatamente {total_questoes} questões obtidas!")
    else:
        print(f"⚠️ [AVISO] Obtidas {len(questoes_com_topico)} questões de {total_questoes} solicitadas")
    
    # Obter todos os tópicos únicos das questões
    topicos_utilizados = list(set([q['id_topico'] for q in questoes_com_topico]))
    print(f"[LOG] Tópicos únicos utilizados: {len(topicos_utilizados)}")
    
    # Organizar questões por tópico
    questions_by_topic = {}
    for q in questoes_com_topico:
        tid = q['id_topico']
        if tid not in questions_by_topic:
            questions_by_topic[tid] = []
        questions_by_topic[tid].append(q)
    
    print(f"[LOG] Questões organizadas por {len(questions_by_topic)} tópicos")
    
    # Construir hierarquia completa dos tópicos utilizados
    print("[LOG] Construindo hierarquia completa dos tópicos...")
    
    # Obter hierarquia completa dos tópicos (incluindo ancestrais)
    topicos_completos = set(topicos_utilizados)
    
    # Para cada tópico utilizado, buscar todos os ancestrais
    for topico_id in topicos_utilizados:
        cursor.execute("""
            WITH RECURSIVE topico_ancestrais AS (
                SELECT id, id_pai, nome, 1 as nivel
                FROM topico 
                WHERE id = %s
                
                UNION ALL
                
                SELECT t.id, t.id_pai, t.nome, ta.nivel + 1
                FROM topico t
                INNER JOIN topico_ancestrais ta ON ta.id_pai = t.id
                WHERE ta.nivel < 10
            )
            SELECT id FROM topico_ancestrais
        """, (topico_id,))
        
        ancestrais = cursor.fetchall()
        for ancestral in ancestrais:
            topicos_completos.add(ancestral['id'])
    
    print(f"[LOG] Tópicos completos (incluindo ancestrais): {len(topicos_completos)}")
    
    # Buscar informações completas dos tópicos
    topicos_completos_list = list(topicos_completos)
    format_strings = ','.join(['%s'] * len(topicos_completos_list))
    
    cursor.execute(f"""
        SELECT id, nome, id_pai
        FROM topico 
        WHERE id IN ({format_strings})
        ORDER BY id
    """, tuple(topicos_completos_list))
    
    topicos_info = {t['id']: t for t in cursor.fetchall()}
    
    # Construir árvores hierárquicas
    def build_topic_tree(topico_id, nivel_atual=1, max_nivel=4):
        if topico_id not in topicos_info:
            return None
            
        topico = topicos_info[topico_id]
        tree_node = {
            'id': topico_id,
            'nome': topico['nome'],
            'nivel': nivel_atual,
            'children': []
        }
        
        # Se já atingiu o nível máximo, não adicionar mais filhos na árvore
        # mas as questões dos descendentes serão reagrupadas neste nível
        if nivel_atual >= max_nivel:
            return tree_node
        
        # Encontrar filhos diretos
        filhos = [t_id for t_id, t_info in topicos_info.items() 
                 if t_info['id_pai'] == topico_id and t_id in topicos_completos]
        
        for filho_id in sorted(filhos):
            child_tree = build_topic_tree(filho_id, nivel_atual + 1, max_nivel)
            if child_tree:
                tree_node['children'].append(child_tree)
        
        return tree_node
    
    # Encontrar tópicos raiz (sem pai ou pai não está no conjunto)
    topicos_raiz = []
    for topico_id in topicos_completos:
        topico = topicos_info[topico_id]
        if topico['id_pai'] is None or topico['id_pai'] not in topicos_completos:
            topicos_raiz.append(topico_id)
    
    print(f"[LOG] Tópicos raiz encontrados: {len(topicos_raiz)}")
    
    # Construir árvores para cada tópico raiz
    topic_trees = []
    for raiz_id in sorted(topicos_raiz):
        tree = build_topic_tree(raiz_id)
        if tree:
            topic_trees.append(tree)
    
    print(f"[LOG] Árvores construídas: {len(topic_trees)}")
    
    # Definir ordem específica das áreas médicas conforme solicitado
    ordem_areas = [
        'Cirurgia',
        'Clínica Médica',
        'Pediatria', 
        'Ginecologia',
        'Obstetrícia',
        'Medicina Preventiva'
    ]
    
    # Função para determinar a área de um tópico baseado nas questões
    def get_area_from_topic(tree, questions_by_topic):
        # Buscar questões do tópico e seus filhos para determinar a área
        def collect_questions_from_tree(node):
            all_questions = []
            if node['id'] in questions_by_topic:
                all_questions.extend(questions_by_topic[node['id']])
            for child in node.get('children', []):
                all_questions.extend(collect_questions_from_tree(child))
            return all_questions
        
        questoes = collect_questions_from_tree(tree)
        if questoes:
            # Usar a área da primeira questão como representativa
            return questoes[0].get('area', 'Outros')
        return 'Outros'
    
    # Organizar árvores por área
    arvores_por_area = {}
    for tree in topic_trees:
        area = get_area_from_topic(tree, questions_by_topic)
        if area not in arvores_por_area:
            arvores_por_area[area] = []
        arvores_por_area[area].append(tree)
    
    print(f"[LOG] Árvores organizadas por área: {list(arvores_por_area.keys())}")
    
    # Ordenar árvores conforme a sequência desejada: 1.Cirurgia, 2.Clínica Médica, 3.Pediatria, 4.Ginecologia, 5.Obstetrícia, 6.Med.Preventiva
    topic_trees_ordenadas = []
    for i, area in enumerate(ordem_areas, 1):
        if area in arvores_por_area:
            # Ordenar árvores da mesma área por nome do tópico
            arvores_area = sorted(arvores_por_area[area], key=lambda x: x['nome'])
            topic_trees_ordenadas.extend(arvores_area)
            print(f"[LOG] {i}. Adicionada área '{area}' com {len(arvores_area)} árvore(s)")
    
    # Adicionar áreas não mapeadas no final
    for area, arvores in arvores_por_area.items():
        if area not in ordem_areas:
            arvores_area = sorted(arvores, key=lambda x: x['nome'])
            topic_trees_ordenadas.extend(arvores_area)
            print(f"[LOG] Adicionada área adicional '{area}' com {len(arvores_area)} árvore(s)")
    
    topic_trees = topic_trees_ordenadas
    print(f"[LOG] Árvores reordenadas conforme sequência solicitada: {len(topic_trees)} árvores")
    
    # Reorganizar questões para tópicos de nível 4 (agrupar descendentes)
    def get_all_descendants(topico_id):
        """Retorna todos os descendentes de um tópico (incluindo ele próprio)"""
        descendants = {topico_id}
        
        # Buscar filhos diretos
        filhos = [t_id for t_id, t_info in topicos_info.items() 
                 if t_info['id_pai'] == topico_id]
        
        for filho_id in filhos:
            descendants.update(get_all_descendants(filho_id))
        
        return descendants
    
    def reorganize_questions_for_level4(tree_node, questions_by_topic, reorganized_questions):
        """Reorganiza questões para que tópicos de nível 4 incluam questões de todos os descendentes"""
        
        if tree_node['nivel'] == 4:
            # Este é um tópico de nível 4, coletar questões de todos os descendentes
            all_descendants = get_all_descendants(tree_node['id'])
            todas_questoes = []
            
            for desc_id in all_descendants:
                if desc_id in questions_by_topic:
                    todas_questoes.extend(questions_by_topic[desc_id])
            
            if todas_questoes:
                reorganized_questions[tree_node['id']] = todas_questoes
                print(f"[LOG] Tópico nível 4 '{tree_node['nome']}': {len(todas_questoes)} questões reagrupadas")
            
        elif tree_node['nivel'] < 4:
            # Para níveis menores que 4, manter questões diretas e processar filhos
            if tree_node['id'] in questions_by_topic:
                reorganized_questions[tree_node['id']] = questions_by_topic[tree_node['id']]
            
            # Processar filhos recursivamente
            for child in tree_node['children']:
                reorganize_questions_for_level4(child, questions_by_topic, reorganized_questions)
    
    # Aplicar reorganização
    reorganized_questions = {}
    for tree in topic_trees:
        reorganize_questions_for_level4(tree, questions_by_topic, reorganized_questions)
    
    print(f"[LOG] Questões reorganizadas para {len(reorganized_questions)} tópicos")
    
    # Criar documento
    document = Document()
    
    # Configurar metadados do documento
    configurar_metadados_documento(document, total_questoes)
    
    # Configurar estilo padrão
    style = document.styles['Normal']
    font = style.font
    font.name = 'Calibri'
    font.size = Pt(12)
    paragraph_format = style.paragraph_format
    paragraph_format.space_after = Pt(3)
    paragraph_format.space_before = Pt(0)
    paragraph_format.line_spacing = 1
    
    # === SEÇÃO 1: CAPA ===
    print("[LOG] Criando seção da capa...")
    
    # Configurar cabeçalho da capa com logotipo
    section_capa = document.sections[0]
    section_capa.header.is_linked_to_previous = False
    header_capa = section_capa.header
    for p in header_capa.paragraphs:
        p.clear()
    
    # Adicionar logotipo no cabeçalho (se disponível)
    img_path = os.path.join(os.path.dirname(__file__), 'img', 'logotipo.png')
    p_header = header_capa.paragraphs[0]
    p_header.clear()
    p_header.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    
    logotipo_adicionado = False
    if os.path.exists(img_path):
        print(f"[LOG] Verificando logotipo: {img_path}")
        run_header = p_header.add_run()
        try:
            # Verificar se é uma imagem válida tentando abrir com PIL
            Image.open(img_path).verify()  # Verificar se é uma imagem válida
            
            run_header.add_picture(img_path, width=Inches(3))
            print(f"[LOG] Logotipo adicionado com sucesso")
            logotipo_adicionado = True
        except Exception as e:
            print(f"[AVISO] Arquivo logotipo.png não é uma imagem válida: {str(e)}")
            print(f"[INFO] Substituir img/logotipo.png por uma imagem PNG/JPG real")
    
    if not logotipo_adicionado:
        print(f"[INFO] Cabeçalho da capa criado sem logotipo")
        # Opcional: adicionar texto de placeholder
        # run_header = p_header.add_run("🏥 BANCO DE QUESTÕES MÉDICAS")
        # run_header.bold = True
    
    # Título da capa
    document.add_paragraph("")  # Espaço no topo
    document.add_paragraph("")
    document.add_paragraph("")
    
    capa_title = document.add_paragraph()
    capa_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = capa_title.add_run(f"Banco de Questões - Consulta SQL Específica")
    run.bold = True
    run.font.size = Pt(24)
    
    document.add_paragraph("")
    subtitle = document.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_sub = subtitle.add_run(f"({len(questoes_com_topico)} Questões)")
    run_sub.font.size = Pt(18)
    
    # === SEÇÃO 2: SUMÁRIO ===
    print("[LOG] Criando seção do sumário...")
    document.add_section(WD_SECTION.NEW_PAGE)
    
    # Configurar cabeçalho da seção sumário (sem logotipo)
    section_sumario = document.sections[-1]
    section_sumario.header.is_linked_to_previous = False
    header_sumario = section_sumario.header
    for p in header_sumario.paragraphs:
        p.clear()
    
    # Título do sumário
    sumario_title = document.add_heading("Sumário", level=1)
    sumario_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    document.add_paragraph("")
    toc_paragraph = document.add_paragraph()
    add_toc(toc_paragraph)
    
    # === SEÇÃO 3: CONTEÚDO DAS QUESTÕES ===
    print("[LOG] Criando seção do conteúdo das questões...")
    document.add_section(WD_SECTION.NEW_PAGE)
    
    # Adicionar questões organizadas hierarquicamente
    questao_num = 1
    questoes_adicionadas = set() if not permitir_repeticao else None
    
    # Processar cada árvore de tópicos
    for idx_tree, tree in enumerate(topic_trees, 1):
        print(f"[LOG] Processando árvore {idx_tree}: {tree['nome']}")
        
        # Usar função recursiva para adicionar seções hierárquicas
        questao_num = add_topic_sections_recursive(
            document,
            tree,
            reorganized_questions,
            level=1,
            numbering=[idx_tree],
            parent_names=[],
            questao_num=questao_num,
            breadcrumb_raiz=None,  # Não usar breadcrumb_raiz, usar lógica específica
            permitir_repeticao=permitir_repeticao,
            questoes_adicionadas=questoes_adicionadas,
            total_questoes_banco=total_questoes
        )
    
    # Adicionar rodapé
    add_footer_with_text_and_page_number(document)
    
    # Salvar documento
    data_atual = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_filename = f"banco_questoes_sql_{len(questoes_com_topico)}_{data_atual}.docx"
    
    document.save(output_filename)
    print(f"[LOG] Arquivo {output_filename} gerado com sucesso.")
    print(f"[LOG] Total de questões no banco: {len(questoes_com_topico)}")
    
    return output_filename

def gerar_banco_area_especifica(conn, id_topico_raiz, total_questoes=1000, permitir_repeticao=True):
    """
    Gera um banco de questões de uma área específica baseada no código do tópico raiz.
    
    Args:
        conn: Conexão com o banco de dados
        id_topico_raiz: ID do tópico raiz que define a área específica
        total_questoes: Número total de questões desejadas
        permitir_repeticao: Se permite questões repetidas
    """
    print(f"[LOG] Gerando banco de questões para área específica - Tópico raiz: {id_topico_raiz}, {total_questoes} questões...")
    
    cursor = conn.cursor(dictionary=True)
    
    # Primeiro, verificar se o tópico raiz existe e obter seu nome
    cursor.execute("SELECT id, nome FROM topico WHERE id = %s", (id_topico_raiz,))
    topico_raiz_info = cursor.fetchone()
    
    if not topico_raiz_info:
        print(f"[ERRO] Tópico raiz com ID {id_topico_raiz} não encontrado!")
        return None
    
    nome_area = topico_raiz_info['nome']
    print(f"[LOG] Área selecionada: {nome_area}")
    
    # Informar comportamento de seções baseado no número de questões
    if total_questoes <= 500:
        print(f"[LOG] Banco COMPACTO ({total_questoes} questões): quebras de página apenas para tópicos de NÍVEL 1")
    else:
        print(f"[LOG] Banco EXPANDIDO ({total_questoes} questões): quebras de página para tópicos de NÍVEIS 1, 2 e 3")
    
    # Obter todos os tópicos descendentes do tópico raiz
    cursor.execute("""
        WITH RECURSIVE topicos_descendentes AS (
            SELECT id, nome, id_pai, 1 as nivel
            FROM topico 
            WHERE id = %s
            
            UNION ALL
            
            SELECT t.id, t.nome, t.id_pai, td.nivel + 1
            FROM topico t
            INNER JOIN topicos_descendentes td ON t.id_pai = td.id
            WHERE td.nivel < 10
        )
        SELECT id FROM topicos_descendentes
    """, (id_topico_raiz,))
    
    topicos_area = [row['id'] for row in cursor.fetchall()]
    print(f"[LOG] Encontrados {len(topicos_area)} tópicos na hierarquia da área")
    
    # Buscar questões dos tópicos da área específica
    format_strings = ','.join(['%s'] * len(topicos_area))
    
    query_questoes = f"""
    SELECT DISTINCT
        q.questao_id,
        q.codigo,
        q.ano,
        q.instituicao,
        q.area,
        q.enunciado,
        q.alternativaA,
        q.alternativaB,
        q.alternativaC,
        q.alternativaD,
        q.alternativaE,
        q.gabarito,
        q.gabarito_texto,
        q.comentario,
        q.dificuldade,
        cq.id_topico
    FROM questaoresidencia q
    INNER JOIN classificacao_questao cq ON q.questao_id = cq.id_questao
    WHERE cq.id_topico IN ({format_strings})
      AND q.alternativaE IS NULL
      AND q.comentario IS NOT NULL
      AND CHAR_LENGTH(q.comentario) >= 500
      AND q.ano >= 2020
    ORDER BY SHA2(CONCAT(q.questao_id, 'SEMENTE_FIXA'), 256)
    LIMIT %s
    """
    
    print(f"[LOG] Executando consulta SQL para selecionar questões da área específica...")
    cursor.execute(query_questoes, tuple(topicos_area + [total_questoes]))
    questoes_selecionadas = cursor.fetchall()
    
    print(f"[LOG] Total de questões selecionadas: {len(questoes_selecionadas)}")
    
    if len(questoes_selecionadas) == 0:
        print(f"[ERRO] Nenhuma questão encontrada para o tópico raiz {id_topico_raiz}")
        return None
    
    if len(questoes_selecionadas) < total_questoes:
        print(f"[AVISO] Encontradas apenas {len(questoes_selecionadas)} questões de {total_questoes} solicitadas")
        print(f"[AVISO] Isso pode indicar que não há questões suficientes na área que atendam aos critérios")
    
    # Obter classificações mais específicas para melhor organização
    questao_ids = [q['questao_id'] for q in questoes_selecionadas]
    format_strings = ','.join(['%s'] * len(questao_ids))
    
    query_topicos_especificos = f"""
    SELECT DISTINCT cq.id_topico, cq.id_questao
    FROM classificacao_questao cq
    WHERE cq.id_questao IN ({format_strings})
    ORDER BY cq.id_questao, cq.id_topico
    """
    
    cursor.execute(query_topicos_especificos, tuple(questao_ids))
    classificacoes_especificas = cursor.fetchall()
    
    # Criar mapeamento de questão -> tópicos específicos
    questao_topicos_especificos = {}
    for classificacao in classificacoes_especificas:
        questao_id = classificacao['id_questao']
        topico_id = classificacao['id_topico']
        if questao_id not in questao_topicos_especificos:
            questao_topicos_especificos[questao_id] = []
        questao_topicos_especificos[questao_id].append(topico_id)
    
    # Usar tópico mais específico se disponível
    for q in questoes_selecionadas:
        topicos_especificos = questao_topicos_especificos.get(q['questao_id'], [])
        if topicos_especificos:
            # Usar o primeiro tópico específico que está dentro da área
            for topico_esp in topicos_especificos:
                if topico_esp in topicos_area:
                    q['id_topico'] = topico_esp
                    break
    
    questoes_com_topico = questoes_selecionadas
    print(f"[LOG] Questões com tópico associado: {len(questoes_com_topico)}")
    
    # Organizar questões por tópico
    questions_by_topic = {}
    for q in questoes_com_topico:
        tid = q['id_topico']
        if tid not in questions_by_topic:
            questions_by_topic[tid] = []
        questions_by_topic[tid].append(q)
    
    print(f"[LOG] Questões organizadas por {len(questions_by_topic)} tópicos")
    
    # Construir hierarquia completa dos tópicos utilizados
    print("[LOG] Construindo hierarquia completa dos tópicos...")
    
    topicos_utilizados = list(set([q['id_topico'] for q in questoes_com_topico]))
    topicos_completos = set(topicos_utilizados)
    
    # Para cada tópico utilizado, buscar todos os ancestrais
    for topico_id in topicos_utilizados:
        cursor.execute("""
            WITH RECURSIVE topico_ancestrais AS (
                SELECT id, id_pai, nome, 1 as nivel
                FROM topico 
                WHERE id = %s
                
                UNION ALL
                
                SELECT t.id, t.id_pai, t.nome, ta.nivel + 1
                FROM topico t
                INNER JOIN topico_ancestrais ta ON ta.id_pai = t.id
                WHERE ta.nivel < 10
            )
            SELECT id FROM topico_ancestrais
        """, (topico_id,))
        
        ancestrais = cursor.fetchall()
        for ancestral in ancestrais:
            topicos_completos.add(ancestral['id'])
    
    print(f"[LOG] Tópicos completos (incluindo ancestrais): {len(topicos_completos)}")
    
    # Buscar informações completas dos tópicos
    topicos_completos_list = list(topicos_completos)
    format_strings = ','.join(['%s'] * len(topicos_completos_list))
    
    cursor.execute(f"""
        SELECT id, nome, id_pai
        FROM topico 
        WHERE id IN ({format_strings})
        ORDER BY id
    """, tuple(topicos_completos_list))
    
    topicos_info = {t['id']: t for t in cursor.fetchall()}
    
    # Construir árvore hierárquica a partir do tópico raiz
    def build_topic_tree(topico_id, nivel_atual=1, max_nivel=4):
        if topico_id not in topicos_info:
            return None
            
        topico = topicos_info[topico_id]
        tree_node = {
            'id': topico_id,
            'nome': topico['nome'],
            'nivel': nivel_atual,
            'children': []
        }
        
        if nivel_atual >= max_nivel:
            return tree_node
        
        # Encontrar filhos diretos
        filhos = [t_id for t_id, t_info in topicos_info.items() 
                 if t_info['id_pai'] == topico_id and t_id in topicos_completos]
        
        for filho_id in sorted(filhos):
            child_tree = build_topic_tree(filho_id, nivel_atual + 1, max_nivel)
            if child_tree:
                tree_node['children'].append(child_tree)
        
        return tree_node
    
    # Construir árvore a partir do tópico raiz
    topic_tree = build_topic_tree(id_topico_raiz)
    
    if not topic_tree:
        print(f"[ERRO] Não foi possível construir hierarquia para o tópico raiz {id_topico_raiz}")
        return None
    
    print(f"[LOG] Árvore hierárquica construída a partir do tópico raiz: {topic_tree['nome']}")
    
    # Reorganizar questões para tópicos de nível 4
    def get_all_descendants(topico_id):
        """Retorna todos os descendentes de um tópico (incluindo ele próprio)"""
        descendants = {topico_id}
        
        filhos = [t_id for t_id, t_info in topicos_info.items() 
                 if t_info['id_pai'] == topico_id]
        
        for filho_id in filhos:
            descendants.update(get_all_descendants(filho_id))
        
        return descendants
    
    def reorganize_questions_for_level4(tree_node, questions_by_topic, reorganized_questions):
        """Reorganiza questões para que tópicos de nível 4 incluam questões de todos os descendentes"""
        
        if tree_node['nivel'] == 4:
            all_descendants = get_all_descendants(tree_node['id'])
            todas_questoes = []
            
            for desc_id in all_descendants:
                if desc_id in questions_by_topic:
                    todas_questoes.extend(questions_by_topic[desc_id])
            
            if todas_questoes:
                reorganized_questions[tree_node['id']] = todas_questoes
                print(f"[LOG] Tópico nível 4 '{tree_node['nome']}': {len(todas_questoes)} questões reagrupadas")
            
        elif tree_node['nivel'] < 4:
            if tree_node['id'] in questions_by_topic:
                reorganized_questions[tree_node['id']] = questions_by_topic[tree_node['id']]
            
            for child in tree_node['children']:
                reorganize_questions_for_level4(child, questions_by_topic, reorganized_questions)
    
    # Aplicar reorganização
    reorganized_questions = {}
    reorganize_questions_for_level4(topic_tree, questions_by_topic, reorganized_questions)
    
    print(f"[LOG] Questões reorganizadas para {len(reorganized_questions)} tópicos")
    
    # Criar documento
    document = Document()
    
    # Configurar metadados do documento
    configurar_metadados_documento(document, len(questoes_com_topico))
    
    # Configurar estilo padrão
    style = document.styles['Normal']
    font = style.font
    font.name = 'Calibri'
    font.size = Pt(12)
    paragraph_format = style.paragraph_format
    paragraph_format.space_after = Pt(3)
    paragraph_format.space_before = Pt(0)
    paragraph_format.line_spacing = 1
    
    # === SEÇÃO 1: CAPA ===
    print("[LOG] Criando seção da capa...")
    
    # Configurar cabeçalho da capa com logotipo
    section_capa = document.sections[0]
    section_capa.header.is_linked_to_previous = False
    header_capa = section_capa.header
    for p in header_capa.paragraphs:
        p.clear()
    
    # Adicionar logotipo no cabeçalho (se disponível)
    img_path = os.path.join(os.path.dirname(__file__), 'img', 'logotipo.png')
    p_header = header_capa.paragraphs[0]
    p_header.clear()
    p_header.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    
    logotipo_adicionado = False
    if os.path.exists(img_path):
        print(f"[LOG] Verificando logotipo: {img_path}")
        run_header = p_header.add_run()
        try:
            Image.open(img_path).verify()
            run_header.add_picture(img_path, width=Inches(3))
            print(f"[LOG] Logotipo adicionado com sucesso")
            logotipo_adicionado = True
        except Exception as e:
            print(f"[AVISO] Arquivo logotipo.png não é uma imagem válida: {str(e)}")
    
    if not logotipo_adicionado:
        print(f"[INFO] Cabeçalho da capa criado sem logotipo")
    
    # Título da capa
    document.add_paragraph("")
    document.add_paragraph("")
    document.add_paragraph("")
    
    capa_title = document.add_paragraph()
    capa_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = capa_title.add_run(f"Banco de Questões - {nome_area}")
    run.bold = True
    run.font.size = Pt(24)
    
    document.add_paragraph("")
    subtitle = document.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_sub = subtitle.add_run(f"({len(questoes_com_topico)} Questões)")
    run_sub.font.size = Pt(18)
    
    # === SEÇÃO 2: SUMÁRIO ===
    print("[LOG] Criando seção do sumário...")
    document.add_section(WD_SECTION.NEW_PAGE)
    
    section_sumario = document.sections[-1]
    section_sumario.header.is_linked_to_previous = False
    header_sumario = section_sumario.header
    for p in header_sumario.paragraphs:
        p.clear()
    
    sumario_title = document.add_heading("Sumário", level=1)
    sumario_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    document.add_paragraph("")
    toc_paragraph = document.add_paragraph()
    add_toc(toc_paragraph)
    
    # === SEÇÃO 3: CONTEÚDO DAS QUESTÕES ===
    print("[LOG] Criando seção do conteúdo das questões...")
    document.add_section(WD_SECTION.NEW_PAGE)
    
    # Adicionar questões organizadas hierarquicamente
    questao_num = 1
    questoes_adicionadas = set() if not permitir_repeticao else None
    
    # Processar a árvore de tópicos
    questao_num = add_topic_sections_recursive(
        document,
        topic_tree,
        reorganized_questions,
        level=1,
        numbering=[1],
        parent_names=[],
        questao_num=questao_num,
        breadcrumb_raiz=None,
        permitir_repeticao=permitir_repeticao,
        questoes_adicionadas=questoes_adicionadas,
        total_questoes_banco=len(questoes_com_topico)
    )
    
    # Adicionar rodapé
    add_footer_with_text_and_page_number(document)
    
    # Salvar documento
    data_atual = datetime.now().strftime("%Y%m%d_%H%M%S")
    nome_arquivo_limpo = nome_area.replace(" ", "_").replace("/", "_")
    output_filename = f"banco_questoes_{nome_arquivo_limpo}_{len(questoes_com_topico)}_{data_atual}.docx"
    
    document.save(output_filename)
    print(f"[LOG] Arquivo {output_filename} gerado com sucesso.")
    print(f"[LOG] Total de questões no banco: {len(questoes_com_topico)}")
    
    return output_filename

if __name__ == "__main__":
    print("=== GERADOR DE BANCO DE QUESTÕES MÉDICAS ===")
    print()
    print("Escolha o modo de geração:")
    print("1 - Banco completo com 6 áreas médicas (Modo original)")
    print("2 - Banco de área específica por tópico raiz")
    print()
    
    # Escolher modo de operação
    try:
        modo = int(input("Digite sua opção (1 ou 2): "))
        if modo not in [1, 2]:
            print("Erro: Opção inválida! Digite 1 ou 2.")
            exit(1)
    except ValueError:
        print("Erro: Digite um número válido (1 ou 2)!")
        exit(1)
    
    # Solicitar número total de questões
    try:
        N = int(input("Número total de questões do banco (ex: 1000, 2000, 3000): "))
        if N <= 0:
            print("Erro: N deve ser um número positivo!")
            exit(1)
    except ValueError:
        print("Erro: N deve ser um número inteiro!")
        exit(1)
    
    # Configurar permitir repetição (fixo como False para evitar questões duplicadas)
    permitir_repeticao = False
    
    # Conectar ao banco
    conn = get_connection()
    print("[LOG] Conexão com o banco estabelecida.")
    
    if modo == 1:
        # MODO 1: Banco com 6 áreas médicas (modo original)
        print(f"\n[LOG] MODO 1: Gerando banco com 6 áreas médicas")
        print(f"[LOG] Distribuição proporcional para {N} questões:")
        print(f"  1. Cirurgia: {round(N * 0.2)} questões (20%)")
        print(f"  2. Clínica Médica: {round(N * 0.2)} questões (20%)")
        print(f"  3. Pediatria: {round(N * 0.2)} questões (20%)")
        print(f"  4. Ginecologia: {round(N * 0.1)} questões (10%)")
        print(f"  5. Obstetrícia: {round(N * 0.1)} questões (10%)")
        print(f"  6. Medicina Preventiva: {round(N * 0.2)} questões (20%)")
        print()
        
        gerar_banco_estratificacao_deterministica(conn, N, permitir_repeticao=permitir_repeticao)
        
    elif modo == 2:
        # MODO 2: Banco de área específica
        print(f"\n[LOG] MODO 2: Gerando banco de área específica")
        print()
        print("Códigos dos tópicos raiz das principais áreas:")
        print("  33  - Cirurgia")
        print("  100 - Clínica Médica")
        print("  48  - Pediatria")
        print("  183 - Ginecologia")
        print("  218 - Obstetrícia")
        print("  29  - Medicina Preventiva")
        print()
        print("Ou informe o código de qualquer outro tópico raiz desejado.")
        print()
        
        try:
            id_topico_raiz = int(input("Digite o código do tópico raiz: "))
            if id_topico_raiz <= 0:
                print("Erro: O código do tópico deve ser um número positivo!")
                exit(1)
        except ValueError:
            print("Erro: Digite um código válido (número inteiro)!")
            exit(1)
        
        print(f"[LOG] Tópico raiz selecionado: {id_topico_raiz}")
        print(f"[LOG] Gerando {N} questões da área específica...")
        print()
        
        resultado = gerar_banco_area_especifica(conn, id_topico_raiz, N, permitir_repeticao=permitir_repeticao)
        
        if not resultado:
            print("[ERRO] Falha na geração do banco de questões!")
            conn.close()
            exit(1)
    
    conn.close()
    print("\n[LOG] Processo concluído!")
