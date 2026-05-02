import codecs
import re

filepath = r"c:\Users\elman\git\gerador-banco-questoes\geradorBancosDeQuestoesPorTopico.py"
with codecs.open(filepath, 'r', 'utf-8') as f:
    text = f.read()

# Chunk 1: mapear_assunto_hierarquicamente
chunk_1_target = """    finally:
        cursor.close()

def processar_classificacao_questoes_sem_topico"""

chunk_1_replace = """    finally:
        cursor.close()

def mapear_assunto_hierarquicamente(assunto_edital, topicos_dict, topicos_raiz_ids):
    \"\"\"
    Mapeia um assunto de edital para a hierarquia de tópicos do banco.
    \"\"\"
    assunto_limpo = assunto_edital.strip()[:1000]

    caminho_topicos = []
    opcoes_atual = list(topicos_raiz_ids)
    nivel = 1
    visitados = set()
    mapeamento_completo = True

    while opcoes_atual:
        lista_opcoes, mapa_indice = montar_lista_opcoes(topicos_dict, opcoes_atual)
        if not mapa_indice:
            print("[AVISO] Lista de tópicos vazia durante o mapeamento.")
            mapeamento_completo = False
            break

        prompt_base = (
            "Você é um classificador estruturado. "
            "Avalie o seguinte assunto/tema extraído de um edital médico e determine em qual dos tópicos listados ele melhor se encaixa. "
            "Se o assunto for amplo demais, ou cobrir vários dos subtópicos, ou nenhum parecer adequado, você pode ter chegado ao nível correto. "
            "No entanto, se um sub-tópico for CLARAMENTE a melhor correspondência, escolha-o. "
            "Se for absolutamente impossível encaixar o tema em qualquer um dos tópicos, responda 0.\\n\\n"
            f"[ASSUNTO DO EDITAL]: {assunto_limpo}\\n\\n"
            f"[TÓPICOS NÍVEL {nivel}]:\\n{lista_opcoes}\\n\\n"
            "Responda APENAS com o número correspondente."
        )

        numero_escolhido = None
        for tentativa in range(2):
            prompt = prompt_base if tentativa == 0 else (
                prompt_base +
                f"\\nResponda APENAS com um número. Opções válidas: 0, {', '.join(str(i) for i in mapa_indice.keys())}."
            )

            resposta = deepseek_chat(
                [{"role": "user", "content": prompt}],
                max_tokens=10
            )

            if not resposta:
                print(f"[AVISO] Sem resposta da IA para mapeamento do assunto '{assunto_limpo}'.")
                numero_escolhido = None
                break

            numero = extrair_primeiro_inteiro(resposta)
            
            if numero == 0:
                numero_escolhido = 0
                break
                
            if numero in mapa_indice:
                numero_escolhido = numero
                break

            print(f"[AVISO] Resposta inválida: '{resposta}'. Tentativa {tentativa + 1}/2.")

        if numero_escolhido is None or numero_escolhido == 0:
            if caminho_topicos:
                print(f"[LOG] Mapeamento concluído no nível {nivel-1} para '{assunto_limpo}'.")
            else:
                print(f"[AVISO] Não foi possível encontrar mapeamento para '{assunto_limpo}'.")
                mapeamento_completo = False
            break

        topico_id = mapa_indice[numero_escolhido]
        nome_topico = topicos_dict[topico_id]['nome']
        print(f"[LOG] Nível {nivel}: escolhido tópico {numero_escolhido} -> ID {topico_id} ({nome_topico})")

        if topico_id in visitados:
            print(f"[AVISO] Ciclo detectado (ID {topico_id}).")
            mapeamento_completo = False
            break

        caminho_topicos.append(topico_id)
        visitados.add(topico_id)

        filhos = topicos_dict[topico_id]['filhos']
        if not filhos:
            break

        opcoes_atual = filhos
        nivel += 1

    return caminho_topicos, mapeamento_completo

def processar_classificacao_questoes_sem_topico"""

text = text.replace(chunk_1_target, chunk_1_replace)

# Chunk 2: gerar_banco_por_edital
chunk_2_target = """    return output_filename

if __name__ == "__main__":"""

chunk_2_replace = """    return output_filename

def gerar_banco_por_edital(conn, caminho_edital, nome_concurso, N, ano_minimo=2016, tamanho_minimo_comentario=500, incluir_comentarios=True):
    import os
    print(f"\\n[LOG] Lendo assuntos do edital em: {caminho_edital}")
    
    if not os.path.exists(caminho_edital):
        print(f"[ERRO] Arquivo não encontrado: {caminho_edital}")
        return None
        
    try:
        with open(caminho_edital, 'r', encoding='utf-8') as f:
            assuntos = [linha.strip() for linha in f if linha.strip()]
    except Exception as e:
        print(f"[ERRO] Falha ao ler arquivo: {e}")
        return None
        
    if not assuntos:
        print("[ERRO] Arquivo do edital está vazio.")
        return None
        
    print(f"[LOG] Total de {len(assuntos)} assuntos encontrados.")
    
    topicos_dict, topicos_raiz = carregar_hierarquia_topicos(conn)
    if not topicos_raiz:
        print("[ERRO] Não foi possível carregar a hierarquia de tópicos.")
        return None
        
    topicos_selecionados = {}
    
    for idx, assunto in enumerate(assuntos, 1):
        print(f"\\n[LOG] ({idx}/{len(assuntos)}) Mapeando: '{assunto}'")
        caminho_topicos, _ = mapear_assunto_hierarquicamente(assunto, topicos_dict, topicos_raiz)
        
        if caminho_topicos:
            topico_alvo = caminho_topicos[-1]
            topicos_selecionados[assunto] = topico_alvo
            print(f"[SUCESSO] '{assunto}' -> Tópico ID {topico_alvo} ({topicos_dict[topico_alvo]['nome']})")
        else:
            print(f"[AVISO] Ignorando '{assunto}' pois não houve mapeamento correspondente.")
            
    ids_unicos = list(set(topicos_selecionados.values()))
    
    if not ids_unicos:
        print("[ERRO] Nenhum assunto mapeado com sucesso para buscar questões.")
        return None
        
    print(f"\\n[LOG] Total de tópicos base mapeados: {len(ids_unicos)}")
    
    cursor = conn.cursor(dictionary=True)
    topicos_e_descendentes = set()
    
    for topico_id in ids_unicos:
        cursor.execute(\"\"\"
            WITH RECURSIVE topico_descendentes AS (
                SELECT id, id_pai, nome, 1 as nivel
                FROM topico 
                WHERE id = %s
                UNION ALL
                SELECT t.id, t.id_pai, t.nome, td.nivel + 1
                FROM topico t
                INNER JOIN topico_descendentes td ON t.id_pai = t.id
                WHERE td.nivel < 10
            )
            SELECT id FROM topico_descendentes
        \"\"\", (topico_id,))
        for desc in cursor.fetchall():
            topicos_e_descendentes.add(desc['id'])
            
    print(f"[LOG] Tópicos expandidos com descendentes: {len(topicos_e_descendentes)} tópicos no total.")
    
    formato_ids = ','.join(['%s'] * len(topicos_e_descendentes))
    
    query = f\"\"\"
        SELECT q.*, cq.id_topico
        FROM questaoresidencia q
        INNER JOIN classificacao_questao cq ON q.questao_id = cq.id_questao
        WHERE cq.id_topico IN ({formato_ids})
          AND q.ano >= %s
          AND (CHAR_LENGTH(q.comentario) >= %s OR (q.gabaritoIA=q.gabarito AND q.comentarioIA IS NOT NULL))
        ORDER BY RAND()
    \"\"\"
    params = list(topicos_e_descendentes) + [ano_minimo, tamanho_minimo_comentario]
    
    print(f"[LOG] Buscando questões no banco de dados...")
    cursor.execute(query, tuple(params))
    questoes = cursor.fetchall()
    
    if not questoes:
        print("[ERRO] Nenhuma questão encontrada para os tópicos do edital com os critérios informados.")
        return None
        
    ids_adicionados = set()
    questoes_finais = []
    
    for q in questoes:
        if q['questao_id'] not in ids_adicionados:
            questoes_finais.append(q)
            ids_adicionados.add(q['questao_id'])
            if len(questoes_finais) >= N:
                break
                
    if len(questoes_finais) < N:
        print(f"[AVISO] Foram encontradas apenas {len(questoes_finais)} questões, embora o objetivo fosse {N}.")
        
    print(f"[LOG] {len(questoes_finais)} questões selecionadas e balanceadas.")
    
    questions_by_topic = {}
    for q in questoes_finais:
        tid = q['id_topico']
        if tid not in questions_by_topic:
            questions_by_topic[tid] = []
        questions_by_topic[tid].append(q)
        
    topicos_utilizados = list(questions_by_topic.keys())
    topicos_completos = set(topicos_utilizados)
    
    print("[LOG] Resolvendo cadeias de árvores para a formatação do documento...")
    for topico_id in topicos_utilizados:
        cursor.execute(\"\"\"
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
        \"\"\", (topico_id,))
        for anc in cursor.fetchall():
            topicos_completos.add(anc['id'])
            
    topicos_completos_list = list(topicos_completos)
    format_strings = ','.join(['%s'] * len(topicos_completos_list))
    
    cursor.execute(f\"\"\"
        SELECT id, nome, id_pai
        FROM topico 
        WHERE id IN ({format_strings})
        ORDER BY id
    \"\"\", tuple(topicos_completos_list))
    
    topicos_info = {t['id']: t for t in cursor.fetchall()}
    
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
        filhos = [t_id for t_id, t_info in topicos_info.items() 
                 if t_info['id_pai'] == topico_id and t_id in topicos_completos]
        for filho_id in sorted(filhos):
            child_tree = build_topic_tree(filho_id, nivel_atual + 1, max_nivel)
            if child_tree:
                tree_node['children'].append(child_tree)
        return tree_node

    topicos_raiz = []
    for topico_id in topicos_completos:
        if topico_id not in topicos_info:
            continue
        topico = topicos_info[topico_id]
        if topico['id_pai'] is None or topico['id_pai'] not in topicos_completos:
            topicos_raiz.append(topico_id)
            
    topic_trees = []
    for raiz_id in sorted(topicos_raiz):
        tree = build_topic_tree(raiz_id)
        if tree:
            topic_trees.append(tree)
            
    reorganized_questions = {}
    
    def reorganize_questions_for_level4(tree_node, questions_by_topic, reorganized_questions):
        if tree_node['nivel'] == 4:
            todas_questoes = []
            questoes_ids_unicos = set()
            
            def get_all_descendants(tid):
                descendants = {tid}
                filhos = [t_id for t_id, t_info in topicos_info.items() if t_info['id_pai'] == tid]
                for filho_id in filhos:
                    descendants.update(get_all_descendants(filho_id))
                return descendants
                
            all_descendants = get_all_descendants(tree_node['id'])
            for desc_id in all_descendants:
                if desc_id in questions_by_topic:
                    for questao in questions_by_topic[desc_id]:
                        if questao['questao_id'] not in questoes_ids_unicos:
                            todas_questoes.append(questao)
                            questoes_ids_unicos.add(questao['questao_id'])
            if todas_questoes:
                reorganized_questions[tree_node['id']] = todas_questoes
        elif tree_node['nivel'] < 4:
            if tree_node['id'] in questions_by_topic:
                reorganized_questions[tree_node['id']] = questions_by_topic[tree_node['id']]
            for child in tree_node['children']:
                reorganize_questions_for_level4(child, questions_by_topic, reorganized_questions)
                
    for tree in topic_trees:
        reorganize_questions_for_level4(tree, questions_by_topic, reorganized_questions)
        
    print("[LOG] Criando documento DocX do Edital...")
    document = Document()
    
    nome_titulo_instituicao = limpar_nome_para_titulo(nome_concurso)
    configurar_metadados_documento(document, len(questoes_finais), nome_titulo_instituicao)
    
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
    header_capa = section_capa.header
    for p in header_capa.paragraphs: p.clear()
    
    img_path = os.path.join(os.path.dirname(__file__), 'img', 'logotipo.png')
    p_header = header_capa.paragraphs[0]
    p_header.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    if os.path.exists(img_path):
        try:
            run_header = p_header.add_run()
            Image.open(img_path).verify()
            run_header.add_picture(img_path, width=Inches(3))
        except:
            pass
            
    for _ in range(3): document.add_paragraph("")
    
    capa_title = document.add_paragraph()
    capa_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = capa_title.add_run(f"Banco de Questões - {nome_concurso}")
    run.bold = True
    run.font.size = Pt(24)
    
    document.add_paragraph("")
    subtitle = document.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_sub = subtitle.add_run(f"({len(questoes_finais)} Questões - Baseado no Edital)")
    run_sub.font.size = Pt(18)
    
    document.add_section(WD_SECTION.NEW_PAGE)
    section_sumario = document.sections[-1]
    section_sumario.header.is_linked_to_previous = False
    for p in section_sumario.header.paragraphs: p.clear()
    
    sumario_title = document.add_heading("Sumário", level=1)
    sumario_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    document.add_paragraph("")
    add_toc(document.add_paragraph())
    
    document.add_section(WD_SECTION.NEW_PAGE)
    questao_num = 1
    
    for idx_tree, tree in enumerate(topic_trees, 1):
        questao_num = add_topic_sections_recursive(
            document,
            tree,
            reorganized_questions,
            level=1,
            numbering=[idx_tree],
            parent_names=[],
            questao_num=questao_num,
            breadcrumb_raiz=None,
            permitir_repeticao=False,
            questoes_adicionadas=set(),
            total_questoes_banco=len(questoes_finais),
            incluir_comentarios=incluir_comentarios
        )
        
    add_footer_with_text_and_page_number(document)
    
    data_atual = datetime.now().strftime("%Y%m%d_%H%M%S")
    nome_arquivo_limpo = nome_titulo_instituicao.replace(" ", "_").upper()
    output_filename = f"banco_questoes_{nome_arquivo_limpo}_{len(questoes_finais)}_{data_atual}.docx"
    
    document.save(output_filename)
    print(f"[LOG] Arquivo {output_filename} gerado com sucesso!")
    
    return output_filename

if __name__ == "__main__":"""

text = text.replace(chunk_2_target, chunk_2_replace)

# Chunk 3: Menu and inputs
chunk_3_target_1 = """    print("5 - Responder questões usando a IA (DeepSeek)")
    print("6 - Classificar questões sem tópico (DeepSeek AI)")
    print()
    
    # Escolher modo de operação
    try:
        modo = int(input("Digite sua opção (1, 2, 3, 4, 5 ou 6): "))
        if modo not in [1, 2, 3, 4, 5, 6]:
            print("Erro: Opção inválida! Digite 1, 2, 3, 4, 5 ou 6.")
            exit(1)
    except ValueError:
        print("Erro: Digite um número válido (1, 2, 3, 4, 5 ou 6)!")
        exit(1)"""

chunk_3_replace_1 = """    print("5 - Responder questões usando a IA (DeepSeek)")
    print("6 - Classificar questões sem tópico (DeepSeek AI)")
    print("7 - Banco a partir de Edital (Lista de Assuntos via DeepSeek)")
    print()
    
    # Escolher modo de operação
    try:
        modo = int(input("Digite sua opção (1 a 7): "))
        if modo not in [1, 2, 3, 4, 5, 6, 7]:
            print("Erro: Opção inválida! Digite de 1 a 7.")
            exit(1)
    except ValueError:
        print("Erro: Digite um número válido!")
        exit(1)"""

text = text.replace(chunk_3_target_1, chunk_3_replace_1)

chunk_3_target_2 = """    # Solicitar ano mínimo para modos 1, 2 e 3
    ano_minimo = None
    if modo in [1, 2, 3]:
        try:
            ano_minimo = int(input("Ano mínimo para filtrar as questões (ex: 2016, 2018, 2020): "))
            if ano_minimo < 1900 or ano_minimo > 2100:
                print("Erro: O ano deve ser um valor razoável (entre 1900 e 2100)!")
                exit(1)
        except ValueError:
            print("Erro: O ano deve ser um número inteiro!")
            exit(1)
    
    # Solicitar tamanho mínimo de comentário para modos 1, 2 e 3
    tamanho_minimo_comentario = 500
    if modo in [1, 2, 3]:"""

chunk_3_replace_2 = """    # Solicitar ano mínimo para modos 1, 2, 3 e 7
    ano_minimo = None
    if modo in [1, 2, 3, 7]:
        try:
            ano_minimo = int(input("Ano mínimo para filtrar as questões (ex: 2016, 2018, 2020): "))
            if ano_minimo < 1900 or ano_minimo > 2100:
                print("Erro: O ano deve ser um valor razoável (entre 1900 e 2100)!")
                exit(1)
        except ValueError:
            print("Erro: O ano deve ser um número inteiro!")
            exit(1)
    
    # Solicitar tamanho mínimo de comentário para modos 1, 2, 3 e 7
    tamanho_minimo_comentario = 500
    if modo in [1, 2, 3, 7]:"""

text = text.replace(chunk_3_target_2, chunk_3_replace_2)

# Chunk 4: execution hook the end of Modo 6
chunk_4_target = """            processar_classificacao_questoes_sem_topico(
                conn,
                limite=limite,
                filtro_instituicao=filtro_instituicao,
                resto_mod5=resto_mod5,
                filtro_ano=filtro_ano,
                filtro_prova=filtro_prova
            )
    
    conn.close()
    print("\\n[LOG] Processo concluído!")"""

chunk_4_replace = """            processar_classificacao_questoes_sem_topico(
                conn,
                limite=limite,
                filtro_instituicao=filtro_instituicao,
                resto_mod5=resto_mod5,
                filtro_ano=filtro_ano,
                filtro_prova=filtro_prova
            )
            
    elif modo == 7:
        # MODO 7: Banco baseado na lista do edital
        print(f"\\n[LOG] MODO 7: Gerando banco a partir de temas de um edital")
        print()
        
        caminho_edital = os.path.join(os.path.dirname(__file__), 'edital.txt')
        
        nome_concurso = ""
        while not nome_concurso:
            nome_concurso = input("Informe o nome do concurso (ex: Concurso EBSERH) para a capa: ").strip()
            if not nome_concurso:
                print("Erro: O nome não pode ser vazio!")
                
        # Perguntar se deve incluir comentários
        print("Opções de geração do documento:")
        print("1) Apenas questões (sem comentários)")
        print("2) Questões com comentários (padrão)")
        opcao_comentarios = input("Escolha a opção (1 ou 2, padrão 2): ").strip()
        incluir_comentarios = opcao_comentarios != '1'
        
        print(f"[LOG] Iniciando fluxo de dados do edital.")
        resultado = gerar_banco_por_edital(conn, caminho_edital, nome_concurso, N, ano_minimo=ano_minimo, tamanho_minimo_comentario=tamanho_minimo_comentario, incluir_comentarios=incluir_comentarios)
        
        if not resultado:
            print("\\n[ERRO] Falha na geração do banco de questões por edital!")
            conn.close()
            exit(1)
    
    conn.close()
    print("\\n[LOG] Processo concluído!")"""

text = text.replace(chunk_4_target, chunk_4_replace)

with codecs.open(filepath, 'w', 'utf-8') as f:
    f.write(text)

print("Patch applied successfully.")
