import os
import sys
import mysql.connector
import json
import time

# Ensure stdout uses UTF-8 if possible, or just replace emojis
def safe_print(msg):
    try:
        print(msg)
    except UnicodeEncodeError:
        try:
            print(msg.encode('ascii', errors='replace').decode('ascii'))
        except Exception:
            pass

def main():
    safe_print("=== EXPORTADOR DE HIERARQUIA DE TOPICOS INTERATIVO ===")
    
    # 1. Configuração e Conexão com o Banco de Dados
    db_config = {
        "host": "localhost",
        "user": "root",
        "password": "El@mysql.32",
        "database": "qconcursos"
    }
    
    safe_print("[LOG] Conectando ao banco de dados MySQL...")
    try:
        conn = mysql.connector.connect(**db_config)
        cur = conn.cursor(dictionary=True)
        safe_print("[LOG] Conexao estabelecida com sucesso! [OK]")
    except Exception as e:
        safe_print(f"[ERRO] Falha ao conectar ao banco de dados: {e}")
        sys.exit(1)
        
    start_time = time.time()
    
    # 2. Carregar todos os tópicos
    safe_print("[LOG] Carregando todos os topicos da tabela 'topico'...")
    cur.execute("SELECT id, id_pai, nome FROM topico")
    topicos_rows = cur.fetchall()
    safe_print(f"[LOG] {len(topicos_rows)} topicos carregados.")
    
    # Inicializar dicionário de tópicos
    topicos = {}
    for r in topicos_rows:
        topicos[r["id"]] = {
            "id": r["id"],
            "pai": r["id_pai"],
            "nome": r["nome"],
            "filhos": [],
            "direct_count": 0,
            "total_count": 0,
            "depth": 1
        }
        
    # 3. Carregar contagem de questões ativas por tópico
    safe_print("[LOG] Calculando contagem de questoes por topico (tabela 'questaoresidencia' + 'classificacao_questao')...")
    cur.execute("""
        SELECT cq.id_topico, COUNT(DISTINCT cq.id_questao) as qtd
        FROM classificacao_questao cq
        INNER JOIN questaoresidencia q ON cq.id_questao = q.questao_id
        GROUP BY cq.id_topico
    """)
    question_counts = cur.fetchall()
    
    total_active_questions = 0
    for qc in question_counts:
        tid = qc["id_topico"]
        if tid in topicos:
            topicos[tid]["direct_count"] = qc["qtd"]
            total_active_questions += qc["qtd"]
            
    safe_print(f"[LOG] Contagem concluida. Total de classificacoes ativas encontradas: {total_active_questions}")
    
    # 4. Montar a hierarquia e calcular profundidades
    safe_print("[LOG] Estruturando a arvore hierarquica...")
    raizes = []
    for tid, node in topicos.items():
        pai_id = node["pai"]
        if pai_id is not None and pai_id in topicos:
            topicos[pai_id]["filhos"].append(tid)
        else:
            raizes.append(tid)
            
    # Função para calcular recursivamente profundidade e contagem cumulativa
    max_depth = 1
    
    def process_node(node_id, current_depth):
        nonlocal max_depth
        node = topicos[node_id]
        node["depth"] = current_depth
        if current_depth > max_depth:
            max_depth = current_depth
            
        # Cumulative question count starts with the direct questions of this node
        cumulative = node["direct_count"]
        
        # Process children
        for child_id in node["filhos"]:
            cumulative += process_node(child_id, current_depth + 1)
            
        node["total_count"] = cumulative
        return cumulative
        
    for r_id in raizes:
        process_node(r_id, 1)
        
    safe_print(f"[LOG] Arvore processada. Profundidade maxima encontrada: {max_depth}")
    
    # 5. Ordenar as raízes de acordo com as 7 áreas prioritárias do projeto
    prioridade_areas = {
        "Cirurgia": 1,
        "Clínica Médica": 2,
        "Pediatria": 3,
        "Ginecologia": 4,
        "Obstetrícia": 5,
        "Medicina Preventiva": 6,
        "Outros": 7
    }
    
    def get_raiz_priority(rid):
        nome = topicos[rid]["nome"]
        # Match matches or prefixes
        for area, prio in prioridade_areas.items():
            if area.lower() in nome.lower():
                return prio
        return 99 # outros
        
    # Ordenar raízes pela prioridade, e em seguida em ordem alfabética do nome
    raizes_ordenadas = sorted(raizes, key=lambda rid: (get_raiz_priority(rid), topicos[rid]["nome"].lower()))
    
    # 6. Gerar a representação HTML recursiva
    safe_print("[LOG] Renderizando estrutura de topicos em HTML...")
    
    # Usaremos uma lista de strings para eficiência
    html_tree_parts = []
    
    # Cores associadas às profundidades para a linha guia lateral
    depth_colors = {
        1: "var(--depth-1)",
        2: "var(--depth-2)",
        3: "var(--depth-3)",
        4: "var(--depth-4)",
        5: "var(--depth-5)",
    }
    
    def render_tree_to_html(node_id):
        node = topicos[node_id]
        nome = node["nome"]
        direct = node["direct_count"]
        total = node["total_count"]
        depth = node["depth"]
        filhos = node["filhos"]
        
        # Classes CSS adicionais baseadas no estado
        has_questions = "has-questions" if total > 0 else "no-questions"
        direct_badge = f'<span class="badge direct" title="Questoes diretas neste topico">{direct}</span>' if direct > 0 else ''
        total_badge = f'<span class="badge total" title="Total de questoes (incluindo subtopicos)">{total}</span>' if total > 0 else '<span class="badge zero">0</span>'
        
        # Escapar nome para HTML
        nome_escaped = nome.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('"', '&quot;')
        
        # Determinar cor da profundidade
        border_color = depth_colors.get(depth, "var(--depth-multi)")
        
        if filhos:
            # Ordenar subtópicos alfabeticamente
            filhos_ordenados = sorted(filhos, key=lambda fid: topicos[fid]["nome"].lower())
            
            html_tree_parts.append(
                f'<details class="topic-node {has_questions}" data-depth="{depth}" data-name="{nome_escaped.lower()} {node_id}" data-id="{node_id}" style="border-left: 2px solid {border_color};">'
            )
            html_tree_parts.append(f'  <summary class="topic-summary">')
            html_tree_parts.append(f'    <span class="caret-icon">▸</span>')
            html_tree_parts.append(f'    <div class="topic-text-wrapper" style="flex: 1; display: flex; align-items: center; gap: 6px; overflow: hidden; min-width: 0;">')
            html_tree_parts.append(f'      <span class="topic-title">{nome_escaped}</span>')
            html_tree_parts.append(f'      <span class="topic-id" style="color: var(--text-muted); font-size: 11px; font-weight: normal; font-family: monospace; flex-shrink: 0;">({node_id})</span>')
            html_tree_parts.append(f'    </div>')
            html_tree_parts.append(f'    <span class="badges-wrapper">{direct_badge}{total_badge}</span>')
            html_tree_parts.append(f'  </summary>')
            html_tree_parts.append(f'  <div class="topic-content">')
            
            for f_id in filhos_ordenados:
                render_tree_to_html(f_id)
                
            html_tree_parts.append(f'  </div>')
            html_tree_parts.append(f'</details>')
        else:
            html_tree_parts.append(
                f'<div class="topic-leaf {has_questions}" data-depth="{depth}" data-name="{nome_escaped.lower()} {node_id}" data-id="{node_id}" style="border-left: 2px solid {border_color};">'
            )
            html_tree_parts.append(f'  <div class="topic-summary leaf-summary">')
            html_tree_parts.append(f'    <span class="bullet-icon">•</span>')
            html_tree_parts.append(f'    <div class="topic-text-wrapper" style="flex: 1; display: flex; align-items: center; gap: 6px; overflow: hidden; min-width: 0;">')
            html_tree_parts.append(f'      <span class="topic-title">{nome_escaped}</span>')
            html_tree_parts.append(f'      <span class="topic-id" style="color: var(--text-muted); font-size: 11px; font-weight: normal; font-family: monospace; flex-shrink: 0;">({node_id})</span>')
            html_tree_parts.append(f'    </div>')
            html_tree_parts.append(f'    <span class="badges-wrapper">{direct_badge}{total_badge}</span>')
            html_tree_parts.append(f'  </div>')
            html_tree_parts.append(f'</div>')
            
    # Renderizar todas as raízes e suas árvores
    for r_id in raizes_ordenadas:
        render_tree_to_html(r_id)
        
    html_tree_rendered = "\n".join(html_tree_parts)
    
    # 7. Construir o documento HTML completo
    safe_print("[LOG] Montando arquivo HTML completo com CSS premium e logica JS...")
    
    # Estatísticas rápidas para o painel superior
    total_active_questions_db = 0
    cur.execute("SELECT COUNT(*) as count FROM questaoresidencia")
    total_active_questions_db = cur.fetchone()["count"]
    
    cur.close()
    conn.close()
    
    # Carregar template HTML
    html_template = f"""<!DOCTYPE html>
<html lang="pt-BR" data-theme="dark">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Hierarquia de Tópicos - Banco de Questões</title>
    
    <!-- Google Fonts -->
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=Outfit:wght@400;500;600;700;800&display=swap" rel="stylesheet">
    
    <style>
        :root {{
            /* Cores - Tema Escuro (Default) */
            --bg-body: #0b0f19;
            --bg-card: #151d30;
            --bg-sidebar: #0f1524;
            --bg-input: #1f293d;
            --border-color: rgba(255, 255, 255, 0.08);
            
            --text-primary: #f3f4f6;
            --text-secondary: #9ca3af;
            --text-muted: #6b7280;
            
            --primary: #6366f1; /* Indigo */
            --primary-hover: #4f46e5;
            --primary-light: rgba(99, 102, 241, 0.15);
            
            --success: #10b981; /* Emerald */
            --success-light: rgba(16, 185, 129, 0.15);
            
            --accent: #f59e0b; /* Amber */
            
            /* Cores por Nível de Profundidade */
            --depth-1: #6366f1; /* Indigo */
            --depth-2: #10b981; /* Emerald */
            --depth-3: #f59e0b; /* Amber */
            --depth-4: #ec4899; /* Pink */
            --depth-5: #3b82f6; /* Blue */
            --depth-multi: #8b5cf6; /* Purple */
            
            --shadow-premium: 0 10px 25px -5px rgba(0, 0, 0, 0.3), 0 8px 10px -6px rgba(0, 0, 0, 0.3);
            --transition-smooth: all 0.25s cubic-bezier(0.4, 0, 0.2, 1);
        }}

        [data-theme="light"] {{
            /* Cores - Tema Claro */
            --bg-body: #f8fafc;
            --bg-card: #ffffff;
            --bg-sidebar: #f1f5f9;
            --bg-input: #e2e8f0;
            --border-color: rgba(0, 0, 0, 0.08);
            
            --text-primary: #1e293b;
            --text-secondary: #475569;
            --text-muted: #94a3b8;
            
            --primary: #4f46e5;
            --primary-hover: #4338ca;
            --primary-light: rgba(79, 70, 229, 0.1);
            
            --success: #059669;
            --success-light: rgba(5, 150, 105, 0.1);
            
            --depth-1: #4f46e5;
            --depth-2: #059669;
            --depth-3: #d97706;
            --depth-4: #db2777;
            --depth-5: #2563eb;
            --depth-multi: #7c3aed;
            
            --shadow-premium: 0 10px 25px -5px rgba(0, 0, 0, 0.05), 0 8px 10px -6px rgba(0, 0, 0, 0.05);
        }}

        * {{
            box-sizing: border-box;
            margin: 0;
            padding: 0;
        }}

        body {{
            font-family: 'Inter', -apple-system, sans-serif;
            background-color: var(--bg-body);
            color: var(--text-primary);
            display: flex;
            height: 100vh;
            overflow: hidden;
            transition: var(--transition-smooth);
        }}

        /* --- SIDEBAR --- */
        .sidebar {{
            width: 280px;
            background-color: var(--bg-sidebar);
            border-right: 1px solid var(--border-color);
            display: flex;
            flex-direction: column;
            flex-shrink: 0;
            transition: var(--transition-smooth);
            z-index: 10;
        }}

        .brand-section {{
            padding: 24px;
            border-bottom: 1px solid var(--border-color);
            display: flex;
            align-items: center;
            gap: 12px;
        }}

        .logo-placeholder {{
            width: 40px;
            height: 40px;
            background: linear-gradient(135deg, var(--primary), var(--success));
            border-radius: 10px;
            display: flex;
            align-items: center;
            justify-content: center;
            color: white;
            font-family: 'Outfit', sans-serif;
            font-weight: 800;
            font-size: 20px;
            box-shadow: 0 4px 10px rgba(99, 102, 241, 0.3);
        }}

        .brand-text h1 {{
            font-family: 'Outfit', sans-serif;
            font-size: 16px;
            font-weight: 700;
            color: var(--text-primary);
            line-height: 1.2;
        }}
        
        .brand-text span {{
            font-size: 11px;
            color: var(--text-muted);
            text-transform: uppercase;
            letter-spacing: 1px;
            font-weight: 600;
        }}

        .nav-section {{
            flex: 1;
            overflow-y: auto;
            padding: 20px 12px;
        }}

        .nav-section-title {{
            font-size: 11px;
            color: var(--text-muted);
            text-transform: uppercase;
            letter-spacing: 1px;
            padding-left: 12px;
            margin-bottom: 12px;
            font-weight: 700;
        }}

        .nav-item {{
            display: flex;
            align-items: center;
            justify-content: space-between;
            padding: 10px 12px;
            border-radius: 8px;
            color: var(--text-secondary);
            text-decoration: none;
            font-size: 14px;
            font-weight: 500;
            margin-bottom: 4px;
            cursor: pointer;
            transition: var(--transition-smooth);
        }}

        .nav-item:hover {{
            background-color: var(--border-color);
            color: var(--text-primary);
        }}

        .nav-item.active {{
            background-color: var(--primary-light);
            color: var(--primary);
            font-weight: 600;
        }}

        .nav-badge {{
            background-color: rgba(255, 255, 255, 0.08);
            color: var(--text-secondary);
            font-size: 11px;
            padding: 2px 6px;
            border-radius: 12px;
            font-weight: 600;
        }}

        [data-theme="light"] .nav-badge {{
            background-color: rgba(0, 0, 0, 0.05);
            color: var(--text-secondary);
        }}

        .nav-item.active .nav-badge {{
            background-color: var(--primary);
            color: white;
        }}

        .sidebar-footer {{
            padding: 20px;
            border-top: 1px solid var(--border-color);
            display: flex;
            justify-content: space-between;
            align-items: center;
        }}

        .theme-toggle {{
            background: none;
            border: 1px solid var(--border-color);
            color: var(--text-secondary);
            cursor: pointer;
            padding: 8px 12px;
            border-radius: 6px;
            font-size: 12px;
            font-weight: 600;
            display: flex;
            align-items: center;
            gap: 6px;
            transition: var(--transition-smooth);
        }}

        .theme-toggle:hover {{
            background-color: var(--border-color);
            color: var(--text-primary);
        }}

        /* --- MAIN CONTENT AREA --- */
        .main-container {{
            flex: 1;
            display: flex;
            flex-direction: column;
            height: 100vh;
            overflow: hidden;
        }}

        /* Sticky Header */
        .header {{
            background-color: rgba(21, 29, 48, 0.85);
            backdrop-filter: blur(12px);
            -webkit-backdrop-filter: blur(12px);
            border-bottom: 1px solid var(--border-color);
            padding: 20px 32px;
            display: flex;
            flex-direction: column;
            gap: 16px;
            z-index: 5;
            transition: var(--transition-smooth);
        }}

        [data-theme="light"] .header {{
            background-color: rgba(255, 255, 255, 0.85);
        }}

        .header-top {{
            display: flex;
            justify-content: space-between;
            align-items: center;
        }}

        .stats-dashboard {{
            display: flex;
            gap: 24px;
        }}

        .stat-card {{
            display: flex;
            flex-direction: column;
        }}

        .stat-val {{
            font-family: 'Outfit', sans-serif;
            font-size: 20px;
            font-weight: 700;
            color: var(--text-primary);
        }}

        .stat-label {{
            font-size: 11px;
            color: var(--text-muted);
            text-transform: uppercase;
            letter-spacing: 0.5px;
            font-weight: 600;
        }}

        /* Search and Controls Bar */
        .controls-row {{
            display: flex;
            align-items: center;
            gap: 16px;
            flex-wrap: wrap;
        }}

        .search-wrapper {{
            position: relative;
            flex: 1;
            min-width: 300px;
        }}

        .search-input {{
            width: 100%;
            background-color: var(--bg-input);
            border: 1px solid var(--border-color);
            border-radius: 8px;
            padding: 12px 16px 12px 42px;
            color: var(--text-primary);
            font-family: inherit;
            font-size: 14px;
            outline: none;
            transition: var(--transition-smooth);
        }}

        .search-input:focus {{
            border-color: var(--primary);
            box-shadow: 0 0 0 3px var(--primary-light);
        }}

        .search-icon {{
            position: absolute;
            left: 14px;
            top: 50%;
            transform: translateY(-50%);
            color: var(--text-muted);
            pointer-events: none;
            font-size: 16px;
        }}

        .clear-search-btn {{
            position: absolute;
            right: 14px;
            top: 50%;
            transform: translateY(-50%);
            background: none;
            border: none;
            color: var(--text-muted);
            cursor: pointer;
            font-size: 16px;
            display: none;
        }}

        .clear-search-btn:hover {{
            color: var(--text-primary);
        }}

        .action-btn {{
            background-color: var(--bg-card);
            border: 1px solid var(--border-color);
            color: var(--text-primary);
            cursor: pointer;
            padding: 11px 18px;
            border-radius: 8px;
            font-size: 13px;
            font-weight: 600;
            display: inline-flex;
            align-items: center;
            gap: 8px;
            transition: var(--transition-smooth);
        }}

        .action-btn:hover {{
            background-color: var(--border-color);
        }}

        .action-btn.active-filter {{
            background-color: var(--success-light);
            border-color: var(--success);
            color: var(--success);
        }}

        .action-btn.primary-btn {{
            background-color: var(--primary);
            border-color: var(--primary);
            color: white;
        }}

        .action-btn.primary-btn:hover {{
            background-color: var(--primary-hover);
        }}

        /* Tree View Container */
        .tree-container {{
            flex: 1;
            overflow-y: auto;
            padding: 32px;
            scroll-behavior: smooth;
        }}

        .tree-wrapper {{
            max-width: 900px;
            margin: 0 auto;
            background-color: var(--bg-card);
            border: 1px solid var(--border-color);
            border-radius: 12px;
            padding: 24px;
            box-shadow: var(--shadow-premium);
            transition: var(--transition-smooth);
        }}

        /* --- TREE ITEM STYLING (NESTED DETAILS) --- */
        .topic-node, .topic-leaf {{
            margin-top: 4px;
            margin-bottom: 4px;
            border-radius: 6px;
            transition: var(--transition-smooth);
            background-color: rgba(255, 255, 255, 0.01);
        }}

        [data-theme="light"] .topic-node, [data-theme="light"] .topic-leaf {{
            background-color: rgba(0, 0, 0, 0.005);
        }}

        .topic-summary {{
            list-style: none; /* Hide default summary caret */
            display: flex;
            align-items: center;
            padding: 8px 12px;
            cursor: pointer;
            border-radius: 6px;
            font-weight: 500;
            font-size: 14px;
            color: var(--text-primary);
            user-select: none;
            gap: 8px;
            transition: var(--transition-smooth);
        }}

        /* Remove default safari summary caret */
        .topic-summary::-webkit-details-marker {{
            display: none;
        }}

        .topic-summary:hover {{
            background-color: rgba(255, 255, 255, 0.04);
        }}

        [data-theme="light"] .topic-summary:hover {{
            background-color: rgba(0, 0, 0, 0.03);
        }}

        .topic-content {{
            padding-left: 20px;
            margin-top: 2px;
            margin-bottom: 6px;
        }}

        /* Custom caret icon rotation */
        .caret-icon {{
            display: inline-block;
            font-size: 10px;
            color: var(--text-muted);
            width: 14px;
            text-align: center;
            transition: transform 0.2s ease;
        }}

        .topic-node[open] > .topic-summary > .caret-icon {{
            transform: rotate(90deg);
        }}

        .bullet-icon {{
            color: var(--text-muted);
            width: 14px;
            text-align: center;
            font-size: 12px;
        }}

        .topic-title {{
            word-break: break-word;
        }}

        /* Badges styling */
        .badges-wrapper {{
            display: flex;
            align-items: center;
            gap: 6px;
            flex-shrink: 0;
        }}

        .badge {{
            font-size: 11px;
            font-weight: 600;
            padding: 2px 6px;
            border-radius: 4px;
            line-height: 1;
        }}

        .badge.direct {{
            background-color: rgba(59, 130, 246, 0.15);
            color: #3b82f6;
            border: 1px solid rgba(59, 130, 246, 0.25);
        }}

        [data-theme="light"] .badge.direct {{
            background-color: rgba(37, 99, 235, 0.08);
            color: #2563eb;
        }}

        .badge.total {{
            background-color: var(--success-light);
            color: var(--success);
            border: 1px solid rgba(16, 185, 129, 0.25);
        }}

        .badge.zero {{
            background-color: rgba(255, 255, 255, 0.03);
            color: var(--text-muted);
            border: 1px solid var(--border-color);
        }}

        [data-theme="light"] .badge.zero {{
            background-color: rgba(0, 0, 0, 0.02);
        }}

        /* Search Match Highlight */
        mark {{
            background-color: var(--accent);
            color: #0b0f19;
            padding: 0 2px;
            border-radius: 2px;
            font-weight: 700;
        }}

        /* Leaf styling (lowest level node with no children) */
        .topic-leaf {{
            background: none;
        }}

        .leaf-summary {{
            cursor: default;
        }}

        .leaf-summary:hover {{
            background-color: rgba(255, 255, 255, 0.02);
        }}

        [data-theme="light"] .leaf-summary:hover {{
            background-color: rgba(0, 0, 0, 0.015);
        }}

        /* Filtering states */
        .no-questions-filter-active .no-questions {{
            display: none !important;
        }}

        .searching .topic-node, .searching .topic-leaf {{
            display: none;
        }}

        .searching .match-found {{
            display: block !important;
        }}

        .searching .child-match-found {{
            display: block !important;
        }}

        /* Soft fade animation on search */
        @keyframes fadeIn {{
            from {{ opacity: 0; }}
            to {{ opacity: 1; }}
        }}

        .topic-node, .topic-leaf {{
            animation: fadeIn 0.15s ease-out;
        }}
    </style>
</head>
<body>

    <!-- SIDEBAR -->
    <div class="sidebar">
        <div class="brand-section">
            <div class="logo-placeholder">Q</div>
            <div class="brand-text">
                <h1>Tópicos</h1>
                <span>Banco de Questões</span>
            </div>
        </div>

        <div class="nav-section">
            <div class="nav-section-title">Grandes Áreas</div>
            <div class="nav-wrapper">
                <div class="nav-item active" onclick="scrollToTopic(null, this)">
                    <span>Todas as Áreas</span>
                    <span class="nav-badge">{len(raizes)}</span>
                </div>
                <div class="nav-item" onclick="scrollToTopic('33', this)">
                    <span>Cirurgia</span>
                    <span class="nav-badge" style="background-color: var(--depth-1); color: white;">{topicos[33]['total_count'] if 33 in topicos else 0}</span>
                </div>
                <div class="nav-item" onclick="scrollToTopic('100', this)">
                    <span>Clínica Médica</span>
                    <span class="nav-badge" style="background-color: var(--depth-2); color: white;">{topicos[100]['total_count'] if 100 in topicos else 0}</span>
                </div>
                <div class="nav-item" onclick="scrollToTopic('48', this)">
                    <span>Pediatria</span>
                    <span class="nav-badge" style="background-color: var(--depth-3); color: white;">{topicos[48]['total_count'] if 48 in topicos else 0}</span>
                </div>
                <div class="nav-item" onclick="scrollToTopic('183', this)">
                    <span>Ginecologia</span>
                    <span class="nav-badge" style="background-color: var(--depth-4); color: white;">{topicos[183]['total_count'] if 183 in topicos else 0}</span>
                </div>
                <div class="nav-item" onclick="scrollToTopic('218', this)">
                    <span>Obstetrícia</span>
                    <span class="nav-badge" style="background-color: var(--depth-5); color: white;">{topicos[218]['total_count'] if 218 in topicos else 0}</span>
                </div>
                <div class="nav-item" onclick="scrollToTopic('29', this)">
                    <span>Medicina Preventiva</span>
                    <span class="nav-badge" style="background-color: var(--depth-multi); color: white;">{topicos[29]['total_count'] if 29 in topicos else 0}</span>
                </div>
                <div class="nav-item" onclick="scrollToTopic('67', this)">
                    <span>Outros</span>
                    <span class="nav-badge">{topicos[67]['total_count'] if 67 in topicos else 0}</span>
                </div>
            </div>
        </div>

        <div class="sidebar-footer">
            <button class="theme-toggle" onclick="toggleTheme()">
                <span id="theme-icon">☀</span> <span id="theme-text">Tema Claro</span>
            </button>
        </div>
    </div>

    <!-- MAIN CONTAINER -->
    <div class="main-container">
        <!-- HEADER -->
        <div class="header">
            <div class="header-top">
                <div class="header-titles">
                    <h2 style="font-family: 'Outfit', sans-serif; font-size: 22px; font-weight: 700;">Hierarquia do Banco de Questões</h2>
                    <p style="font-size: 13px; color: var(--text-muted);">Navegue, pesquise e analise a distribuição de tópicos do banco de dados em tempo real</p>
                </div>
                <div class="stats-dashboard">
                    <div class="stat-card">
                        <span class="stat-val">{len(topicos):,}</span>
                        <span class="stat-label">Total de Tópicos</span>
                    </div>
                    <div class="stat-card" style="border-left: 1px solid var(--border-color); padding-left: 20px;">
                        <span class="stat-val">{total_active_questions_db:,}</span>
                        <span class="stat-label">Questões Ativas</span>
                    </div>
                    <div class="stat-card" style="border-left: 1px solid var(--border-color); padding-left: 20px;">
                        <span class="stat-val">{max_depth}</span>
                        <span class="stat-label">Nível Máximo</span>
                    </div>
                </div>
            </div>

            <div class="controls-row">
                <div class="search-wrapper">
                    <span class="search-icon">🔍</span>
                    <input type="text" class="search-input" id="search-box" placeholder="Pesquisar por assunto ou tópico (Ex: Sutura)..." oninput="handleSearch()">
                    <button class="clear-search-btn" id="clear-search-btn" onclick="clearSearch()">✕</button>
                </div>

                <button class="action-btn" id="filter-questions-btn" onclick="toggleFilterQuestions()">
                    <span>🎯</span> Apenas com Questões
                </button>

                <button class="action-btn" onclick="expandAll()">
                    <span>📂</span> Expandir Tudo
                </button>

                <button class="action-btn" onclick="collapseAll()">
                    <span>📁</span> Contrair Tudo
                </button>
            </div>
        </div>

        <!-- TREE VIEW -->
        <div class="tree-container">
            <div class="tree-wrapper" id="tree-root">
                {html_tree_rendered}
            </div>
        </div>
    </div>

    <!-- JAVASCRIPT LOGIC -->
    <script>
        // Alternar entre Tema Claro e Tema Escuro
        function toggleTheme() {{
            const html = document.documentElement;
            const currentTheme = html.getAttribute('data-theme');
            const newTheme = currentTheme === 'dark' ? 'light' : 'dark';
            html.setAttribute('data-theme', newTheme);
            
            const icon = document.getElementById('theme-icon');
            const text = document.getElementById('theme-text');
            
            if (newTheme === 'dark') {{
                icon.textContent = '☀';
                text.textContent = 'Tema Claro';
            }} else {{
                icon.textContent = '🌙';
                text.textContent = 'Tema Escuro';
            }}
        }}

        // Rolar até um tópico principal e ativá-lo na barra lateral
        function scrollToTopic(topicId, element) {{
            // Remover classe active de todos
            document.querySelectorAll('.nav-item').forEach(item => item.classList.remove('active'));
            element.classList.add('active');

            if (topicId === null) {{
                document.querySelector('.tree-container').scrollTop = 0;
                return;
            }}

            const target = document.querySelector(`[data-id="${{topicId}}"]`);
            if (target) {{
                // Se for um detalhes colapsado, garantir que está aberto
                if (target.tagName === 'DETAILS') {{
                    target.open = true;
                }}
                target.scrollIntoView({{ behavior: 'smooth', block: 'start' }});
                
                // Dar um breve efeito visual de destaque
                const summary = target.querySelector('.topic-summary');
                if (summary) {{
                    const origBg = summary.style.backgroundColor;
                    summary.style.backgroundColor = 'var(--primary-light)';
                    setTimeout(() => {{
                        summary.style.backgroundColor = origBg;
                    }}, 1500);
                }}
            }}
        }}

        // Expandir todos os detalhes
        function expandAll() {{
            document.querySelectorAll('.topic-node').forEach(detail => {{
                detail.open = true;
            }});
        }}

        // Contrair todos os detalhes
        function collapseAll() {{
            document.querySelectorAll('.topic-node').forEach(detail => {{
                detail.open = false;
            }});
        }}

        // Filtrar apenas tópicos com questões
        function toggleFilterQuestions() {{
            const btn = document.getElementById('filter-questions-btn');
            const root = document.getElementById('tree-root');
            
            if (btn.classList.contains('active-filter')) {{
                btn.classList.remove('active-filter');
                root.classList.remove('no-questions-filter-active');
            }} else {{
                btn.classList.add('active-filter');
                root.classList.add('no-questions-filter-active');
            }}
        }}

        // Normalização em português (remover acentos e cedilhas)
        function normalizeText(text) {{
            return text.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();
        }}

        // Lógica de pesquisa em tempo real
        let searchTimeout = null;

        function handleSearch() {{
            clearTimeout(searchTimeout);
            
            // Usar debouncing para digitação rápida e responsiva
            searchTimeout = setTimeout(() => {{
                const searchBox = document.getElementById('search-box');
                const clearBtn = document.getElementById('clear-search-btn');
                const root = document.getElementById('tree-root');
                const query = normalizeText(searchBox.value.trim());

                if (query.length > 0) {{
                    clearBtn.style.display = 'block';
                    root.classList.add('searching');
                    performSearch(query);
                }} else {{
                    clearBtn.style.display = 'none';
                    root.classList.remove('searching');
                    clearHighlights();
                }}
            }}, 150);
        }}

        function clearSearch() {{
            const searchBox = document.getElementById('search-box');
            searchBox.value = '';
            handleSearch();
            searchBox.focus();
        }}

        function performSearch(query) {{
            // 1. Limpar estados anteriores de match
            document.querySelectorAll('.match-found, .child-match-found').forEach(el => {{
                el.classList.remove('match-found', 'child-match-found');
            }});
            clearHighlights();

            // 2. Localizar nós que correspondem diretamente à busca
            const allElements = document.querySelectorAll('.topic-node, .topic-leaf');
            
            allElements.forEach(el => {{
                const name = el.getAttribute('data-name');
                if (name.includes(query)) {{
                    el.classList.add('match-found');
                    
                    // Destacar termo de busca no título
                    highlightTitle(el, query);
                    
                    // 3. Subir na hierarquia de pais para marcar como 'child-match-found' e abrir os details
                    let parent = el.parentElement.closest('.topic-node');
                    while (parent) {{
                        parent.classList.add('child-match-found');
                        parent.open = true; // Abrir o ramo para expor a busca
                        parent = parent.parentElement.closest('.topic-node');
                    }}
                }}
            }});
        }}

        function highlightTitle(element, query) {{
            const titleEl = element.querySelector('.topic-title');
            if (!titleEl) return;

            const text = titleEl.textContent;
            const normalized = normalizeText(text);
            
            // Encontrar os índices do termo buscado na string normalizada
            let index = normalized.indexOf(query);
            if (index === -1) return;

            // Guardar texto original com marcação
            let html = '';
            let lastIndex = 0;
            
            while (index !== -1) {{
                html += text.substring(lastIndex, index);
                html += `<mark>${{text.substring(index, index + query.length)}}</mark>`;
                lastIndex = index + query.length;
                index = normalized.indexOf(query, lastIndex);
            }}
            html += text.substring(lastIndex);
            
            titleEl.innerHTML = html;
        }}

        function clearHighlights() {{
            document.querySelectorAll('mark').forEach(mark => {{
                const parent = mark.parentNode;
                parent.replaceChild(document.createTextNode(mark.textContent), mark);
                parent.normalize(); // Juntar nós de texto adjacentes
            }});
        }}
    </script>
</body>
</html>
"""
    
    # Gravar o arquivo HTML
    output_filename = "hierarquia_topicos_interativo.html"
    output_path = os.path.join(os.getcwd(), output_filename)
    
    safe_print(f"[LOG] Salvando o arquivo gerado em {output_path}...")
    try:
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(html_template)
        safe_print("[LOG] SUCESSO! O arquivo interativo de topicos foi exportado com sucesso. [OK]")
        safe_print(f"[LOG] Caminho do arquivo: {output_path}")
        safe_print(f"[LOG] Tempo de execucao: {time.time() - start_time:.2f} segundos.")
        safe_print("\nProntinho! Voce ja pode abrir o arquivo 'hierarquia_topicos_interativo.html' em qualquer navegador para interagir com a arvore.")
    except Exception as e:
        safe_print(f"[ERRO] Falha ao salvar o arquivo HTML: {e}")

if __name__ == "__main__":
    main()
