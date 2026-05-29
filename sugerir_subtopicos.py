import os
import sys
import mysql.connector
import json
import requests
import time
import re
from bs4 import BeautifulSoup

def safe_print(msg):
    try:
        print(msg)
    except UnicodeEncodeError:
        try:
            print(msg.encode('ascii', errors='replace').decode('ascii'))
        except Exception:
            pass

# 1. Carregar senhas e chaves
def load_db_password():
    password_path = os.path.join(os.path.dirname(__file__), 'db_password.txt')
    try:
        with open(password_path, 'r', encoding='utf-8') as f:
            pwd = f.read().strip()
            if not pwd:
                raise ValueError("Senha esta vazia no arquivo db_password.txt")
            return pwd
    except FileNotFoundError:
        safe_print(f"[ERRO] Arquivo db_password.txt nao encontrado em {password_path}")
        safe_print("[ERRO] Crie o arquivo db_password.txt na raiz do projeto com a senha do banco")
        sys.exit(1)
    except Exception as e:
        safe_print(f"[ERRO] Erro ao ler db_password.txt: {e}")
        sys.exit(1)

def load_api_key():
    api_key_path = os.path.join(os.path.dirname(__file__), 'api_key_deepseek.txt')
    try:
        with open(api_key_path, 'r', encoding='utf-8') as f:
            api_key = f.read().strip()
            if not api_key:
                raise ValueError("API key esta vazia no arquivo api_key.txt")
            return api_key
    except FileNotFoundError:
        safe_print(f"[ERRO] Arquivo api_key.txt nao encontrado em {api_key_path}")
        safe_print("[ERRO] Crie o arquivo api_key.txt na raiz do projeto com sua API key do DeepSeek")
        sys.exit(1)
    except Exception as e:
        safe_print(f"[ERRO] Erro ao ler api_key.txt: {e}")
        sys.exit(1)

# Configurações globais
DB_CONFIG = {
    "host": "localhost",
    "user": "root",
    "password": load_db_password(),
    "database": "qconcursos"
}

DEEPSEEK_CONFIG = {
    "api_key": load_api_key(),
    "model": "deepseek-chat",
    "temperature": 0.1,
    "url": "https://api.deepseek.com/v1/chat/completions"
}

def deepseek_chat(messages, max_tokens=1500, temperature=None):
    if temperature is None:
        temperature = DEEPSEEK_CONFIG["temperature"]

    headers = {
        "Authorization": f"Bearer {DEEPSEEK_CONFIG['api_key']}",
        "Content-Type": "application/json"
    }

    payload = {
        "model": DEEPSEEK_CONFIG["model"],
        "messages": messages,
        "temperature": temperature,
        "max_tokens": max_tokens
    }

    try:
        response = requests.post(DEEPSEEK_CONFIG["url"], headers=headers, json=payload, timeout=60)
        response.raise_for_status()
        data = response.json()
        content = data['choices'][0]['message']['content']
        return content
    except Exception as e:
        safe_print(f"[ERRO] Falha ao chamar a API do DeepSeek: {e}")
        return None

def limpar_html(html_text):
    if not html_text:
        return ""
    try:
        soup = BeautifulSoup(html_text, 'html.parser')
        for s in soup(['script', 'style']):
            s.decompose()
        return soup.get_text(separator=' ').strip()
    except Exception:
        return re.sub(r'<[^>]+>', ' ', html_text).strip()

def extract_json(text):
    if not text:
        return None
    # Procurar blocos de markdown ```json ... ```
    match = re.search(r'```json\s*(.*?)\s*```', text, re.DOTALL | re.IGNORECASE)
    if not match:
        match = re.search(r'```\s*(.*?)\s*```', text, re.DOTALL | re.IGNORECASE)
    if match:
        text_clean = match.group(1)
    else:
        text_clean = text
    
    try:
        return json.loads(text_clean.strip())
    except json.JSONDecodeError:
        # Fallback: tentar encontrar o primeiro '[' e o último ']'
        try:
            start = text_clean.find('[')
            end = text_clean.rfind(']')
            if start != -1 and end != -1:
                return json.loads(text_clean[start:end+1].strip())
        except Exception:
            pass
        return None

# Funções auxiliares de Banco
def get_caminho_ancestrais(cur, topico_id):
    caminho = []
    curr_id = topico_id
    while curr_id:
        cur.execute("SELECT id, nome, id_pai FROM topico WHERE id = %s", (curr_id,))
        row = cur.fetchone()
        if not row:
            break
        caminho.insert(0, row['nome'])
        curr_id = row['id_pai']
    return " > ".join(caminho)

def buscar_topico_por_nome_ou_id(cur, termo):
    # Se for inteiro, busca por ID diretamente
    if termo.isdigit():
        cur.execute("SELECT id, nome, id_pai FROM topico WHERE id = %s", (int(termo),))
        row = cur.fetchone()
        if row:
            return [row]
    
    # Caso contrário, busca por nome (LIKE)
    cur.execute("SELECT id, nome, id_pai FROM topico WHERE nome LIKE %s LIMIT 30", (f"%{termo}%",))
    return cur.fetchall()

def sugerir_e_criar_subtopicos(cur, conn, parent_id, parent_nome, hierarquia_pai):
    # 2. Coletar contexto de questões sob este tópico
    safe_print("\n[LOG] Buscando questoes classificadas neste topico para contexto da IA...")
    cur.execute("""
        SELECT q.questao_id, q.enunciado, q.alternativaA, q.alternativaB, q.alternativaC, q.alternativaD, q.alternativaE, q.gabarito
        FROM classificacao_questao cq
        INNER JOIN questaoresidencia q ON cq.id_questao = q.questao_id
        WHERE cq.id_topico = %s
        LIMIT 6
    """, (parent_id,))
    questoes = cur.fetchall()
    
    contexto_questoes = ""
    if questoes:
        safe_print(f"[LOG] Encontradas {len(questoes)} questoes para contexto de amostragem.")
        for idx, q in enumerate(questoes, 1):
            enunciado_limpo = limpar_html(q['enunciado'])[:400]
            contexto_questoes += f"Questao {idx} (ID {q['questao_id']}):\nEnunciado: {enunciado_limpo}...\n\n"
    else:
        safe_print("[LOG] Nenhuma questao diretamente associada a este topico no banco.")
        contexto_questoes = "Nenhuma questao cadastrada no momento."

    # 3. Geração de Sugestões via DeepSeek
    suggestions = []
    prompt_custom_instruction = ""
    
    while True:
        safe_print("\n[LOG] Solicitando sugestoes de subtopicos a API do DeepSeek...")
        
        prompt = f"""
Você é um médico especialista e organizador de currículos de medicina e provas de residência médica.
O tópico pai é: "{parent_nome}" (Caminho completo: {hierarquia_pai}).
Queremos criar subtópicos (filhos) específicos para esse assunto para que possamos classificar melhor as questões no banco.

Abaixo estão alguns exemplos de enunciados de questões reais classificadas sob este tópico para você entender o escopo real das perguntas:
---
{contexto_questoes}
---

{prompt_custom_instruction}

Com base nisso, sugira entre 2 e 10 subtópicos específicos, clinicamente corretos, mutuamente exclusivos e abrangentes para dividir este assunto.
Evite tópicos excessivamente longos ou genéricos (como "Outros", "Introdução", "Geral", "Miscelânea").

Responda APENAS com um array JSON de objetos contendo os campos "nome" e "descricao" (sem tags markdown de texto extras, fora do bloco de código):
```json
[
  {{"nome": "Nome do Subtópico 1", "descricao": "Breve descrição do escopo clínico"}},
  {{"nome": "Nome do Subtópico 2", "descricao": "Breve descrição do escopo clínico"}}
]
```
"""
        messages = [{"role": "user", "content": prompt}]
        resposta = deepseek_chat(messages)
        
        if not resposta:
            safe_print("[ERRO] Nao foi possivel obter sugestoes da IA. Tentando novamente...")
            continue
            
        suggestions = extract_json(resposta)
        if not suggestions or not isinstance(suggestions, list):
            safe_print("[ERRO] A IA nao retornou um JSON valido no formato correto. Resposta bruta:")
            safe_print(resposta)
            input("\nPressione Enter para tentar novamente...")
            continue
            
        # 4. Loop de Revisão CLI
        while True:
            safe_print(f"\n=========================================================")
            safe_print(f"=== SUBTOPICOS SUGERIDOS PARA: {parent_nome} ===")
            safe_print(f"=========================================================")
            for idx, s in enumerate(suggestions, 1):
                safe_print(f"  {idx}. {s['nome']} - {s.get('descricao', '')}")
            
            safe_print("\nOpcoes:")
            safe_print("  [1] Aceitar e criar TODOS os subtopicos sugeridos")
            safe_print("  [2] Selecionar indices especificos para criar (ex: 1, 2, 4)")
            safe_print("  [3] Adicionar um subtopico customizado manualmente")
            safe_print("  [4] Regenerar sugestoes com nova instrucao para IA")
            safe_print("  [5] Cancelar e sair")
            
            opcao = input("\nEscolha uma opcao (1-5): ").strip()
            
            if opcao == '1':
                subtopicos_aprovados = [s['nome'] for s in suggestions]
                break
            elif opcao == '2':
                indices_str = input("Digite os numeros dos subtopicos desejados separados por virgula (ex: 1,3,5): ").strip()
                indices = [int(i.strip()) for i in indices_str.split(',') if i.strip().isdigit()]
                subtopicos_aprovados = [suggestions[i - 1]['nome'] for i in indices if 1 <= i <= len(suggestions)]
                if not subtopicos_aprovados:
                    safe_print("[AVISO] Nenhum subtopico valido selecionado.")
                    continue
                break
            elif opcao == '3':
                novo_nome = input("Digite o nome do novo subtopico customizado: ").strip()
                if novo_nome:
                    suggestions.append({"nome": novo_nome, "descricao": "Adicionado manualmente pelo usuario"})
                    safe_print(f"[LOG] Subtopico '{novo_nome}' adicionado a lista.")
                continue
            elif opcao == '4':
                instrucao = input("Digite a instrucao para a IA (ex: 'Focar apenas em patologias infecciosas'): ").strip()
                if instrucao:
                    prompt_custom_instruction = f"Instrução adicional do usuário para focar a divisão:\n{instrucao}\n"
                else:
                    prompt_custom_instruction = ""
                break # Sai da revisão interna e volta para a regeneração
            elif opcao == '5':
                cur.close()
                conn.close()
                safe_print("Operacao cancelada pelo usuario.")
                sys.exit(0)
            else:
                safe_print("[AVISO] Opcao invalida.")
                continue
        
        # Se escolheu a opção 4, o loop de regeneração continua, caso contrário quebramos o loop principal
        if opcao != '4':
            break

    # 5. Gravação no Banco de Dados
    safe_print(f"\n[LOG] Preparando para criar {len(subtopicos_aprovados)} subtopicos sob '{parent_nome}'...")
    novos_topicos_ids = []
    
    try:
        for nome_sub in subtopicos_aprovados:
            # Verificar se já existe com esse nome sob este pai
            cur.execute("SELECT id FROM topico WHERE nome = %s AND id_pai = %s", (nome_sub, parent_id))
            row = cur.fetchone()
            if row:
                safe_print(f"  - Subtopico '{nome_sub}' ja existe (ID: {row['id']}). Ignorando.")
                novos_topicos_ids.append((row['id'], nome_sub))
            else:
                cur.execute("INSERT INTO topico (nome, id_pai) VALUES (%s, %s)", (nome_sub, parent_id))
                new_id = cur.lastrowid
                safe_print(f"  + Subtopico '{nome_sub}' criado com sucesso! (ID: {new_id})")
                novos_topicos_ids.append((new_id, nome_sub))
        
        conn.commit()
        safe_print("[LOG] Gravacao dos subtopicos concluida e commitada no banco de dados. [OK]")
        return novos_topicos_ids
    except Exception as e:
        conn.rollback()
        safe_print(f"[ERRO] Falha ao inserir subtopicos no banco de dados: {e}")
        cur.close()
        conn.close()
        sys.exit(1)

def main():
    safe_print("=========================================================")
    safe_print("=== GERADOR E SUGERIDOR DE SUBTOPICOS VIA DEEPSEEK ===")
    safe_print("=========================================================")

    # Conectar ao banco
    try:
        conn = mysql.connector.connect(**DB_CONFIG)
        cur = conn.cursor(dictionary=True)
        safe_print("[LOG] Conexao com banco de dados estabelecida. [OK]")
    except Exception as e:
        safe_print(f"[ERRO] Falha ao conectar ao banco de dados: {e}")
        sys.exit(1)

    # 1. Selecionar o tópico pai
    topico_selecionado = None
    while not topico_selecionado:
        termo = input("\nDigite o ID ou parte do NOME do topico pai (ex: Neuropediatria ou 2354): ").strip()
        if not termo:
            continue
        
        resultados = buscar_topico_por_nome_ou_id(cur, termo)
        if not resultados:
            safe_print(f"[AVISO] Nenhum topico encontrado para '{termo}'. Tente novamente.")
            continue
        
        if len(resultados) == 1:
            topico_selecionado = resultados[0]
        else:
            safe_print(f"\nForam encontrados {len(resultados)} topicos. Escolha o ID correto:")
            for idx, r in enumerate(resultados, 1):
                # Pegar caminho do pai
                caminho = get_caminho_ancestrais(cur, r['id'])
                safe_print(f"  [{idx}] ID: {r['id']} | Hierarquia: {caminho}")
            
            escolha = input("\nSelecione o numero da opcao correta (ou pressione Enter para pesquisar de novo): ").strip()
            if escolha.isdigit() and 1 <= int(escolha) <= len(resultados):
                topico_selecionado = resultados[int(escolha) - 1]
            else:
                continue

    parent_id = topico_selecionado['id']
    parent_nome = topico_selecionado['nome']
    hierarquia_pai = get_caminho_ancestrais(cur, parent_id)
    
    safe_print(f"\n[LOG] Topico Pai Selecionado: {parent_nome} (ID: {parent_id})")
    safe_print(f"[LOG] Hierarquia Completa: {hierarquia_pai}")

    # Verificar se já possui filhos
    cur.execute("SELECT id, nome FROM topico WHERE id_pai = %s", (parent_id,))
    filhos_existentes = cur.fetchall()
    
    modo = None
    if filhos_existentes:
        safe_print(f"\n[LOG] Este topico ja possui {len(filhos_existentes)} subtopicos filhos cadastrados:")
        for idx, f in enumerate(filhos_existentes[:15], 1):
            safe_print(f"  - [{f['id']}] {f['nome']}")
        if len(filhos_existentes) > 15:
            safe_print(f"  ... e mais {len(filhos_existentes) - 15} filhos.")
            
        safe_print("\nEscolha a acao desejada:")
        safe_print("  [1] Pipeline Completo (Sugerir novos subtopicos com IA + Classificar questoes)")
        safe_print("  [2] Apenas Classificacao (Classificar questoes nos subtopicos ja existentes)")
        safe_print("  [3] Cancelar e sair")
        
        while True:
            escolha_modo = input("\nEscolha uma opcao (1-3): ").strip()
            if escolha_modo == '1':
                modo = 'completo'
                break
            elif escolha_modo == '2':
                modo = 'apenas_classificacao'
                break
            elif escolha_modo == '3':
                cur.close()
                conn.close()
                safe_print("Operacao cancelada pelo usuario.")
                sys.exit(0)
            else:
                safe_print("[AVISO] Opcao invalida.")
    else:
        safe_print("\n[LOG] Este topico nao possui subtopicos filhos cadastrados.")
        safe_print("Iniciando o Pipeline Completo para sugerir novos subtopicos...")
        modo = 'completo'

    if modo == 'completo':
        novos_topicos_ids = sugerir_e_criar_subtopicos(cur, conn, parent_id, parent_nome, hierarquia_pai)
    else:
        # Apenas classificação nos subtopicos já existentes
        novos_topicos_ids = [(f['id'], f['nome']) for f in filhos_existentes]
        safe_print(f"\n[LOG] Usando {len(novos_topicos_ids)} subtopicos ja existentes para a classificacao.")

    # 6. Reclassificação Automática de Questões Existentes
    # Contar total de questões diretamente associadas ao pai
    cur.execute("""
        SELECT COUNT(DISTINCT cq.id_questao) as count
        FROM classificacao_questao cq
        INNER JOIN questaoresidencia q ON cq.id_questao = q.questao_id
        WHERE cq.id_topico = %s
    """, (parent_id,))
    total_questoes_pai = cur.fetchone()['count']
    
    if total_questoes_pai > 0:
        safe_print(f"\n=========================================================")
        safe_print(f"=== CLASSIFICACAO DE QUESTOES ===")
        safe_print(f"=========================================================")
        safe_print(f"Existem atualmente {total_questoes_pai} questoes associadas diretamente ao topico pai '{parent_nome}'.")
        
        if modo == 'apenas_classificacao':
            msg_pergunta = "Deseja classificar essas questoes nos subtopicos existentes? (s/n, padrao: s): "
        else:
            msg_pergunta = "Deseja reclassificar essas questoes automaticamente nos novos subtopicos criados? (s/n, padrao: s): "
            
        reclassificar = input(msg_pergunta).strip().lower()
        
        if reclassificar != 'n':
            safe_print("\n[LOG] Buscando dados das questoes a classificar...")
            cur.execute("""
                SELECT q.questao_id, q.codigo, q.enunciado, q.alternativaA, q.alternativaB, q.alternativaC, q.alternativaD, q.alternativaE, q.gabarito
                FROM classificacao_questao cq
                INNER JOIN questaoresidencia q ON cq.id_questao = q.questao_id
                WHERE cq.id_topico = %s
            """, (parent_id,))
            questoes_para_reclassificar = cur.fetchall()
            
            # Montar a lista de opções de reclassificação
            subtopicos_opcoes = "\n".join([f"{idx}. {nome} (ID: {tid})" for idx, (tid, nome) in enumerate(novos_topicos_ids, 1)])
            num_opcoes = len(novos_topicos_ids)
            
            safe_print(f"[LOG] Iniciando classificacao de {len(questoes_para_reclassificar)} questoes...")
            reclassificadas = 0
            
            for idx_q, q in enumerate(questoes_para_reclassificar, 1):
                qid = q['questao_id']
                enunciado_limpo = limpar_html(q['enunciado'])[:3500]
                
                # Montar alternativas formatadas
                alts = ""
                for alt_letra in ['A', 'B', 'C', 'D', 'E']:
                    alt_val = q.get(f'alternativa{alt_letra}')
                    if alt_val:
                        alts += f"  {alt_letra}) {limpar_html(alt_val)}\n"
                
                prompt_reclass = f"""
Você é um classificador estruturado de questões de medicina e residência médica.
A questão abaixo pertencia ao tópico abrangente "{parent_nome}". O tópico possui os seguintes subtópicos específicos:
{subtopicos_opcoes}

Analise a questão e determine a qual subtópico ela pertence mais especificamente.
Se a questão for muito ampla, genérica ou se encaixar em múltiplos subtópicos sem um vencedor claro, responda com 0 para mantê-la no tópico pai.

[ENUNCIADO]: {enunciado_limpo}
[ALTERNATIVAS]:
{alts}
[GABARITO CORRETO]: {q['gabarito']}

Responda APENAS com o número correspondente ao subtópico na lista acima (1 a {num_opcoes}), ou 0 para manter no pai. Não inclua nenhuma outra palavra ou explicação.
"""
                try:
                    res_reclass = deepseek_chat([{"role": "user", "content": prompt_reclass}], max_tokens=10)
                    if res_reclass:
                        # Extrair o primeiro número
                        num_match = re.search(r'\d+', res_reclass)
                        if num_match:
                            escolha_idx = int(num_match.group())
                            if 1 <= escolha_idx <= num_opcoes:
                                novo_topico_id, novo_topico_nome = novos_topicos_ids[escolha_idx - 1]
                                
                                # Fazer UPDATE na tabela classificacao_questao para esta questão
                                # Troca a classificação de cq.id_topico = parent_id para cq.id_topico = novo_topico_id
                                cur.execute("""
                                    UPDATE classificacao_questao
                                    SET id_topico = %s
                                    WHERE id_questao = %s AND id_topico = %s
                                """, (novo_topico_id, qid, parent_id))
                                
                                safe_print(f"  [{idx_q}/{len(questoes_para_reclassificar)}] Questao {qid} reclassificada: '{parent_nome}' -> '{novo_topico_nome}'")
                                reclassificadas += 1
                                continue
                                
                    # Se respondeu 0 ou falhou
                    safe_print(f"  [{idx_q}/{len(questoes_para_reclassificar)}] Questao {qid} mantida no topico pai '{parent_nome}'")
                except Exception as ex_q:
                    safe_print(f"  [AVISO] Falha ao reclassificar questao {qid}: {ex_q}")
            
            # Salvar reclassificações
            conn.commit()
            safe_print(f"\n[LOG] Classificacao concluida! {reclassificadas} de {len(questoes_para_reclassificar)} questoes foram classificadas com sucesso. [OK]")
    else:
        safe_print(f"\n[LOG] Nenhuma questao diretamente associada ao topico pai '{parent_nome}' encontrada para classificar.")
    
    cur.close()
    conn.close()
    safe_print("\n=========================================================")
    safe_print("=== OPERACAO FINALIZADA COM SUCESSO! ===")
    safe_print("=========================================================")
    safe_print("Todos os subtópicos foram verificados e as questões organizadas!")

if __name__ == "__main__":
    main()
