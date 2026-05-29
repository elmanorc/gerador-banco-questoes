import os
import sys
import mysql.connector
import time

def safe_print(msg):
    try:
        print(msg)
    except UnicodeEncodeError:
        try:
            print(msg.encode('ascii', errors='replace').decode('ascii'))
        except Exception:
            pass

# 1. Carregar senha do banco
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
        sys.exit(1)
    except Exception as e:
        safe_print(f"[ERRO] Erro ao ler db_password.txt: {e}")
        sys.exit(1)

# Configurações globais
DB_CONFIG = {
    "host": "localhost",
    "user": "root",
    "password": load_db_password(),
    "database": "qconcursos"
}

def main():
    safe_print("=========================================================")
    safe_print("=== VERIFICADOR DE REDUNDANCIA HIERARQUICA DE TOPICOS ===")
    safe_print("=========================================================")
    
    # Verificar se foi passado argumento --fix
    execute_fix = "--fix" in sys.argv
    
    # 1. Conectar ao banco
    try:
        conn = mysql.connector.connect(**DB_CONFIG)
        cur = conn.cursor(dictionary=True)
        safe_print("[LOG] Conexao com banco de dados estabelecida. [OK]")
    except Exception as e:
        safe_print(f"[ERRO] Falha ao conectar ao banco de dados: {e}")
        sys.exit(1)
        
    start_time = time.time()
    
    # 2. Carregar todos os tópicos
    safe_print("[LOG] Carregando topicos para mapear ancestrais...")
    cur.execute("SELECT id, id_pai, nome FROM topico")
    topicos_rows = cur.fetchall()
    
    topicos = {r["id"]: {"pai": r["id_pai"], "nome": r["nome"]} for r in topicos_rows}
    safe_print(f"[LOG] {len(topicos)} topicos carregados.")
    
    # 3. Mapear ancestrais de cada tópico (com cache)
    ancestrais_cache = {}
    
    def get_ancestrais(tid):
        if tid in ancestrais_cache:
            return ancestrais_cache[tid]
            
        ancestrais = set()
        curr_id = topicos[tid]["pai"]
        visitados = {tid} # Detecção de ciclos
        
        while curr_id is not None and curr_id in topicos:
            if curr_id in visitados:
                # Ciclo detectado na hierarquia!
                break
            visitados.add(curr_id)
            ancestrais.add(curr_id)
            curr_id = topicos[curr_id]["pai"]
            
        ancestrais_cache[tid] = ancestrais
        return ancestrais
        
    # Pre-popular cache de ancestrais
    for tid in topicos:
        get_ancestrais(tid)
        
    # 4. Carregar todas as classificações de questões
    safe_print("[LOG] Carregando associacoes questao-topico da tabela 'classificacao_questao'...")
    cur.execute("SELECT id_questao, id_topico FROM classificacao_questao")
    classifs_rows = cur.fetchall()
    safe_print(f"[LOG] {len(classifs_rows)} registros de classificacao carregados.")
    
    # Agrupar classificações por ID de questão
    classifs = {}
    for r in classifs_rows:
        qid = r["id_questao"]
        tid = r["id_topico"]
        classifs.setdefault(qid, set()).add(tid)
        
    # 5. Analisar inconsistências de redundância
    safe_print("[LOG] Analisando consistencia da redundancia (transitive closure)...")
    
    total_questoes_analisadas = len(classifs)
    questoes_anomalas = 0
    total_redundancias_faltantes = 0
    insercoes_necessarias = [] # lista de tuplas (id_questao, id_topico_faltante)
    
    for qid, tids_associados in classifs.items():
        # Calcular quais ancestrais deveriam estar associados
        ancestrais_necessarios = set()
        for tid in tids_associados:
            if tid in topicos:
                ancestrais_necessarios.update(get_ancestrais(tid))
                
        # Identificar quais ancestrais necessários estão ausentes nas associações atuais
        ancestrais_faltantes = ancestrais_necessarios - tids_associados
        
        if ancestrais_faltantes:
            questoes_anomalas += 1
            total_redundancias_faltantes += len(ancestrais_faltantes)
            for tid_faltante in ancestrais_faltantes:
                insercoes_necessarias.append((qid, tid_faltante))
                
    safe_print(f"\n🎯 RESULTADOS DA VERIFICACAO:")
    safe_print(f"  • Total de Questoes Analisadas: {total_questoes_analisadas:,}")
    safe_print(f"  • Total de Registros de Classificacao: {len(classifs_rows):,}")
    
    if questoes_anomalas == 0:
        safe_print("  • STATUS: Redundancia 100% integra! Todas as questoes possuem associacao completa com seus ancestrais. [INTEGRO]")
        cur.close()
        conn.close()
        return
        
    safe_print(f"  • STATUS: Redundancia INCOMPLETA! [ANOMALIA DETECTADA]")
    safe_print(f"  • Questoes com Redundancia Ausente: {questoes_anomalas:,} ({(questoes_anomalas/total_questoes_analisadas*100):.2f}% das questoes)")
    safe_print(f"  • Total de Classificacoes Hierarquicas Faltantes: {total_redundancias_faltantes:,} registros")
    
    # 6. Exibir exemplos de anomalias
    safe_print("\nPrimeiros 5 exemplos de questoes com classificacao incompleta:")
    exemplos_impressos = 0
    for qid, tids_associados in classifs.items():
        ancestrais_necessarios = set()
        for tid in tids_associados:
            if tid in topicos:
                ancestrais_necessarios.update(get_ancestrais(tid))
        ancestrais_faltantes = ancestrais_necessarios - tids_associados
        
        if ancestrais_faltantes:
            exemplos_impressos += 1
            safe_print(f"  Questao {qid}:")
            # Mostrar os tópicos que ela tem classificados
            nomes_atuais = [f"'{topicos[tid]['nome']}' ({tid})" for tid in tids_associados if tid in topicos]
            safe_print(f"    - Associacoes atuais: {', '.join(nomes_atuais)}")
            # Mostrar os tópicos pais faltantes
            nomes_faltantes = [f"'{topicos[tid]['nome']}' ({tid})" for tid in ancestrais_faltantes if tid in topicos]
            safe_print(f"    - Ancestrais faltantes a adicionar: {', '.join(nomes_faltantes)}")
            if exemplos_impressos >= 5:
                break

    # 7. Executar a correção caso solicitado
    if execute_fix:
        safe_print(f"\n[LOG] Iniciando correcao no banco de dados...")
        safe_print(f"[LOG] Preparando para inserir {len(insercoes_necessarias)} registros faltantes...")
        
        conn.autocommit = False
        try:
            # Usar INSERT IGNORE para garantir segurança absoluta
            insert_query = "INSERT IGNORE INTO classificacao_questao (id_questao, id_topico) VALUES (%s, %s)"
            
            for i, (qid, tid) in enumerate(insercoes_necessarias, 1):
                cur.execute(insert_query, (qid, tid))
                if i % 10000 == 0:
                    safe_print(f"  - Inseridos {i}/{len(insercoes_necessarias)} registros ({(i/len(insercoes_necessarias)*100):.1f}%)")
                    
            conn.commit()
            safe_print(f"\n[LOG] SUCESSO! {len(insercoes_necessarias)} classificacoes redundantes foram inseridas e salvas. [OK]")
        except Exception as ex_db:
            conn.rollback()
            safe_print(f"[ERRO] Falha durante a execucao da correcao: {ex_db}")
        finally:
            cur.close()
            conn.close()
    else:
        cur.close()
        conn.close()
        safe_print(f"\n[DICA] Para corrigir automaticamente essas anomalias, rode o script passando o parametro '--fix':")
        safe_print(f"       python verificar_redundancia_topicos.py --fix")
        
    safe_print(f"\n[LOG] Tempo de execucao: {time.time() - start_time:.2f} segundos.")

if __name__ == "__main__":
    main()
