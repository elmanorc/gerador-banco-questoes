import mysql.connector

conn = mysql.connector.connect(
    host="localhost",
    user="root",
    password="El@mysql.32",
    database="qconcursos"
)
cur = conn.cursor(dictionary=True)

# 1. Carregar todos os tópicos
cur.execute("SELECT id, id_pai, nome FROM topico")
topicos = {row["id"]: {"pai": row["id_pai"], "nome": row["nome"]} for row in cur.fetchall()}

def get_ancestrais(tid):
    """Retorna lista de ancestrais até a raiz (raiz é o último elemento da lista)."""
    ancestrais = []
    while topicos[tid]["pai"] is not None:
        tid = topicos[tid]["pai"]
        ancestrais.append(tid)
    return ancestrais

def get_raiz(tid):
    """Retorna o ID do tópico raiz de um tópico."""
    while topicos[tid]["pai"] is not None:
        tid = topicos[tid]["pai"]
    return tid

def get_raiz_nome(tid):
    """Retorna o nome do tópico raiz de um tópico."""
    return topicos[get_raiz(tid)]["nome"]

def get_profundidade(tid):
    """Profundidade = nº de ancestrais + 1"""
    return len(get_ancestrais(tid)) + 1

# 2. Definir prioridades entre áreas
prioridade_areas = {
    "Pediatria": 1,
    "Clínica Médica": 2,
    "Cirurgia": 3,
    "Ginecologia": 4,
    "Obstetrícia": 5,
    "Medicina Preventiva": 6,
    "Outros": 99  # sempre por último
}

# 3. Carregar classificações das questões
cur.execute("SELECT id_questao, id_topico FROM classificacao_questao")
classifs = {}
for row in cur.fetchall():
    classifs.setdefault(row["id_questao"], []).append(row["id_topico"])

# 4. Obter a área atual de cada questão
cur.execute("SELECT questao_id, area FROM questaoresidencia")
areas_atual = {row["questao_id"]: row["area"] for row in cur.fetchall()}

# 5. Analisar questões que precisam ter a área corrigida
correcoes = []
for qid, tids in classifs.items():
    # Escolhe pelo critério: maior profundidade -> prioridade
    tid_escolhido = max(
        tids,
        key=lambda tid: (
            get_profundidade(tid),
            - (1000 - prioridade_areas.get(get_raiz_nome(tid), 999))  # menor número = maior prioridade
        )
    )
    raiz_nome = get_raiz_nome(tid_escolhido)
    area_atual = areas_atual.get(qid)
    if area_atual != raiz_nome:
        correcoes.append((qid, area_atual, raiz_nome))

# 6. Mostrar análise das correções
print(f"Total de questões que precisam ter a área corrigida: {len(correcoes)}\n")

# Mostrar estatísticas por área
estatisticas = {}
for qid, atual, novo in correcoes:
    key = f"{atual} -> {novo}"
    estatisticas[key] = estatisticas.get(key, 0) + 1

print("Estatísticas das correções por área:")
for mudanca, count in sorted(estatisticas.items(), key=lambda x: x[1], reverse=True):
    print(f"  {mudanca}: {count} questões")

# Mostrar exemplos das primeiras 20 correções
print(f"\nPrimeiras 20 correções que serão feitas:")
for qid, atual, novo in correcoes[:20]:
    print(f"  Questão {qid}: '{atual}' -> '{novo}'")

if len(correcoes) > 20:
    print(f"  ... e mais {len(correcoes) - 20} questões.")

# 7. Confirmar antes de executar as alterações
if len(correcoes) == 0:
    print("\nNenhuma correção necessária. Todas as áreas já estão corretas.")
    cur.close()
    conn.close()
    exit()

print(f"\n{'='*60}")
print(f"ATENÇÃO: Esta operação irá alterar {len(correcoes)} questões no banco de dados!")
print(f"{'='*60}")

confirmacao = input("\nDeseja prosseguir com as alterações? (digite 'SIM' para confirmar): ").strip()

if confirmacao != 'SIM':
    print("Operação cancelada pelo usuário.")
    cur.close()
    conn.close()
    exit()

# 8. Executar as correções no banco de dados
print(f"\nIniciando correções no banco de dados...")
print(f"Processando {len(correcoes)} questões...")

# Usar transação para garantir consistência
conn.autocommit = False

try:
    # Preparar statement de UPDATE
    update_query = "UPDATE questaoresidencia SET area = %s WHERE questao_id = %s"
    
    correcoes_executadas = 0
    for i, (qid, area_atual, nova_area) in enumerate(correcoes, 1):
        # Executar UPDATE
        cur.execute(update_query, (nova_area, qid))
        correcoes_executadas += 1
        
        # Mostrar progresso a cada 100 questões
        if i % 100 == 0 or i == len(correcoes):
            print(f"  Processadas {i}/{len(correcoes)} questões ({(i/len(correcoes)*100):.1f}%)")
    
    # Fazer commit das alterações
    conn.commit()
    print(f"\n✅ SUCESSO: {correcoes_executadas} questões atualizadas com sucesso!")
    
    # Verificar algumas alterações
    print(f"\nVerificando algumas alterações...")
    verificacoes = correcoes[:5]  # Verificar primeiras 5
    
    for qid, area_antiga, area_nova in verificacoes:
        cur.execute("SELECT area FROM questaoresidencia WHERE questao_id = %s", (qid,))
        resultado = cur.fetchone()
        if resultado and resultado['area'] == area_nova:
            print(f"  ✓ Questão {qid}: área atualizada para '{area_nova}'")
        else:
            print(f"  ✗ Questão {qid}: ERRO na atualização!")
    
    print(f"\n🎯 RESUMO FINAL:")
    print(f"  • Questões analisadas: {len(classifs)}")
    print(f"  • Questões corrigidas: {correcoes_executadas}")
    print(f"  • Transação commitada com sucesso")
    
except Exception as e:
    # Fazer rollback em caso de erro
    conn.rollback()
    print(f"\n❌ ERRO durante a atualização: {str(e)}")
    print(f"Todas as alterações foram revertidas (rollback).")
    
finally:
    cur.close()
    conn.close()
    print(f"\nConexão com banco de dados fechada.")
