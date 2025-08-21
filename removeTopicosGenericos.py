import mysql.connector

conn = mysql.connector.connect(
    host="localhost",
    user="root",
    password="El@mysql.32",
    database="qconcursos"
)
cur = conn.cursor(dictionary=True)

# 1. Carregar todos os tópicos e montar hierarquia
cur.execute("SELECT id, id_pai, nome FROM topico")
topicos = {row["id"]: {"pai": row["id_pai"], "nome": row["nome"]} for row in cur.fetchall()}

def get_ancestrais(tid):
    """Retorna todos os ancestrais de um tópico"""
    ancestrais = []
    while topicos[tid]["pai"] is not None:
        tid = topicos[tid]["pai"]
        ancestrais.append(tid)
    return ancestrais

# 2. Carregar classificações
cur.execute("SELECT id_questao, id_topico FROM classificacao_questao")
classifs = {}
for row in cur.fetchall():
    classifs.setdefault(row["id_questao"], []).append(row["id_topico"])

# 3. Identificar quais devem ser removidos
remover = []
for qid, tids in classifs.items():
    todos_ancestrais = set()
    for tid in tids:
        todos_ancestrais.update(get_ancestrais(tid))
    for tid in tids:
        if tid in todos_ancestrais:
            remover.append((qid, tid))

print(f"Total de classificações redundantes a remover: {len(remover)}")

# 4. Executar as remoções
for qid, tid in remover:
    cur.execute(
        "DELETE FROM classificacao_questao WHERE id_questao = %s AND id_topico = %s",
        (qid, tid)
    )

conn.commit()
cur.close()
conn.close()

print("Remoções concluídas com sucesso ✅")
