import mysql.connector
import os

def load_db_password():
    password_path = os.path.join(os.path.dirname(__file__), 'db_password.txt')
    with open(password_path, 'r', encoding='utf-8') as f:
        return f.read().strip()

def atualizar_areas():
    print("Conectando ao banco de dados...")
    conn = mysql.connector.connect(
        host="localhost",
        user="root",
        password=load_db_password(),
        database="qconcursos"
    )
    cursor = conn.cursor(dictionary=True)
    
    print("Buscando questões com área nula ou vazia...")
    cursor.execute("""
        SELECT questao_id 
        FROM questaoresidencia 
        WHERE area IS NULL OR area = '' OR area = 'None'
    """)
    questoes = cursor.fetchall()
    
    print(f"Encontradas {len(questoes)} questões sem área definida.")
    
    if not questoes:
        print("Nenhuma atualização necessária.")
        conn.close()
        return

    # Cache para evitar consultas repetidas (topico_id -> nome_raiz)
    raiz_cache = {}
    
    def get_nome_raiz(tid):
        if tid in raiz_cache:
            return raiz_cache[tid]
            
        current_id = tid
        current_nome = None
        
        while current_id is not None:
            cursor.execute("SELECT id, nome, id_pai FROM topico WHERE id = %s", (current_id,))
            row = cursor.fetchone()
            if not row:
                break
            current_nome = row['nome']
            if row['id_pai'] is None:
                break
            current_id = row['id_pai']
            
        raiz_cache[tid] = current_nome
        return current_nome

    atualizadas = 0
    sem_classificacao = 0

    for q in questoes:
        qid = q['questao_id']
        
        # Buscar as classificações da questão
        cursor.execute("SELECT id_topico FROM classificacao_questao WHERE id_questao = %s", (qid,))
        classificacoes = cursor.fetchall()
        
        if not classificacoes:
            sem_classificacao += 1
            continue
            
        # Pega a primeira classificação (assumindo que leva à área raiz correta)
        topico_id = classificacoes[0]['id_topico']
        nome_raiz = get_nome_raiz(topico_id)
        
        if nome_raiz:
            cursor.execute("UPDATE questaoresidencia SET area = %s WHERE questao_id = %s", (nome_raiz, qid))
            atualizadas += 1
            if atualizadas % 100 == 0:
                print(f"Atualizadas {atualizadas} questões...")
                conn.commit()
                
    conn.commit()
    print(f"Processo concluído!")
    print(f"Total de questões atualizadas: {atualizadas}")
    print(f"Questões sem classificação (continuam sem área): {sem_classificacao}")
    
    conn.close()

if __name__ == '__main__':
    atualizar_areas()
