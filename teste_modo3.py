#!/usr/bin/env python3
"""
Script de teste para verificar a consulta SQL do modo 3
"""

import mysql.connector

# Configurações do banco
DB_CONFIG = {
    "host": "localhost",
    "user": "root",
    "password": "El@mysql.32",
    "database": "qconcursos"
}

def testar_consulta_modo3():
    """Testa a nova consulta SQL do modo 3"""
    
    print("=== TESTE DA CONSULTA SQL DO MODO 3 ===")
    print()
    
    # Conectar ao banco
    try:
        conn = mysql.connector.connect(**DB_CONFIG)
        cursor = conn.cursor(dictionary=True)
        print("✅ Conexão com o banco estabelecida")
    except Exception as e:
        print(f"❌ Erro ao conectar: {e}")
        return
    
    # Testar para ENARE
    print("\n--- Testando para ENARE ---")
    query_enare = """
    SELECT 
        q.*
    FROM questaoresidencia q
    WHERE q.alternativaE IS NULL
      AND q.comentario IS NOT NULL
      AND CHAR_LENGTH(q.comentario) >= 400
      AND q.ano >= 2016
      AND (q.instituicao LIKE '%ENARE%' OR q.instituicao LIKE '%REVALIDA NACIONAL%')
    ORDER BY q.ano DESC, q.questao_id
    LIMIT 5
    """
    
    try:
        cursor.execute(query_enare)
        questoes_enare = cursor.fetchall()
        print(f"✅ Questões ENARE encontradas: {len(questoes_enare)}")
        
        if questoes_enare:
            print("Primeiras 3 questões ENARE:")
            for i, q in enumerate(questoes_enare[:3], 1):
                print(f"  {i}. ID: {q['questao_id']}, Ano: {q['ano']}, Instituição: {q['instituicao']}")
                print(f"     Área: {q['area']}, Comentário: {len(q['comentario'])} chars")
    except Exception as e:
        print(f"❌ Erro na consulta ENARE: {e}")
    
    # Testar para REVALIDA NACIONAL
    print("\n--- Testando para REVALIDA NACIONAL ---")
    query_revalida = """
    SELECT 
        q.*
    FROM questaoresidencia q
    WHERE q.alternativaE IS NULL
      AND q.comentario IS NOT NULL
      AND CHAR_LENGTH(q.comentario) >= 400
      AND q.ano >= 2016
      AND (q.instituicao LIKE '%REVALIDA NACIONAL%' OR q.instituicao LIKE '%REVALIDA NACIONAL%')
    ORDER BY q.ano DESC, q.questao_id
    LIMIT 5
    """
    
    try:
        cursor.execute(query_revalida)
        questoes_revalida = cursor.fetchall()
        print(f"✅ Questões REVALIDA NACIONAL encontradas: {len(questoes_revalida)}")
        
        if questoes_revalida:
            print("Primeiras 3 questões REVALIDA NACIONAL:")
            for i, q in enumerate(questoes_revalida[:3], 1):
                print(f"  {i}. ID: {q['questao_id']}, Ano: {q['ano']}, Instituição: {q['instituicao']}")
                print(f"     Área: {q['area']}, Comentário: {len(q['comentario'])} chars")
    except Exception as e:
        print(f"❌ Erro na consulta REVALIDA: {e}")
    
    # Testar contagem total
    print("\n--- Contagem total ---")
    query_count = """
    SELECT 
        COUNT(*) as total,
        COUNT(CASE WHEN q.instituicao LIKE '%ENARE%' THEN 1 END) as enare_count,
        COUNT(CASE WHEN q.instituicao LIKE '%REVALIDA NACIONAL%' THEN 1 END) as revalida_count
    FROM questaoresidencia q
    WHERE q.alternativaE IS NULL
      AND q.comentario IS NOT NULL
      AND CHAR_LENGTH(q.comentario) >= 400
      AND q.ano >= 2016
      AND (q.instituicao LIKE '%ENARE%' OR q.instituicao LIKE '%REVALIDA NACIONAL%')
    """
    
    try:
        cursor.execute(query_count)
        resultado = cursor.fetchone()
        print(f"✅ Total de questões: {resultado['total']}")
        print(f"   - ENARE: {resultado['enare_count']}")
        print(f"   - REVALIDA NACIONAL: {resultado['revalida_count']}")
    except Exception as e:
        print(f"❌ Erro na contagem: {e}")
    
    # Fechar conexão
    cursor.close()
    conn.close()
    print("\n✅ Teste concluído!")

if __name__ == "__main__":
    testar_consulta_modo3()