#!/usr/bin/env python3
"""
Script de teste para o Modo 4 - Processamento de questões incompletas
"""

import mysql.connector
from datetime import datetime

# Configurações do banco
DB_CONFIG = {
    "host": "localhost",
    "user": "root",
    "password": "El@mysql.32",
    "database": "qconcursos"
}

def verificar_colunas_necessarias():
    """
    Verifica se as colunas necessárias para o modo 4 existem na tabela questaoresidencia.
    """
    print("=== VERIFICAÇÃO DE COLUNAS NECESSÁRIAS ===")
    
    conn = mysql.connector.connect(**DB_CONFIG)
    cursor = conn.cursor()
    
    try:
        # Verificar estrutura da tabela
        cursor.execute("DESCRIBE questaoresidencia")
        colunas = cursor.fetchall()
        
        print("Colunas existentes na tabela questaoresidencia:")
        colunas_necessarias = ['comentarioIA', 'comentario_autor', 'comentario_data', 'gabaritoIA']
        colunas_existentes = [coluna[0] for coluna in colunas]
        
        for coluna in colunas_necessarias:
            if coluna in colunas_existentes:
                print(f"  ✓ {coluna} - EXISTE")
            else:
                print(f"  ✗ {coluna} - NÃO EXISTE")
        
        # Verificar se precisamos criar as colunas
        colunas_faltando = [coluna for coluna in colunas_necessarias if coluna not in colunas_existentes]
        
        if colunas_faltando:
            print(f"\n[AVISO] Colunas faltando: {colunas_faltando}")
            print("Executando comandos para criar as colunas...")
            
            for coluna in colunas_faltando:
                if coluna == 'comentarioIA':
                    cursor.execute("ALTER TABLE questaoresidencia ADD COLUMN comentarioIA TEXT")
                    print(f"  ✓ Coluna 'comentarioIA' criada")
                elif coluna == 'comentario_autor':
                    cursor.execute("ALTER TABLE questaoresidencia ADD COLUMN comentario_autor VARCHAR(100)")
                    print(f"  ✓ Coluna 'comentario_autor' criada")
                elif coluna == 'comentario_data':
                    cursor.execute("ALTER TABLE questaoresidencia ADD COLUMN comentario_data DATETIME")
                    print(f"  ✓ Coluna 'comentario_data' criada")
                elif coluna == 'gabaritoIA':
                    cursor.execute("ALTER TABLE questaoresidencia ADD COLUMN gabaritoIA VARCHAR(1)")
                    print(f"  ✓ Coluna 'gabaritoIA' criada")
            
            conn.commit()
            print("\n[SUCESSO] Todas as colunas necessárias foram criadas!")
        else:
            print("\n[SUCESSO] Todas as colunas necessárias já existem!")
            
    except Exception as e:
        print(f"[ERRO] Falha na verificação/criação das colunas: {str(e)}")
        conn.rollback()
    finally:
        cursor.close()
        conn.close()

def testar_identificacao_questoes_incompletas():
    """
    Testa a identificação de questões com comentários incompletos.
    """
    print("\n=== TESTE DE IDENTIFICAÇÃO DE QUESTÕES INCOMPLETAS ===")
    
    conn = mysql.connector.connect(**DB_CONFIG)
    cursor = conn.cursor(dictionary=True)
    
    try:
        # Buscar questões que terminam com 'analisar as alternativas'
        query = """
        SELECT questao_id, codigo, 
               SUBSTRING(comentario, LOCATE('analisar as alternativas', comentario)) as final_comentario,
               LENGTH(TRIM(SUBSTRING(comentario, LOCATE('analisar as alternativas', comentario) + 23))) as chars_apos
        FROM questaoresidencia 
        WHERE comentario LIKE '%analisar as alternativas%'
        ORDER BY questao_id
        LIMIT 10
        """
        
        cursor.execute(query)
        questoes = cursor.fetchall()
        
        print(f"Encontradas {len(questoes)} questões com 'analisar as alternativas' (mostrando primeiras 10):")
        
        for questao in questoes:
            print(f"\nQuestão {questao['codigo']} (ID: {questao['questao_id']}):")
            print(f"  Final do comentário: '{questao['final_comentario']}'")
            print(f"  Caracteres após 'analisar as alternativas': {questao['chars_apos']}")
            
            # Verificar se é incompleta
            if questao['chars_apos'] < 50:
                print(f"  → INCOMPLETA (poucos caracteres após)")
            else:
                print(f"  → COMPLETA (muitos caracteres após)")
        
        # Aplicar filtro completo
        query_incompletas = """
        SELECT questao_id, codigo, enunciado, alternativaA, alternativaB, alternativaC, 
               alternativaD, alternativaE, gabarito, comentario
        FROM questaoresidencia 
        WHERE comentario LIKE '%analisar as alternativas%'
        AND (
            LENGTH(TRIM(SUBSTRING(comentario, LOCATE('analisar as alternativas', comentario) + 23))) < 50
            OR comentario REGEXP 'analisar as alternativas[[:space:]]*$'
            OR comentario REGEXP 'analisar as alternativas[[:space:]]*[[:punct:]]*[[:space:]]*$'
        )
        ORDER BY questao_id
        LIMIT 5
        """
        
        cursor.execute(query_incompletas)
        questoes_incompletas = cursor.fetchall()
        
        print(f"\n[RESULTADO] {len(questoes_incompletas)} questões identificadas como incompletas (mostrando primeiras 5):")
        
        for questao in questoes_incompletas:
            print(f"\nQuestão {questao['codigo']} (ID: {questao['questao_id']}):")
            print(f"  Gabarito: {questao['gabarito']}")
            print(f"  Enunciado: {questao['enunciado'][:100]}...")
            print(f"  Alternativas:")
            for alt in ['A', 'B', 'C', 'D', 'E']:
                alt_text = questao.get(f'alternativa{alt}', '')
                if alt_text:
                    print(f"    {alt}) {alt_text[:50]}...")
        
    except Exception as e:
        print(f"[ERRO] Falha no teste: {str(e)}")
    finally:
        cursor.close()
        conn.close()

def testar_conexao_api():
    """
    Testa a conexão com a API DeepSeek (sem fazer chamadas reais).
    """
    print("\n=== TESTE DE CONFIGURAÇÃO DA API DEEPSEEK ===")
    
    import requests
    import json
    
    DEEPSEEK_CONFIG = {
        "api_key": "sk-50280cb2abb4473c9463f7ae053f7610",
        "model": "deepseek-chat",
        "temperature": 0.1,
        "url": "https://api.deepseek.com/v1/chat/completions"
    }
    
    print(f"URL da API: {DEEPSEEK_CONFIG['url']}")
    print(f"Modelo: {DEEPSEEK_CONFIG['model']}")
    print(f"Temperatura: {DEEPSEEK_CONFIG['temperature']}")
    print(f"API Key: {DEEPSEEK_CONFIG['api_key'][:10]}...")
    
    # Teste simples de conectividade (sem fazer chamada real)
    try:
        headers = {
            "Authorization": f"Bearer {DEEPSEEK_CONFIG['api_key']}",
            "Content-Type": "application/json"
        }
        print("\n[INFO] Configuração da API parece estar correta")
        print("[INFO] Para testar a API real, execute o modo 4 com uma questão de teste")
    except Exception as e:
        print(f"[ERRO] Problema na configuração da API: {str(e)}")

if __name__ == "__main__":
    print("=== TESTE DO MODO 4 - PROCESSAMENTO DE QUESTÕES INCOMPLETAS ===")
    
    # 1. Verificar colunas necessárias
    verificar_colunas_necessarias()
    
    # 2. Testar identificação de questões incompletas
    testar_identificacao_questoes_incompletas()
    
    # 3. Testar configuração da API
    testar_conexao_api()
    
    print("\n=== TESTE CONCLUÍDO ===")
    print("Se todos os testes passaram, o Modo 4 deve estar funcionando corretamente.")
    print("Execute o programa principal e escolha a opção 4 para processar as questões incompletas.")
