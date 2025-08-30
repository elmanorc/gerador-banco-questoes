import mysql.connector

# Conexão com o banco
conn = mysql.connector.connect(
    host="localhost",
    user="root",
    password="El@mysql.32",
    database="qconcursos"
)
cursor = conn.cursor()

# Texto do comentário (use """ para multilinha)
comentario = """Quando uma criança está gravemente desidratada, o tratamento é feito em 3 etapas:

1️⃣ **EXPANSÃO** 💧

*   Se menos que 5 anos: realizar 20 ml/Kg de cristalóide em Soro Fisiológico a 0,9% em 30 min. 
*   Sendo maior que 5 anos: realizar 30 ml/kg de Sf 0,9% em 30 min, seguido por 70 ml/kg de Ringer Lactato em 2 horas e 30 minutos.

2️⃣ **Manutenção** ⚖️

*   Nas próximas 24 horas, é preciso restabelecer o equilíbrio hídrico com uma solução de manutenção, calculada com a regra de Holliday-Segar.

<img src="https://estrategia-prod-questoes.s3.amazonaws.com/images/2AFF3A76-78A3-2CA1-3EDB-4BAE6C5C25B9/2AFF3A76-78A3-2CA1-3EDB-4BAE6C5C25B9-400.png" class="questions-inline-img cursor-pointer">

*   Essa regra define a necessidade hídrica diária para as próximas 24 horas.
*   A escolha da solução depende da necessidade:
    *   **Solução hipotônica:** SG 5% (4 partes) + SF 0,9% (1 parte) + 2 ml de KCl 19,1% para cada 100 ml da solução.
    *   **Solução isotônica:** Necessidade hídrica em SG5%, para cada 1000 ml de soro, usamos 40ml de Nacl 20% e 10 ml de KCl 19,1%.

3️⃣ **Reposição** 🔄

*   Para repor as perdas nas próximas 24 horas, usar: SG 5% (1 parte) + SF 0,9% (1 parte), 50 ml/kg."""

# Atualiza a questão com ID específico
sql = "UPDATE questaoresidencia SET comentario = %s WHERE codigo = %s"
cursor.execute(sql, (comentario, 400170723))  # substitua 123 pelo ID da questão

# Confirma no banco
conn.commit()

print(cursor.rowcount, "registro atualizado.")

cursor.close()
conn.close()
