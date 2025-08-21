# Gerador de Banco de Questões - Consulta SQL Específica

Este programa gera bancos de questões em formato DOCX usando uma consulta SQL específica com distribuição proporcional entre áreas médicas e organização hierárquica.

## Funcionalidades

### 1. Consulta SQL Específica
- **Query**: Usa `ROW_NUMBER() OVER (PARTITION BY q.area ORDER BY q.questao_id)` para seleção determinística
- **Filtros**: Questions sem alternativa E, com comentários ≥500 chars, ano ≥2020
- **Proporções**: Distribuição automática baseada no total N informado pelo usuário

### 2. Distribuição Proporcional por Área
- **Cirurgia**: 20% das questões
- **Clínica Médica**: 20% das questões  
- **Pediatria**: 20% das questões
- **Ginecologia**: 10% das questões
- **Obstetrícia**: 10% das questões
- **Medicina Preventiva e Social**: 20% das questões

### 3. Organização Hierárquica
- **Profundidade máxima**: Nível 4 (tópicos de nível 5+ são reagrupados no nível 4)
- **Estrutura**: Mantém hierarquia natural do banco de dados
- **Sequência**: Ordem específica das áreas médicas conforme solicitado

### 4. Estrutura do Documento
#### **Seção 1: Capa**
- Cabeçalho com logotipo centralizado (`/img/logotipo.png`)
- Título principal do banco
- Subtítulo com número total de questões

#### **Seção 2: Sumário**
- Nova página dedicada
- Índice automático (TOC) com links
- Cabeçalho limpo sem logotipo

#### **Seção 3+: Conteúdo das Questões**
- **Tópicos Nível 1**: Nova seção com quebra de página
  - Cabeçalho: Nome do tópico (ex: "Cirurgia")
- **Tópicos Nível 2**: Nova seção com quebra de página  
  - Cabeçalho: "Tópico Nível 1 > Tópico Nível 2" (ex: "Cirurgia > Cirurgia Geral")
- **Tópicos Nível 3-4**: Sem quebra de página, breadcrumbs no cabeçalho da seção pai

### 5. Controle de Qualidade
- **Controle de repetição**: Opção para evitar questões duplicadas
- **Reagrupamento**: Questões de níveis 5+ aparecem no nível 4 pai
- **Formatação profissional**: Estilos consistentes, imagens, comentários

## Estrutura de Arquivos

```
/
├── geradorBancosDeQuestoesPorTopico.py  # Script principal
├── img/
│   └── logotipo.png                     # Logotipo para cabeçalho da capa
├── files/                               # Arquivos de exemplo (opcionais)
└── README.md                            # Esta documentação
```

## Dependências

- `python-docx`: Geração de documentos DOCX
- `mysql.connector`: Conexão com banco de dados MySQL
- `markdown2`: Conversão de Markdown para HTML
- `BeautifulSoup4`: Parsing de HTML
- `Pillow`: Manipulação de imagens

## Uso

1. **Configure a conexão** no arquivo (variável `DB_CONFIG`)
2. **Adicione o logotipo** em `/img/logotipo.png`
3. **Execute**: `python geradorBancosDeQuestoesPorTopico.py`
4. **Informe o número N** de questões totais desejado
5. **Configure repetição** conforme necessário
6. **Arquivo gerado**: `banco_questoes_sql_{N}_{timestamp}.docx`

## Exemplo de Uso

```bash
python geradorBancosDeQuestoesPorTopico.py

# Saída:
=== GERADOR DE BANCO DE QUESTÕES COM CONSULTA SQL ESPECÍFICA ===
Número total de questões do banco (ex: 1000, 2000, 3000): 1000

[LOG] Distribuição proporcional para 1000 questões:
  - Cirurgia: 200 questões (20%)
  - Clínica Médica: 200 questões (20%)
  - Pediatria: 200 questões (20%)
  - Ginecologia: 100 questões (10%)
  - Obstetrícia: 100 questões (10%)
  - Medicina Preventiva e Social: 200 questões (20%)

Permitir repetição de questões no documento? (s/n, Enter para sim): n
```

## Estrutura do Documento Gerado

1. **Capa** (Seção 1)
   - Logotipo no cabeçalho
   - Título e subtítulo centralizados

2. **Sumário** (Seção 2)
   - Índice automático
   - Links para seções

3. **Conteúdo** (Seções 3+)
   - **1. Cirurgia** (Nova seção)
     - **1.1 Cirurgia Geral** (Nova seção)
       - **1.1.1 Técnicas Básicas** (Mesma seção)
         - **1.1.1.1 Sutura** (Questões agrupadas aqui)

## Logs Detalhados

O programa fornece logs informativos sobre:
- Conexão com banco de dados
- Execução da consulta SQL
- Construção da hierarquia de tópicos
- Reorganização de questões por nível
- Criação de seções e breadcrumbs
- Salvamento do arquivo final
