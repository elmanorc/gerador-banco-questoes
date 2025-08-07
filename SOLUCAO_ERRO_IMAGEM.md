# Solução para Erro UnrecognizedImageError

## Problema
O erro `UnrecognizedImageError` estava ocorrendo na linha 421 durante a adição de imagens no documento Word usando `document.add_picture(img_path, width=max_width)`.

## Causas Principais

### 1. **Formato de Imagem Não Suportado**
- O `python-docx` suporta apenas formatos específicos: PNG, JPEG, GIF, BMP, TIFF
- Se a imagem estiver em formato não suportado (como WebP, SVG, etc.), o erro ocorrerá

### 2. **Arquivo de Imagem Corrompido**
- O arquivo pode estar corrompido ou incompleto
- O arquivo pode ter extensão incorreta (ex: arquivo .jpg com conteúdo PNG)

### 3. **Problemas de Permissão**
- O arquivo pode não estar acessível devido a permissões
- O caminho pode estar incorreto

### 4. **Dependências Ausentes**
- Falta de bibliotecas de processamento de imagem (Pillow/PIL)

## Soluções Implementadas

### 1. **Função de Verificação Robusta**
Foi criada a função `verificar_e_adicionar_imagem()` que:
- Verifica se o arquivo existe
- Verifica se é um arquivo válido
- Verifica o tamanho do arquivo (não pode ser 0)
- Verifica o formato MIME da imagem
- Captura e trata especificamente o erro `UnrecognizedImageError`

### 2. **Tratamento de Erro Específico**
```python
try:
    document.add_picture(img_path, width=max_width)
except UnrecognizedImageError as e:
    print(f"[ERRO] Formato de imagem não reconhecido: {img_path}")
    print(f"[ERRO] Detalhes: {str(e)}")
except Exception as e:
    print(f"[ERRO] Erro ao adicionar imagem {img_path}: {str(e)}")
```

### 3. **Logs Detalhados**
- Logs informativos para cada etapa do processo
- Mensagens de erro específicas para diferentes tipos de problema
- Identificação clara de qual imagem está causando problema

### 4. **Fallback Graceful**
- Quando uma imagem falha, o sistema continua funcionando
- Adiciona uma mensagem indicando que a imagem não foi encontrada ou é inválida
- Não interrompe a geração do documento

## Melhorias Implementadas

### 1. **Verificação de Formato MIME**
```python
mime_type, _ = mimetypes.guess_type(img_path)
if mime_type and not mime_type.startswith('image/'):
    print(f"[AVISO] Arquivo não parece ser uma imagem válida: {img_path}")
```

### 2. **Verificação de Tamanho**
```python
file_size = os.path.getsize(img_path)
if file_size == 0:
    print(f"[AVISO] Arquivo de imagem vazio: {img_path}")
```

### 3. **Import de Exceção Específica**
```python
from docx.exceptions import UnrecognizedImageError
```

## Como Usar

O código agora trata automaticamente os erros de imagem. Se uma imagem falhar:

1. **Logs detalhados** serão exibidos no console
2. **O processo continua** sem interrupção
3. **Uma mensagem** será inserida no documento indicando o problema
4. **O documento será gerado** normalmente

## Recomendações Adicionais

### 1. **Verificar Formatos de Imagem**
- Certifique-se de que as imagens estão em formatos suportados (PNG, JPEG, GIF, BMP, TIFF)
- Converta imagens WebP, SVG ou outros formatos não suportados

### 2. **Verificar Integridade dos Arquivos**
- Teste se as imagens abrem corretamente em outros programas
- Verifique se os arquivos não estão corrompidos

### 3. **Instalar Dependências**
```bash
pip install Pillow
```

### 4. **Verificar Permissões**
- Certifique-se de que o script tem permissão para acessar os arquivos de imagem
- Verifique se os caminhos estão corretos

## Exemplo de Uso

```python
# Antes (pode causar erro)
document.add_picture(img_path, width=max_width)

# Depois (tratamento robusto)
if not verificar_e_adicionar_imagem(document, img_path, max_width):
    document.add_paragraph(f"[Imagem não encontrada ou inválida: {img_filename}]")
```

Esta solução garante que o gerador de bancos de questões continue funcionando mesmo quando encontrar imagens problemáticas, fornecendo logs detalhados para facilitar a identificação e correção dos problemas. 