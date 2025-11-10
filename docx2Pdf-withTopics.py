import os
import win32com.client

# Solicita o caminho do arquivo DOCX
input_path = input("Digite o caminho completo do arquivo DOCX: ").strip('"')

# Verifica se o arquivo existe
if not os.path.isfile(input_path) or not input_path.lower().endswith('.docx'):
    print("❌ Arquivo inválido. Verifique o caminho e tente novamente.")
else:
    # Define o caminho do PDF no mesmo diretório
    base_name = os.path.splitext(input_path)[0]
    output_path = base_name + ".pdf"

    # Inicia o Word
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False  # oculta a janela do Word

    # Abre o documento
    doc = word.Documents.Open(input_path)

    # Exporta para PDF com estrutura de tópicos (headings)
    doc.ExportAsFixedFormat(
        OutputFileName=output_path,
        ExportFormat=17,          # wdExportFormatPDF
        OpenAfterExport=False,    # não abrir automaticamente
        OptimizeFor=0,            # wdExportOptimizeForPrint
        CreateBookmarks=1         # 0=None, 1=Headings, 2=Word Bookmarks
    )

    # Fecha o documento e encerra o Word
    doc.Close(False)
    word.Quit()

    print(f"✅ PDF gerado com sucesso:\n{output_path}")
