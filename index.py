import PyPDF2
import csv
import unicodedata
import os
import zipfile
import re

def remover_acentos_e_sinais(texto):
    """Remove acentos, sinais, números ordinais e substitui Ç por C"""
    if not isinstance(texto, str):
        return texto
    
    # Normaliza para forma decomposed (separa caracteres de seus acentos)
    texto = unicodedata.normalize('NFD', texto)
    # Remove todos os caracteres que não são ASCII (acento, til, etc.)
    texto = ''.join(c for c in texto if unicodedata.category(c) != 'Mn')
    
    # Substitui Ç por C (maiúsculo e minúsculo)
    texto = texto.replace('Ç', 'C').replace('ç', 'c')
    
    # Remove símbolos como 'nº', 'ª', etc.
    texto = re.sub(r'nº|ª|º', '', texto)
    
    # Remove outros caracteres não alfanuméricos (somente letras e números)
    texto = re.sub(r'[^a-zA-Z0-9\s]', '', texto)
    
    return texto

def extrair_texto_pdf(caminho_pdf):
    """Extrai o texto do PDF"""
    with open(caminho_pdf, 'rb') as arquivo_pdf:
        leitor_pdf = PyPDF2.PdfReader(arquivo_pdf)
        texto = ""
        for pagina in leitor_pdf.pages:
            texto += pagina.extract_text()
    return texto

def salvar_csv(dados, caminho_csv):
    """Salva os dados em um arquivo CSV"""
    with open(caminho_csv, mode='w', newline='', encoding='utf-8') as arquivo:
        escritor_csv = csv.writer(arquivo)
        for linha in dados:
            escritor_csv.writerow(linha)

def compactar_arquivo(caminho_csv, caminho_zip):
    """Compacta o arquivo CSV em um arquivo ZIP"""
    with zipfile.ZipFile(caminho_zip, 'w', zipfile.ZIP_DEFLATED) as zipf:
        zipf.write(caminho_csv, os.path.basename(caminho_csv))
        
        
def processar_dados(texto):
    """Processa o texto extraído do PDF, mantendo a estrutura de tabela e expandindo abreviações"""
    # Divide o texto em linhas
    linhas = texto.split("\n")
    
    # Nova lista para armazenar linhas processadas
    linhas_consolidadas = []
    
    # Variável para acumular linhas parciais
    linha_atual = []
    
    for linha in linhas:
        # Remove espaços extras e divide a linha
        colunas = linha.split()
        
        # Verifica se a linha contém informações relevantes
        if colunas:
            # Se a primeira coluna corresponde a um procedimento conhecido
            if any(proc in linha.upper() for proc in [
                "TESTE", "TESTES", "PROCEDIMENTO"
            ]):
                # Se há uma linha parcial acumulada, tenta consolidá-la
                if linha_atual:
                    linhas_consolidadas.append(' '.join(linha_atual))
                
                # Reinicia a linha atual
                linha_atual = colunas
            else:
                # Adiciona as colunas à linha atual
                linha_atual.extend(colunas)
    
    # Adiciona a última linha se existir
    if linha_atual:
        linhas_consolidadas.append(' '.join(linha_atual))
    
    # Processa as linhas consolidadas
    dados_processados = []
    
    # Adiciona cabeçalho com todas as colunas
    cabecalho = [
        "PROCEDIMENTO", "RN", "VIGENCIA", "OD", "AMB", "HCO", 
        "HSO", "REF", "PAC", "DUT", "SUBGRUPO", "GRUPO", "CAPITULO"
    ]
    dados_processados.append(cabecalho)
    
    for linha in linhas_consolidadas:
        # Tenta dividir a linha mantendo a estrutura de colunas
        colunas = linha.split()
        
        # Processa cada coluna removendo acentos e sinais
        linha_limpa = [remover_acentos_e_sinais(celula.strip()) for celula in colunas]
        
        # Substituição de abreviações
        for i, coluna in enumerate(linha_limpa):
            if coluna == "OD":
                linha_limpa[i] = "SEG_ODONTOLOGICA"
            elif coluna == "AMB":
                linha_limpa[i] = "SEG_AMBULATORIAL"
        
        # Preenche com vazios se não tiver todas as colunas
        while len(linha_limpa) < len(cabecalho):
            linha_limpa.append("")
        
        # Trunca para o número de colunas do cabeçalho se necessário
        linha_limpa = linha_limpa[:len(cabecalho)]
        
        if any(linha_limpa):  # Adiciona apenas se não for totalmente vazia
            dados_processados.append(linha_limpa)
    
    return dados_processados
        

#def processar_dados(texto):
#    """Processa o texto extraído do PDF, divide em linhas e remove acentos e sinais"""
#    linhas = texto.split("\n")
#    dados_processados = []
    
#    for linha in linhas:
#        # Limpeza de acentos, sinais e adição dos dados processados
#        linha_limpa = [remover_acentos_e_sinais(celula.strip()) for celula in linha.split() if celula.strip()]
#        
#        if linha_limpa:  # Evita adicionar linhas vazias
#            dados_processados.append(linha_limpa)
#    
#    return dados_processados

def main():
    # Caminho do PDF fixo
    caminho_pdf = "D:\\Intuitive\\b-teste\\anexo.pdf"  # Substitua com o nome real do arquivo PDF
    caminho_csv = "Teste_Willian_de_Sousa_Mota.csv"  # Nome do arquivo CSV de saída
    caminho_zip = "Teste_Willian_de_Sousa_Mota.zip"  # Nome do arquivo ZIP de saída
    
    # Extrai o texto do PDF
    print("Iniciando extração de texto do PDF...")
    texto_pdf = extrair_texto_pdf(caminho_pdf)
    
    # Processa os dados extraídos
    print("Processando os dados extraídos...")
    dados_processados = processar_dados(texto_pdf)
    
    # Salva os dados processados no CSV
    print("Salvando os dados em CSV...")
    salvar_csv(dados_processados, caminho_csv)
    
    # Compacta o arquivo CSV
    print("Compactando o arquivo CSV...")
    compactar_arquivo(caminho_csv, caminho_zip)
    
    # Remove o arquivo CSV temporário
    os.remove(caminho_csv)
    
    print(f"Dados salvos e compactados com sucesso em {caminho_zip}")

if __name__ == "__main__":
    main()
