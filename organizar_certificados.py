import os
import shutil
import pandas as pd
import re 
from unidecode import unidecode


# --- CONFIGURAÇÕES ---
PASTA_ORIGEM = r"C:\Caminho\Para\Pasta_Baixada"
PASTA_DESTINO = r"C:\Caminho\Para\Pasta_Organizada"
ARQUIVO_PLANILHA = r"C:\Caminho\Para\Lista_Alunos.xlsx"

# Nomes das colunas na sua planilha
COLUNA_NOME = "ALUNO"
COLUNA_ESCOLA = "ESTABELECIMENTO"

def normalizar_texto(texto):
   
    if not isinstance(texto, str):
        return ""
    return unidecode(texto).lower().strip()

def organizar_arquivos():
    # 1. Carregar a planilha
    print("Lendo planilha de alunos...")
    try:
        df = pd.read_excel(ARQUIVO_PLANILHA)
    except Exception as e:
        print(f"Erro ao ler planilha: {e}")
        return

    # 2. Mapeamento
    mapa_alunos = {}
    for index, row in df.iterrows():
        nome_norm = normalizar_texto(row[COLUNA_NOME])
        escola = row[COLUNA_ESCOLA]
        mapa_alunos[nome_norm] = escola

    # 3. Preparar pastas
    if not os.path.exists(PASTA_DESTINO):
        os.makedirs(PASTA_DESTINO)

    arquivos = [f for f in os.listdir(PASTA_ORIGEM) if f.lower().endswith('.pdf')]
    print(f"Total de arquivos na pasta: {len(arquivos)}")

    sucessos = 0
    ignorados = 0
    nao_encontrados = 0

    # Padrão REGEX para identificar (1), (20), etc no final do nome
    padrao_duplicata = re.compile(r'\(\d+\)$') 

    for arquivo in arquivos:
        nome_arquivo_sem_ext = os.path.splitext(arquivo)[0] # Remove .pdf

        # --- NOVA LÓGICA DE FILTRO ---
        # Verifica se termina com (numero). Ex: "NOME(1)"
        if padrao_duplicata.search(nome_arquivo_sem_ext):
            print(f"[IGNORADO] Duplicata detectada: {arquivo}")
            ignorados += 1
            continue # Pula para o próximo arquivo e não faz nada com este

        # Se passou pelo filtro, assume que o nome do arquivo é o nome do aluno
        nome_para_busca = nome_arquivo_sem_ext
        
        # Opcional: Se ainda houver casos com "Certificado - ", removemos aqui por segurança
        if "Certificado -" in nome_para_busca:
             nome_para_busca = nome_para_busca.replace("Certificado -", "")

        nome_arquivo_norm = normalizar_texto(nome_para_busca)
        escola_destino = mapa_alunos.get(nome_arquivo_norm)

        if escola_destino:
            caminho_escola = os.path.join(PASTA_DESTINO, str(escola_destino).strip())
            
            if not os.path.exists(caminho_escola):
                os.makedirs(caminho_escola)
            
            shutil.copy2(os.path.join(PASTA_ORIGEM, arquivo), os.path.join(caminho_escola, arquivo))
            sucessos += 1
            print(f"[OK] {arquivo} -> {escola_destino}")
        else:
            # Move para _NAO_IDENTIFICADOS
            pasta_erro = os.path.join(PASTA_DESTINO, "_NAO_IDENTIFICADOS")
            if not os.path.exists(pasta_erro):
                os.makedirs(pasta_erro)
            
            shutil.copy2(os.path.join(PASTA_ORIGEM, arquivo), os.path.join(pasta_erro, arquivo))
            nao_encontrados += 1
            print(f"[ERRO] Nome não bateu com planilha: {nome_para_busca}")

    print("-" * 30)
    print("Resumo:")
    print(f"Arquivos organizados: {sucessos}")
    print(f"Duplicatas ignoradas: {ignorados}")
    print(f"Não identificados: {nao_encontrados}")

if __name__ == "__main__":
    organizar_arquivos()
