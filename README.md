# 📂 Organizador de Certificados por Escola

Este script em Python automatiza a organização de certificados escolares. Ele lê uma pasta cheia de PDFs misturados, cruza as informações com uma planilha de alunos e move cada arquivo para uma pasta específica com o nome da Escola (Estabelecimento).

## ✨ O que o script faz

1.  **Lê** uma planilha Excel contendo nomes de alunos e suas respectivas escolas.
2.  **Verifica** uma pasta de origem contendo os certificados (PDFs).
3.  **Ignora** arquivos duplicados gcontidos no Drive (ex: `JOAO(1).pdf`).
4.  **Move** os arquivos originais para pastas organizadas por nome da escola.
5.  **Separa** arquivos que não foram encontrados na planilha em uma pasta `_NAO_IDENTIFICADOS`.

-----

## 🛠️ Pré-requisitos

Você precisa ter o **Python** instalado no computador. Além disso, é necessário instalar algumas bibliotecas que o script utiliza para ler Excel e tratar textos.

Abra seu terminal (CMD ou PowerShell) e rode o comando abaixo:

```bash
pip install pandas openpyxl unidecode
```

-----

## 📋 Como preparar a Planilha (Excel)

Para que o script funcione, sua planilha deve estar formatada corretamente. O script espera duas colunas principais (o nome pode ser ajustado no código, mas o padrão atual é):

  * **ALUNO**: Coluna com o nome completo do aluno.
  * **ESTABELECIMENTO**: Coluna com o nome da escola.

**Exemplo visual da planilha:**

| ALUNO | ESTABELECIMENTO |
| :--- | :--- |
| João da Silva | C.E. Machado de Assis |
| Maria Oliveira | C.E. Monteiro Lobato |
| Willian Maciel | C.E. Machado de Assis |

> **Nota:** Não se preocupe com acentos ou letras maiúsculas/minúsculas. O script trata isso automaticamente (ex: "joão" será igual a "JOAO").

-----

## ⚙️ Configuração do Script

Antes de rodar, abra o arquivo `.py` em um editor de texto (Bloco de Notas, VS Code, etc.) e edite a seção **CONFIGURAÇÕES** no topo do arquivo:

```python
# --- CONFIGURAÇÕES ---
PASTA_ORIGEM = r"C:\Users\Voce\Downloads\Certificados_Misturados"
PASTA_DESTINO = r"C:\Users\Voce\Documents\Certificados_Organizados"
ARQUIVO_PLANILHA = r"C:\Users\Voce\Documents\Lista_Alunos.xlsx"

# Nomes exatos das colunas na sua planilha
COLUNA_NOME = "ALUNO"
COLUNA_ESCOLA = "ESTABELECIMENTO"
```

  * **Dica:** O `r` antes das aspas é importante no Windows para evitar erros com as barras `\`.

-----

## 🚀 Como Executar

1.  Certifique-se de que a planilha está fechada (o Excel bloqueia a leitura se estiver aberto).
2.  Abra o terminal na pasta onde está o script.
3.  Execute o comando:

<!-- end list -->

```bash
python organizar_certificados.py
```

Aguarde o processamento. O terminal mostrará quais arquivos foram movidos com sucesso (`[OK]`), quais foram ignorados (`[IGNORADO]`) e quais deram erro.



  * Ele criará uma pasta chamada `_NAO_IDENTIFICADOS` dentro do destino.
  * Todos os arquivos "sobras" serão copiados para lá.
  * **Ação recomendada:** Abra essa pasta e organize esses poucos arquivos manualmente.
