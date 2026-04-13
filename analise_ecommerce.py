"""
Agente de análise de e-commerce VTEX
=====================================
Lê a planilha "Cópia de Analise E-commerce.xlsx" e calcula:
  - Total de pedidos únicos
  - Soma dos valores dos pedidos
  - Ticket médio

Para ajustar o mapeamento de colunas, edite as variáveis abaixo:
"""

import os
import re
import sys

import pandas as pd

# ---------------------------------------------------------------------------
# CONFIGURAÇÃO — edite aqui para mapear nomes de colunas da sua planilha
# ---------------------------------------------------------------------------
ARQUIVO_PLANILHA = "Cópia de Analise E-commerce.xlsx"   # caminho relativo ao script

COLUNA_PEDIDO = "Order"         # coluna com o número/id do pedido
COLUNA_VALOR  = "Total Value"   # coluna com o valor total do pedido
# ---------------------------------------------------------------------------


def limpa_valor(valor) -> float:
    """Converte um valor monetário em float.

    Aceita formatos variados, por exemplo:
      - "251.54"          (ponto decimal)
      - "1.234,56"        (milhar com ponto, decimal com vírgula)
      - "R$ 1.234,56"     (com símbolo de moeda)
      - "1234,56"         (só vírgula decimal)
    """
    if pd.isnull(valor):
        return 0.0

    val = str(valor).strip()

    # Remove símbolo de moeda "R$" e espaços em branco
    val = re.sub(r"R\$|\s", "", val)

    if not val:
        return 0.0

    # Detecta formato: se houver vírgula, é decimal brasileiro ("1.234,56")
    if "," in val:
        val = val.replace(".", "")   # remove separadores de milhar
        val = val.replace(",", ".")  # transforma vírgula decimal em ponto
    # Caso contrário já usa ponto como decimal — não faz nada

    try:
        return float(val)
    except ValueError:
        return 0.0


def main():
    # Localiza a planilha em relação ao diretório do script
    base_dir = os.path.dirname(os.path.abspath(__file__))
    caminho = os.path.join(base_dir, ARQUIVO_PLANILHA)

    if not os.path.exists(caminho):
        print(f"[ERRO] Planilha não encontrada: {caminho}", file=sys.stderr)
        sys.exit(1)

    print(f"Lendo planilha: {caminho}")

    # Lê tudo como texto para não perder precisão nem ter problemas de formato
    extensao = os.path.splitext(caminho)[1].lower()
    if extensao == ".csv":
        df = pd.read_csv(caminho, dtype=str, encoding="utf-8")
    elif extensao in (".xlsx", ".xls"):
        df = pd.read_excel(caminho, dtype=object, engine="openpyxl")
    else:
        print(f"[ERRO] Formato de arquivo não suportado: {extensao}", file=sys.stderr)
        sys.exit(1)

    # Valida colunas
    colunas_faltando = [c for c in (COLUNA_PEDIDO, COLUNA_VALOR) if c not in df.columns]
    if colunas_faltando:
        print(
            f"[ERRO] Coluna(s) não encontrada(s) na planilha: {colunas_faltando}\n"
            f"       Colunas disponíveis: {list(df.columns)}",
            file=sys.stderr,
        )
        sys.exit(1)

    # Remove duplicatas — considera apenas 1 linha por pedido
    df_unicos = df.drop_duplicates(subset=COLUNA_PEDIDO).copy()

    # Limpa e converte a coluna de valor
    df_unicos["_valor_num"] = df_unicos[COLUNA_VALOR].apply(limpa_valor)

    # Calcula métricas
    total_pedidos = len(df_unicos)
    soma_valores  = df_unicos["_valor_num"].sum()
    ticket_medio  = soma_valores / total_pedidos if total_pedidos > 0 else 0.0

    # Imprime resultado
    print()
    print("=" * 40)
    print("  RESULTADO — Análise E-commerce VTEX")
    print("=" * 40)
    print(f"  Total de pedidos únicos : {total_pedidos:,}")
    print(f"  Soma dos valores        : R$ {soma_valores:,.2f}")
    print(f"  Ticket médio            : R$ {ticket_medio:,.2f}")
    print("=" * 40)


if __name__ == "__main__":
    main()
