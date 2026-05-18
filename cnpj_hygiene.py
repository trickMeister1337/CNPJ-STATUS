#!/usr/bin/env python3
"""
CNPJ-HYGIENE — Higienização de planilhas antes da consulta via OpenCNPJ.

Uso:
    python cnpj_hygiene.py planilha.xlsx
    python cnpj_hygiene.py planilha.csv

Saída:
    cnpjs_validos.xlsx    — CNPJs prontos para rodar com cnpj_status.py
    cnpjs_revisao.xlsx    — Planilha multi-aba para conferência manual
                            (abas: Filiais | CPF | Inconsistentes)
"""

import re
import sys
import pandas as pd
from collections import Counter


# ── Classificações possíveis ──────────────────────────────────────────────────

VALIDO       = "VALIDO"
FILIAL       = "FILIAL"
CPF          = "CPF"
INCOMPLETO   = "INCOMPLETO"
ALFANUMERICO = "ALFANUMERICO"
INVALIDO     = "INVALIDO"   # dígitos verificadores incorretos
VAZIO        = "VAZIO"


# ── Algoritmo de validação dos dígitos verificadores ─────────────────────────

def _calcular_digito(sequencia, pesos):
    soma = sum(int(d) * p for d, p in zip(sequencia, pesos))
    resto = soma % 11
    return 0 if resto < 2 else 11 - resto


def validar_cnpj(digitos):
    """Retorna True se os dois dígitos verificadores do CNPJ de 14 dígitos forem válidos."""
    if len(digitos) != 14 or not digitos.isdigit():
        return False
    # CNPJs com todos os dígitos iguais são inválidos (ex: 00000000000000)
    if len(set(digitos)) == 1:
        return False
    pesos1 = [5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2]
    pesos2 = [6, 5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2]
    d1 = _calcular_digito(digitos[:12], pesos1)
    d2 = _calcular_digito(digitos[:13], pesos2)
    return digitos[12] == str(d1) and digitos[13] == str(d2)


# ── Formatação ────────────────────────────────────────────────────────────────

def formatar_cnpj(digitos):
    c = str(digitos).zfill(14)
    return f"{c[:2]}.{c[2:5]}.{c[5:8]}/{c[8:12]}-{c[12:14]}"


def formatar_cpf(digitos):
    c = str(digitos).zfill(11)
    return f"{c[:3]}.{c[3:6]}.{c[6:9]}-{c[9:11]}"


def formatar_raiz(digitos):
    c = str(digitos[:8]).zfill(8)
    return f"{c[:2]}.{c[2:5]}.{c[5:8]}"


# ── Classificador principal ───────────────────────────────────────────────────

def classificar(cnpj_raw):
    """
    Recebe o valor bruto da célula e retorna (classificacao, digitos, observacao).

    classificacao : uma das constantes definidas acima
    digitos       : string somente dígitos (melhor esforço)
    observacao    : texto legível explicando o diagnóstico
    """
    valor = str(cnpj_raw).strip()

    if not valor or valor.lower() in ("nan", "none", ""):
        return VAZIO, "", "Campo vazio"

    # Verifica presença de letras (possível novo formato alfanumérico 2026)
    tem_letras = bool(re.search(r'[A-Za-z]', valor))

    # Extrai apenas os dígitos
    digitos = re.sub(r'\D', '', valor)
    n = len(digitos)

    if n == 0:
        return VAZIO, digitos, "Nenhum dígito encontrado"

    if tem_letras:
        return ALFANUMERICO, digitos, (
            "Contém letras — provável CNPJ alfanumérico (novo formato 2026). "
            "Verificação manual necessária"
        )

    if n == 11:
        return CPF, digitos, f"Número com 11 dígitos — é um CPF ({formatar_cpf(digitos)})"

    if n < 11:
        return INCOMPLETO, digitos, f"Apenas {n} dígito(s) — número muito curto para CNPJ ou CPF"

    if n > 14:
        return INCOMPLETO, digitos[:14], f"{n} dígitos — excede o tamanho de CNPJ (14). Truncado para verificação"

    if n < 14:
        return INCOMPLETO, digitos, f"{n} dígitos — CNPJ incompleto (esperado: 14)"

    # n == 14 — verifica dígitos verificadores
    if not validar_cnpj(digitos):
        return INVALIDO, digitos, (
            f"Dígitos verificadores incorretos em {formatar_cnpj(digitos)}. "
            "CNPJ pode estar digitado errado ou ser inválido"
        )

    # CNPJ válido — verifica se é filial
    estabelecimento = digitos[8:12]
    if estabelecimento != "0001":
        raiz = formatar_raiz(digitos)
        return FILIAL, digitos, (
            f"Estabelecimento filial (/{estabelecimento}). "
            f"Raiz do CNPJ: {raiz} — consulte o estabelecimento /0001 desta raiz"
        )

    return VALIDO, digitos, ""


# ── I/O ───────────────────────────────────────────────────────────────────────

def carregar_planilha(caminho):
    if caminho.lower().endswith(".csv"):
        return pd.read_csv(caminho, dtype=str)
    elif caminho.lower().endswith((".xlsx", ".xls")):
        return pd.read_excel(caminho, dtype=str)
    raise ValueError(f"Formato não suportado: {caminho}")


# ── Main ──────────────────────────────────────────────────────────────────────

def main(caminho):
    print(f"[*] Carregando: {caminho}")
    df = carregar_planilha(caminho)
    df.columns = [c.strip().lower() for c in df.columns]

    if "cnpj" not in df.columns:
        raise ValueError("Coluna 'cnpj' não encontrada na planilha.")

    total = len(df)
    print(f"[*] {total} registros encontrados. Analisando...\n")

    # Listas por categoria
    validos       = []
    filiais       = []
    cpfs          = []
    inconsistentes = []   # INCOMPLETO + INVALIDO + ALFANUMERICO + VAZIO

    for _, row in df.iterrows():
        classe, digitos, obs = classificar(row["cnpj"])
        extra = {
            "cnpj_original":    row["cnpj"],
            "cnpj_normalizado": formatar_cnpj(digitos) if len(digitos) == 14 else digitos,
            "classificacao":    classe,
            "observacao":       obs,
        }
        # Preserva colunas extras originais
        for col in df.columns:
            if col != "cnpj":
                extra[col] = row[col]

        if classe == VALIDO:
            validos.append({**{"cnpj": formatar_cnpj(digitos)}, **{c: row[c] for c in df.columns if c != "cnpj"}})

        elif classe == FILIAL:
            raiz_digitos = digitos[:8]
            extra["cnpj_raiz"]           = formatar_raiz(raiz_digitos)
            extra["estabelecimento"]      = digitos[8:12]
            extra["cnpj_filial_fmt"]      = formatar_cnpj(digitos)
            filiais.append(extra)

        elif classe == CPF:
            extra["cpf_formatado"] = formatar_cpf(digitos) if len(digitos) == 11 else digitos
            cpfs.append(extra)

        else:
            inconsistentes.append(extra)

    # ── Salva cnpjs_validos.xlsx ─────────────────────────────────────────────
    df_validos = pd.DataFrame(validos)
    df_validos.to_excel("cnpjs_validos.xlsx", index=False)

    # ── Salva cnpjs_revisao.xlsx (multi-aba) ─────────────────────────────────
    with pd.ExcelWriter("cnpjs_revisao.xlsx", engine="openpyxl") as writer:

        # Aba Filiais
        if filiais:
            cols_filiais = ["cnpj_filial_fmt", "cnpj_raiz", "estabelecimento",
                            "observacao", "cnpj_original"]
            cols_filiais += [c for c in df.columns if c not in ("cnpj",) and c not in cols_filiais]
            pd.DataFrame(filiais).reindex(columns=cols_filiais).to_excel(
                writer, sheet_name="Filiais", index=False
            )
        else:
            pd.DataFrame(columns=["cnpj_filial_fmt","cnpj_raiz","estabelecimento","observacao"]).to_excel(
                writer, sheet_name="Filiais", index=False
            )

        # Aba CPF
        if cpfs:
            cols_cpf = ["cpf_formatado", "cnpj_original", "observacao"]
            cols_cpf += [c for c in df.columns if c not in ("cnpj",) and c not in cols_cpf]
            pd.DataFrame(cpfs).reindex(columns=cols_cpf).to_excel(
                writer, sheet_name="CPF", index=False
            )
        else:
            pd.DataFrame(columns=["cpf_formatado","cnpj_original","observacao"]).to_excel(
                writer, sheet_name="CPF", index=False
            )

        # Aba Inconsistentes
        if inconsistentes:
            cols_inc = ["cnpj_original", "cnpj_normalizado", "classificacao", "observacao"]
            cols_inc += [c for c in df.columns if c not in ("cnpj",) and c not in cols_inc]
            pd.DataFrame(inconsistentes).reindex(columns=cols_inc).to_excel(
                writer, sheet_name="Inconsistentes", index=False
            )
        else:
            pd.DataFrame(columns=["cnpj_original","cnpj_normalizado","classificacao","observacao"]).to_excel(
                writer, sheet_name="Inconsistentes", index=False
            )

    # ── Quadro resumo no terminal ─────────────────────────────────────────────
    n_val  = len(validos)
    n_fil  = len(filiais)
    n_cpf  = len(cpfs)
    n_inc  = len(inconsistentes)
    n_rev  = n_fil + n_cpf + n_inc

    pct = lambda n: f"{n/total*100:.1f}%" if total else "0%"
    bar = lambda n, w=22: "█" * int(n/total*w) + "░" * (w - int(n/total*w)) if total else "░"*w

    # Detalha inconsistentes por sub-tipo
    sub = Counter(r["classificacao"] for r in inconsistentes)

    print()
    print("╔═════════════════════════════════════════════════════════╗")
    print("║        CNPJ-HYGIENE — Resultado da Higienização         ║")
    print("╠═════════════════════════════════════════════════════════╣")
    print(f"║  Total analisado  : {total:>6}                              ║")
    print("╠═════════════════════════════════════════════════════════╣")
    print(f"║  ✓  Válidos (matriz)  : {n_val:>5}  {bar(n_val)}  {pct(n_val):>6}  ║")
    print("╠═════════════════════════════════════════════════════════╣")
    print(f"║  ↳  Filiais           : {n_fil:>5}  {bar(n_fil)}  {pct(n_fil):>6}  ║")
    print(f"║  ✗  CPF               : {n_cpf:>5}  {bar(n_cpf)}  {pct(n_cpf):>6}  ║")
    for classe, qtd in sub.most_common():
        label = {
            INCOMPLETO:   "Incompleto",
            INVALIDO:     "Dígito inválido",
            ALFANUMERICO: "Alfanumérico",
            VAZIO:        "Vazio",
        }.get(classe, classe)
        print(f"║  ✗  {label:<17}: {qtd:>5}  {bar(qtd)}  {pct(qtd):>6}  ║")
    print("╠═════════════════════════════════════════════════════════╣")
    print(f"║  Para revisão manual  : {n_rev:>5}  {bar(n_rev)}  {pct(n_rev):>6}  ║")
    print("╠═════════════════════════════════════════════════════════╣")
    print(f"║  📄 cnpjs_validos.xlsx   → {n_val} CNPJs prontos para consulta  ║")
    print(f"║  📋 cnpjs_revisao.xlsx   → {n_rev} itens para conferência manual ║")
    print("╚═════════════════════════════════════════════════════════╝")
    print()
    print("Próximo passo:")
    print(f"  python cnpj_status.py cnpjs_validos.xlsx")


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Uso: python cnpj_hygiene.py <arquivo.csv|arquivo.xlsx>")
        sys.exit(1)
    main(sys.argv[1])
