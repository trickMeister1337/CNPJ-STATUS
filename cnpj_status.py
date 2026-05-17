"""
CNPJ-STATUS — Consulta situação cadastral de CNPJs via OpenCNPJ API.

Uso:
    python cnpj_status.py planilha.csv
    python cnpj_status.py planilha.xlsx

Saída:
    resultado_cnpjs.xlsx        — planilha original + colunas enriquecidas
    relatorio_cnpj_status.html  — relatório Big4 com quadro geral e empresas não-ativas

Política de limites OpenCNPJ:
    A API não exige chave e aceita picos pontuais acima de 100 req/s.
    O bloqueio (429) ocorre quando UMA MESMA ORIGEM mantém volume contínuo
    por período prolongado. Para 10k CNPJs, processamos em lotes com pausas
    entre eles para quebrar o padrão de varredura contínua.
"""

import sys
import time
import requests
import pandas as pd
from report_generator import gerar_relatorio


API_URL = "https://kitana.opencnpj.com/cnpj/{cnpj}"
TIMEOUT_SEGUNDOS = 10

# --- Controle de cadência ---
DELAY_ENTRE_REQUESTS = 0.3
TAMANHO_LOTE = 50
PAUSA_ENTRE_LOTES = 45           # segundos

# Backoff para 429: esperas longas porque o bloqueio só cede ao normalizar o tráfego
RETRY_DELAYS = [60, 120, 180]    # segundos — tentativas 1, 2, 3

MAPA_SITUACAO = {
    "ativa":    "Ativa",
    "baixada":  "Baixada",
    "suspensa": "Inativa",
    "inapta":   "Inativa",
    "nula":     "Inativa",
}


# ── Helpers ──────────────────────────────────────────────────────────────────

def _get(data, *keys):
    """Extrai o primeiro valor não-vazio dentre as chaves tentadas."""
    for k in keys:
        v = data.get(k)
        if v and str(v).strip() not in ("", "None", "null"):
            return str(v).strip()
    return ""


def normalizar_cnpj(valor):
    """Remove pontuação e preenche com zeros à esquerda até 14 dígitos."""
    cnpj = "".join(c for c in str(valor) if c.isdigit())
    return cnpj.zfill(14)


def mapear_situacao(situacao):
    """Mapeia o valor de situacaoCadastral para o status normalizado."""
    chave = situacao.strip().lower()
    if chave in MAPA_SITUACAO:
        return MAPA_SITUACAO[chave]
    return f"Status Desconhecido: {situacao}"


def _resultado_erro(msg):
    """Retorna dicionário padronizado para casos de erro na requisição."""
    return {
        "status":                    f"Erro na consulta: {msg}",
        "nome_fantasia":             "",
        "razao_social":              "",
        "data_situacao_cadastral":   "",
        "motivo_situacao_cadastral": "",
    }


# ── Requisição e retry ────────────────────────────────────────────────────────

class _RateLimitError(Exception):
    """Sinaliza que a API retornou 429 (bloqueio temporário por volume contínuo)."""
    def __init__(self, retry_after=None):
        self.retry_after = retry_after


def _executar_requisicao(cnpj):
    """
    Faz a requisição HTTP e retorna um dicionário com todos os campos relevantes.
    Lança _RateLimitError em caso de 429.
    """
    url = API_URL.format(cnpj=cnpj)
    resposta = requests.get(url, timeout=TIMEOUT_SEGUNDOS)

    if resposta.status_code == 429:
        raise _RateLimitError(resposta.headers.get("Retry-After"))

    if resposta.status_code == 404:
        return _resultado_erro("CNPJ não encontrado (404)")

    if not resposta.ok:
        return _resultado_erro(f"HTTP {resposta.status_code}")

    dados = resposta.json()
    data  = dados.get("data", {})

    situacao = _get(data, "situacaoCadastral", "situacao_cadastral", "situacao")
    if not situacao:
        return _resultado_erro("campo situacaoCadastral ausente na resposta")

    return {
        "status": mapear_situacao(situacao),
        # Tenta variações de nome usadas por diferentes versões da API
        "nome_fantasia": _get(
            data, "nomeFantasia", "nome_fantasia", "fantasia",
        ),
        "razao_social": _get(
            data, "razaoSocial", "razao_social", "nome",
        ),
        "data_situacao_cadastral": _get(
            data,
            "dataSituacaoCadastral", "data_situacao_cadastral",
            "dataSituacao", "data_situacao",
        ),
        "motivo_situacao_cadastral": _get(
            data,
            "motivoSituacaoCadastral", "motivo_situacao_cadastral",
            "motivoSituacao", "descricaoSituacaoCadastral",
            "descricao_situacao_cadastral",
        ),
    }


def consultar_cnpj(cnpj):
    """
    Consulta com até 3 retentativas em caso de 429.
    As esperas são longas (60/120/180s) porque o bloqueio só cede quando
    o tráfego da origem normaliza — esperas curtas re-triggeram o bloqueio.
    """
    ultimo_erro = None

    for tentativa, pausa in enumerate(RETRY_DELAYS, start=1):
        try:
            return _executar_requisicao(cnpj)

        except _RateLimitError as e:
            espera = int(e.retry_after) if e.retry_after else pausa
            print(f"      [429] tentativa {tentativa}/3 — aguardando {espera}s para tráfego normalizar...")
            time.sleep(espera)
            ultimo_erro = _resultado_erro("rate limit (429) após 3 tentativas")

        except requests.exceptions.Timeout:
            return _resultado_erro("timeout")

        except requests.exceptions.ConnectionError:
            return _resultado_erro("falha de conexão")

        except requests.exceptions.RequestException as e:
            return _resultado_erro(str(e))

        except (ValueError, KeyError) as e:
            return _resultado_erro(f"resposta inesperada ({e})")

    return ultimo_erro


# ── I/O ───────────────────────────────────────────────────────────────────────

def carregar_planilha(caminho):
    """Carrega CSV ou XLSX e retorna um DataFrame."""
    if caminho.lower().endswith(".csv"):
        return pd.read_csv(caminho, dtype=str)
    elif caminho.lower().endswith((".xlsx", ".xls")):
        return pd.read_excel(caminho, dtype=str)
    else:
        raise ValueError(f"Formato não suportado: {caminho}. Use .csv ou .xlsx")


def estimar_tempo(total):
    """Retorna string com estimativa de duração total para o volume informado."""
    num_lotes = (total + TAMANHO_LOTE - 1) // TAMANHO_LOTE
    segundos  = (total * DELAY_ENTRE_REQUESTS) + ((num_lotes - 1) * PAUSA_ENTRE_LOTES)
    horas, resto = divmod(int(segundos), 3600)
    minutos = resto // 60
    if horas:
        return f"~{horas}h {minutos}min"
    return f"~{minutos}min"


# ── Main ──────────────────────────────────────────────────────────────────────

def main(caminho_entrada):
    print(f"[*] Carregando planilha: {caminho_entrada}")
    df = carregar_planilha(caminho_entrada)

    df.columns = [c.strip().lower() for c in df.columns]
    if "cnpj" not in df.columns:
        raise ValueError("Coluna 'cnpj' não encontrada na planilha.")

    total     = len(df)
    largura   = len(str(total))
    estimativa = estimar_tempo(total)

    print(f"[*] {total} CNPJs encontrados.")
    print(f"[*] Cadência: {DELAY_ENTRE_REQUESTS}s entre requests, "
          f"pausa de {PAUSA_ENTRE_LOTES}s a cada {TAMANHO_LOTE} consultas.")
    print(f"[*] Tempo estimado: {estimativa}\n")

    resultados = []

    for i, valor in enumerate(df["cnpj"], start=1):
        cnpj   = normalizar_cnpj(valor)
        result = consultar_cnpj(cnpj)
        resultados.append(result)

        icone = "✓" if not result["status"].startswith("Erro") else "✗"
        print(f"  [{i:>{largura}}/{total}] {icone} {cnpj} → {result['status']}")

        if i < total:
            if i % TAMANHO_LOTE == 0:
                lote_atual  = i // TAMANHO_LOTE
                total_lotes = (total + TAMANHO_LOTE - 1) // TAMANHO_LOTE
                print(f"\n  [lote {lote_atual}/{total_lotes} concluído] "
                      f"aguardando {PAUSA_ENTRE_LOTES}s...\n")
                time.sleep(PAUSA_ENTRE_LOTES)
            else:
                time.sleep(DELAY_ENTRE_REQUESTS)

    # Enriquece o DataFrame com todos os campos retornados pela API
    df["status"]                    = [r["status"]                    for r in resultados]
    df["nome_fantasia"]             = [r["nome_fantasia"]             for r in resultados]
    df["razao_social"]              = [r["razao_social"]              for r in resultados]
    df["data_situacao_cadastral"]   = [r["data_situacao_cadastral"]   for r in resultados]
    df["motivo_situacao_cadastral"] = [r["motivo_situacao_cadastral"] for r in resultados]

    # Salva planilha enriquecida
    saida_xlsx = "resultado_cnpjs.xlsx"
    df.to_excel(saida_xlsx, index=False)
    print(f"\n[+] Planilha salva em: {saida_xlsx}")

    # Gera relatório HTML no padrão Big4
    saida_html = gerar_relatorio(df, caminho_entrada)
    print(f"[+] Relatório salvo em: {saida_html}")

    # Resumo no terminal
    contagem = df["status"].value_counts()
    print("\n--- Resumo ---")
    for status, qtd in contagem.items():
        print(f"  {status}: {qtd}")


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Uso: python cnpj_status.py <arquivo.csv|arquivo.xlsx>")
        sys.exit(1)

    main(sys.argv[1])
