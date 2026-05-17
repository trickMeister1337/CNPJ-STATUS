"""
CNPJ-STATUS — Consulta situação cadastral de CNPJs via OpenCNPJ API.

Uso:
    python cnpj_status.py planilha.csv
    python cnpj_status.py planilha.xlsx

Saída: resultado_cnpjs.xlsx (mesmas colunas + coluna 'status')

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


API_URL = "https://kitana.opencnpj.com/cnpj/{cnpj}"
TIMEOUT_SEGUNDOS = 10

# --- Controle de cadência ---
# Delay curto entre requisições dentro de um lote (a API tolera bursts pontuais)
DELAY_ENTRE_REQUESTS = 0.3       # segundos

# A cada N requisições, pausa mais longa para quebrar padrão de varredura contínua
TAMANHO_LOTE = 50
PAUSA_ENTRE_LOTES = 45           # segundos

# Backoff para 429: esperas longas porque o bloqueio só cede ao normalizar o tráfego
RETRY_DELAYS = [60, 120, 180]    # segundos — tentativas 1, 2, 3

# Mapeamento de situação cadastral para status normalizado
MAPA_SITUACAO = {
    "ativa":    "Ativa",
    "baixada":  "Baixada",
    "suspensa": "Inativa",
    "inapta":   "Inativa",
    "nula":     "Inativa",
}


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


class _RateLimitError(Exception):
    """Sinaliza que a API retornou 429 (bloqueio temporário por volume contínuo)."""
    def __init__(self, retry_after=None):
        self.retry_after = retry_after


def _executar_requisicao(cnpj):
    """
    Faz a requisição HTTP e retorna o status mapeado.
    Lança _RateLimitError se a API retornar 429.
    """
    url = API_URL.format(cnpj=cnpj)
    resposta = requests.get(url, timeout=TIMEOUT_SEGUNDOS)

    if resposta.status_code == 429:
        retry_after = resposta.headers.get("Retry-After")
        raise _RateLimitError(retry_after)

    if resposta.status_code == 404:
        return "Erro na consulta: CNPJ não encontrado (404)"

    if not resposta.ok:
        return f"Erro na consulta: HTTP {resposta.status_code}"

    dados = resposta.json()
    situacao = dados.get("data", {}).get("situacaoCadastral")
    if situacao is None:
        return "Erro na consulta: campo situacaoCadastral ausente na resposta"

    return mapear_situacao(situacao)


def consultar_cnpj(cnpj):
    """
    Consulta com até 3 retentativas em caso de 429.
    As esperas são longas (60/120/180s) porque o bloqueio só cede quando
    o tráfego da origem normaliza — esperas curtas apenas re-triggeram o bloqueio.
    """
    ultimo_erro = None

    for tentativa, pausa in enumerate(RETRY_DELAYS, start=1):
        try:
            return _executar_requisicao(cnpj)

        except _RateLimitError as e:
            # Prioriza o Retry-After da API; se ausente, usa backoff progressivo
            espera = int(e.retry_after) if e.retry_after else pausa
            print(f"      [429] tentativa {tentativa}/3 — aguardando {espera}s para tráfego normalizar...")
            time.sleep(espera)
            ultimo_erro = "Erro na consulta: rate limit (429) após 3 tentativas"

        except requests.exceptions.Timeout:
            return "Erro na consulta: timeout"

        except requests.exceptions.ConnectionError:
            return "Erro na consulta: falha de conexão"

        except requests.exceptions.RequestException as e:
            return f"Erro na consulta: {e}"

        except (ValueError, KeyError) as e:
            return f"Erro na consulta: resposta inesperada ({e})"

    return ultimo_erro


def estimar_tempo(total):
    """Retorna string com estimativa de duração total para o volume informado."""
    num_lotes = (total + TAMANHO_LOTE - 1) // TAMANHO_LOTE
    segundos = (total * DELAY_ENTRE_REQUESTS) + ((num_lotes - 1) * PAUSA_ENTRE_LOTES)
    horas, resto = divmod(int(segundos), 3600)
    minutos = resto // 60
    if horas:
        return f"~{horas}h {minutos}min"
    return f"~{minutos}min"


def carregar_planilha(caminho):
    """Carrega CSV ou XLSX e retorna um DataFrame."""
    if caminho.lower().endswith(".csv"):
        return pd.read_csv(caminho, dtype=str)
    elif caminho.lower().endswith((".xlsx", ".xls")):
        return pd.read_excel(caminho, dtype=str)
    else:
        raise ValueError(f"Formato não suportado: {caminho}. Use .csv ou .xlsx")


def main(caminho_entrada):
    print(f"[*] Carregando planilha: {caminho_entrada}")
    df = carregar_planilha(caminho_entrada)

    if "cnpj" not in df.columns:
        raise ValueError("Coluna 'cnpj' não encontrada na planilha.")

    total = len(df)
    largura = len(str(total))
    estimativa = estimar_tempo(total)

    print(f"[*] {total} CNPJs encontrados.")
    print(f"[*] Cadência: {DELAY_ENTRE_REQUESTS}s entre requests, "
          f"pausa de {PAUSA_ENTRE_LOTES}s a cada {TAMANHO_LOTE} consultas.")
    print(f"[*] Tempo estimado: {estimativa}\n")

    resultados = []

    for i, valor in enumerate(df["cnpj"], start=1):
        cnpj = normalizar_cnpj(valor)
        status = consultar_cnpj(cnpj)
        resultados.append(status)

        icone = "✓" if not status.startswith("Erro") else "✗"
        print(f"  [{i:>{largura}}/{total}] {icone} {cnpj} → {status}")

        # Pausa longa ao completar um lote (exceto no último item)
        if i < total:
            if i % TAMANHO_LOTE == 0:
                lote_atual = i // TAMANHO_LOTE
                total_lotes = (total + TAMANHO_LOTE - 1) // TAMANHO_LOTE
                print(f"\n  [lote {lote_atual}/{total_lotes} concluído] "
                      f"aguardando {PAUSA_ENTRE_LOTES}s...\n")
                time.sleep(PAUSA_ENTRE_LOTES)
            else:
                time.sleep(DELAY_ENTRE_REQUESTS)

    df["status"] = resultados

    saida = "resultado_cnpjs.xlsx"
    df.to_excel(saida, index=False)
    print(f"\n[+] Concluído. Resultado salvo em: {saida}")

    contagem = df["status"].value_counts()
    print("\n--- Resumo ---")
    for status, qtd in contagem.items():
        print(f"  {status}: {qtd}")


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Uso: python cnpj_status.py <arquivo.csv|arquivo.xlsx>")
        sys.exit(1)

    main(sys.argv[1])
