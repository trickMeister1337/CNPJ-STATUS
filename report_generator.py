#!/usr/bin/env python3
"""
CNPJ-STATUS — Gerador de Relatório (Big4 Style).

Uso direto: chamado automaticamente pelo cnpj_status.py ao final da execução.
"""

import html as H
import os
from datetime import datetime

esc = lambda s: H.escape(str(s)) if s else ""

# Cores por status (mesmo palette do SWARM RED)
_CORES = {
    "ativa":   "#6e8f72",
    "baixada": "#4a7c8c",
    "inativa": "#d4833a",
    "erro":    "#b34e4e",
    "outro":   "#95a5a6",
}

_CSS = """
body{font-family:'Segoe UI',Arial,sans-serif;margin:0;padding:20px;background:#f0f2f5}
.ctn{max-width:1200px;margin:0 auto;background:#fff;border-radius:10px;overflow:hidden;box-shadow:0 2px 10px rgba(0,0,0,.1)}
.hdr{background:#1a3a4f;color:#fff;padding:40px 30px;text-align:center}
.hdr h1{margin:0 0 5px;font-size:1.6em;letter-spacing:2px;text-transform:uppercase}
.hdr .sub{font-size:1.1em;opacity:.9;margin:5px 0}
.hdr .meta{font-size:.85em;opacity:.7}
.hdr .cls{display:inline-block;border:2px solid #e74c3c;color:#e74c3c;padding:4px 16px;border-radius:4px;font-weight:700;font-size:.8em;margin-top:12px;letter-spacing:1px}
.cnt{padding:30px}
h2{color:#1a3a4f;border-bottom:2px solid #e0e0e0;padding-bottom:8px;margin-top:30px}
h3{color:#2c3e50;margin-top:20px}
.sts{display:flex;gap:12px;margin:20px 0;flex-wrap:wrap}
.sc{flex:1;padding:18px;text-align:center;color:#fff;border-radius:8px;min-width:85px}
.sc .n{font-size:32px;font-weight:bold}
.sc .l{font-size:.75em;text-transform:uppercase;letter-spacing:.5px;opacity:.9}
.sc .p{font-size:.85em;opacity:.85;margin-top:3px}
.s-te{background:#2c3e50}
.s-at{background:#6e8f72}
.s-bx{background:#4a7c8c}
.s-in{background:#d4833a}
.s-er{background:#b34e4e}
.ib{background:#e8f4f8;padding:15px 20px;border-radius:8px;margin:15px 0;border-left:4px solid #1a3a4f}
.ib.n{border-left-color:#2c3e50;background:#f9f9f9}
.ib.g{border-left-color:#27ae60;background:#f0faf4}
.ib.w{border-left-color:#d4833a;background:#fef9f0}
table{width:100%;border-collapse:collapse;margin:10px 0}
th,td{border:1px solid #ddd;padding:10px;text-align:left;vertical-align:middle}
th{background:#f5f5f5;font-weight:600;font-size:.85em;color:#2c3e50}
tr:nth-child(even) td{background:#fafafa}
.sb{display:inline-block;padding:3px 10px;border-radius:4px;font-size:.75em;font-weight:700;color:#fff;white-space:nowrap}
.rb{background:#e0e0e0;border-radius:4px;height:12px;margin:5px 0;min-width:60px}
.ri{height:12px;border-radius:4px;transition:width .3s}
.ft{background:#f5f5f5;padding:20px;text-align:center;font-size:.8em;color:#666}
.toc{background:#f8f9fa;padding:15px 20px;border-radius:8px;margin:15px 0}
.toc a{color:#1a3a4f;text-decoration:none}
.toc a:hover{text-decoration:underline}
.toc li{margin:4px 0}
code{background:#f4f4f4;padding:1px 5px;border-radius:3px;font-size:.85em;font-family:'Cascadia Code',monospace}
.filtro-wrap{display:flex;align-items:center;gap:10px;margin-bottom:12px}
.filtro-wrap input{padding:8px 12px;border:1px solid #ddd;border-radius:6px;width:320px;font-size:.9em;outline:none}
.filtro-wrap input:focus{border-color:#1a3a4f}
.cnt-badge{background:#1a3a4f;color:#fff;padding:2px 10px;border-radius:10px;font-size:.75em;font-weight:700}
.dist-table td:nth-child(4){min-width:160px}
"""

_JS = """
<script>
function filtrarTabela() {
    var q = document.getElementById('filtro').value.toLowerCase();
    var rows = document.querySelectorAll('#tbl-nao-ativas tbody tr');
    var vis = 0;
    rows.forEach(function(r) {
        var show = r.textContent.toLowerCase().includes(q);
        r.style.display = show ? '' : 'none';
        if (show) vis++;
    });
    document.getElementById('cnt-vis').textContent = vis;
}
</script>
"""


def _cor_status(status):
    s = str(status).lower()
    if s == "ativa":         return _CORES["ativa"]
    if s == "baixada":       return _CORES["baixada"]
    if s == "inativa":       return _CORES["inativa"]
    if s.startswith("erro"): return _CORES["erro"]
    return _CORES["outro"]


def _formatar_cnpj(cnpj):
    # Remove pontuação antes de formatar — o valor pode já vir formatado do Excel
    c = "".join(ch for ch in str(cnpj) if ch.isdigit()).zfill(14)
    return f"{c[:2]}.{c[2:5]}.{c[5:8]}/{c[8:12]}-{c[12:14]}"


def _formatar_data(valor):
    """Converte YYYY-MM-DD ou YYYYMMDD para DD/MM/YYYY; passa outros formatos sem alteração."""
    s = str(valor).strip() if valor else ""
    if not s or s == "nan":
        return "—"
    if len(s) == 10 and s[4] == "-":
        try:
            return datetime.strptime(s, "%Y-%m-%d").strftime("%d/%m/%Y")
        except ValueError:
            pass
    if len(s) == 8 and s.isdigit():
        try:
            return datetime.strptime(s, "%Y%m%d").strftime("%d/%m/%Y")
        except ValueError:
            pass
    return s


def _nome_display(row):
    """Retorna nome fantasia se preenchido, caso contrário razão social."""
    nf = str(row.get("nome_fantasia", "")).strip()
    rs = str(row.get("razao_social", "")).strip()
    if nf and nf.lower() not in ("nan", "none", ""):
        return nf
    if rs and rs.lower() not in ("nan", "none", ""):
        return rs
    return "—"


def _sort_key_status(status):
    s = str(status).lower()
    if s == "baixada":         return 0
    if s == "inativa":         return 1
    if s.startswith("status"): return 2
    if s.startswith("erro"):   return 3
    return 4


def gerar_relatorio(df, nome_arquivo):
    """
    Gera o relatório HTML no padrão Big4.

    Parâmetros:
        df            — DataFrame com colunas: cnpj, status, nome_fantasia,
                        razao_social, data_situacao_cadastral, motivo_situacao_cadastral
        nome_arquivo  — caminho da planilha de entrada (usado no cabeçalho)

    Retorna o caminho do arquivo HTML gerado.
    """
    now = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    total = len(df)

    contagem = df["status"].value_counts().to_dict()
    ativas   = contagem.get("Ativa", 0)
    baixadas = contagem.get("Baixada", 0)
    inativas = contagem.get("Inativa", 0)
    erros    = sum(v for k, v in contagem.items()
                   if k.startswith("Erro") or k.startswith("Status Desconhecido"))

    def pct(n):
        return f"{n / total * 100:.1f}" if total > 0 else "0.0"

    # Filtra não-ativas e ordena: Baixada → Inativa → Desconhecido → Erro
    df_nao = df[df["status"] != "Ativa"].copy()
    df_nao = df_nao.sort_values(
        by=["status", "cnpj"],
        key=lambda col: col.map(_sort_key_status) if col.name == "status" else col,
    )
    n_nao = len(df_nao)

    # ── HTML ──────────────────────────────────────────────────────────────────
    nome_base = os.path.basename(nome_arquivo)

    h = f"""<!DOCTYPE html>
<html lang="pt-br">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>CNPJ-STATUS — Situação Cadastral</title>
<style>{_CSS}</style>
</head>
<body>
<div class="ctn">

<div class="hdr">
  <h1>CNPJ-STATUS — Situação Cadastral</h1>
  <div class="sub">Arquivo: {esc(nome_base)}</div>
  <div class="meta">Data: {now} &nbsp;|&nbsp; Total consultado: {total:,} CNPJs</div>
  <div class="cls">CONFIDENCIAL — DISTRIBUIÇÃO RESTRITA</div>
</div>

<div class="cnt">

<div class="toc">
  <strong>Índice</strong>
  <ol>
    <li><a href="#s1">Quadro Geral</a></li>
    <li><a href="#s2">Empresas Não-Ativas ({n_nao:,})</a></li>
  </ol>
</div>

"""

    # ── Seção 1: Quadro Geral ────────────────────────────────────────────────
    h += '<h2 id="s1">1. Quadro Geral</h2>\n'

    h += '<div class="sts">\n'
    h += f'<div class="sc s-te"><div class="n">{total:,}</div><div class="l">Total</div></div>\n'
    h += f'<div class="sc s-at"><div class="n">{ativas:,}</div><div class="l">Ativa</div><div class="p">{pct(ativas)}%</div></div>\n'
    h += f'<div class="sc s-bx"><div class="n">{baixadas:,}</div><div class="l">Baixada</div><div class="p">{pct(baixadas)}%</div></div>\n'
    h += f'<div class="sc s-in"><div class="n">{inativas:,}</div><div class="l">Inativa</div><div class="p">{pct(inativas)}%</div></div>\n'
    h += f'<div class="sc s-er"><div class="n">{erros:,}</div><div class="l">Erro</div><div class="p">{pct(erros)}%</div></div>\n'
    h += '</div>\n'

    # Tabela de distribuição com barra de progresso
    dist_rows = [
        ("Ativa",   ativas,   _CORES["ativa"],   "s-at"),
        ("Baixada", baixadas, _CORES["baixada"],  "s-bx"),
        ("Inativa", inativas, _CORES["inativa"],  "s-in"),
        ("Erro / Desconhecido", erros, _CORES["erro"], "s-er"),
    ]

    h += '<table class="dist-table">\n'
    h += '<tr><th>Status</th><th style="width:90px;text-align:right">Qtd</th><th style="width:70px;text-align:right">%</th><th>Distribuição</th></tr>\n'
    for label, qtd, cor, cls in dist_rows:
        p = float(pct(qtd))
        h += (
            f'<tr>'
            f'<td><span class="sb" style="background:{cor}">{esc(label)}</span></td>'
            f'<td style="text-align:right"><strong>{qtd:,}</strong></td>'
            f'<td style="text-align:right">{p:.1f}%</td>'
            f'<td><div class="rb"><div class="ri" style="background:{cor};width:{min(p,100):.1f}%"></div></div></td>'
            f'</tr>\n'
        )
    h += '</table>\n'

    # Parágrafo de contexto
    if ativas == total:
        h += '<div class="ib g"><p><strong>Todos os CNPJs consultados estão com situação Ativa.</strong></p></div>\n'
    else:
        h += (
            f'<div class="ib w"><p>'
            f'<strong>{n_nao:,} empresa(s) ({pct(n_nao)}%)</strong> não estão com situação Ativa '
            f'— detalhamento na seção 2.'
            f'</p></div>\n'
        )

    # ── Seção 2: Empresas Não-Ativas ─────────────────────────────────────────
    h += f'<h2 id="s2">2. Empresas Não-Ativas <span class="cnt-badge">{n_nao:,}</span></h2>\n'

    if n_nao == 0:
        h += '<div class="ib g"><p>Nenhuma empresa fora da situação Ativa.</p></div>\n'
    else:
        h += (
            '<div class="filtro-wrap">'
            '<input type="text" id="filtro" placeholder="Filtrar por CNPJ, nome ou status..." '
            'onkeyup="filtrarTabela()">'
            f'<span>Exibindo <strong id="cnt-vis">{n_nao}</strong> de {n_nao:,}</span>'
            '</div>\n'
        )

        h += '<table id="tbl-nao-ativas">\n'
        h += (
            '<thead><tr>'
            '<th style="white-space:nowrap;width:160px">CNPJ</th>'
            '<th>Nome Fantasia / Razão Social</th>'
            '<th style="white-space:nowrap">Status</th>'
            '<th style="white-space:nowrap;width:130px">Data de Alteração</th>'
            '<th>Motivo</th>'
            '</tr></thead>\n'
            '<tbody>\n'
        )

        for _, row in df_nao.iterrows():
            cnpj_fmt  = _formatar_cnpj(row["cnpj"])
            nome      = esc(_nome_display(row))
            status    = str(row["status"])
            cor       = _cor_status(status)
            data_alt  = _formatar_data(row.get("data_situacao_cadastral", ""))
            motivo    = esc(str(row.get("motivo_situacao_cadastral", "") or "—"))
            if motivo in ("", "nan", "None"):
                motivo = "—"

            h += (
                f'<tr>'
                f'<td style="white-space:nowrap;width:160px"><code>{esc(cnpj_fmt)}</code></td>'
                f'<td>{nome}</td>'
                f'<td style="white-space:nowrap"><span class="sb" style="background:{cor}">{esc(status)}</span></td>'
                f'<td style="white-space:nowrap;width:130px">{esc(data_alt)}</td>'
                f'<td>{motivo}</td>'
                f'</tr>\n'
            )

        h += '</tbody>\n</table>\n'

    h += '</div>\n'  # .cnt

    h += f'<div class="ft">CNPJ-STATUS &nbsp;|&nbsp; {now} &nbsp;|&nbsp; <strong>CONFIDENCIAL</strong></div>\n'
    h += '</div>\n'  # .ctn

    h += _JS
    h += '</body>\n</html>'

    saida = "relatorio_cnpj_status.html"
    with open(saida, "w", encoding="utf-8") as f:
        f.write(h)

    return saida
