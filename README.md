# CNPJ-STATUS

Consulta em lote a situação cadastral de CNPJs diretamente pela API pública [OpenCNPJ](https://opencnpj.org), sem necessidade de chave ou cadastro. Ideal para higienização e enriquecimento de bases de dados.

---

## Sumário

- [Funcionalidades](#funcionalidades)
- [Pré-requisitos](#pré-requisitos)
- [Instalação](#instalação)
- [Uso](#uso)
- [Formato da planilha de entrada](#formato-da-planilha-de-entrada)
- [Saída gerada](#saída-gerada)
- [Mapeamento de status](#mapeamento-de-status)
- [Política de limites e cadência](#política-de-limites-e-cadência)
- [Tratamento de erros](#tratamento-de-erros)
- [Estrutura do projeto](#estrutura-do-projeto)

---

## Funcionalidades

- Lê planilhas `.csv` ou `.xlsx` com uma coluna `cnpj`
- Normaliza automaticamente o CNPJ (remove pontuação, preenche com zeros à esquerda)
- Consulta o endpoint `GET /cnpj/{cnpj}` da API OpenCNPJ
- Mapeia a situação cadastral para um status padronizado em português
- Processa em **lotes com pausas**, respeitando a política anti-bloqueio da API
- Retenativas automáticas com backoff progressivo em caso de `429`
- Salva o resultado em `resultado_cnpjs.xlsx` com a coluna `status` adicionada
- Exibe progresso em tempo real e estimativa de duração total

---

## Pré-requisitos

- Python **3.8+**
- pip

---

## Instalação

```bash
# 1. Clone o repositório
git clone https://github.com/trickMeister1337/CNPJ-STATUS.git
cd CNPJ-STATUS

# 2. (Opcional, recomendado) Crie um ambiente virtual
python -m venv .venv
source .venv/bin/activate       # Linux / macOS
.venv\Scripts\activate          # Windows

# 3. Instale as dependências
pip install requests pandas openpyxl
```

> **Sem `requirements.txt`?** O projeto usa apenas bibliotecas amplamente disponíveis. Se preferir fixar versões, rode:
> ```bash
> pip freeze > requirements.txt
> ```

---

## Uso

```bash
python cnpj_status.py <arquivo>
```

### Exemplos

```bash
# Arquivo CSV
python cnpj_status.py clientes.csv

# Arquivo Excel
python cnpj_status.py fornecedores.xlsx
```

### Saída no terminal

```
[*] Carregando planilha: clientes.xlsx
[*] 10000 CNPJs encontrados.
[*] Cadência: 0.3s entre requests, pausa de 45s a cada 50 consultas.
[*] Tempo estimado: ~2h 40min

  [    1/10000] ✓ 00000000000191 → Ativa
  [    2/10000] ✓ 33000167000101 → Ativa
  [    3/10000] ✗ 00000000000000 → Erro na consulta: CNPJ não encontrado (404)
  ...
  [   50/10000] ✓ 60701190000104 → Baixada

  [lote 1/200 concluído] aguardando 45s...

  [   51/10000] ✓ ...
  ...

[+] Concluído. Resultado salvo em: resultado_cnpjs.xlsx

--- Resumo ---
  Ativa: 7832
  Baixada: 1541
  Inativa: 498
  Erro na consulta: CNPJ não encontrado (404): 129
```

---

## Formato da planilha de entrada

A planilha deve conter **obrigatoriamente** uma coluna chamada `cnpj`. As demais colunas são preservadas intactas na saída.

| cnpj | razao_social | ... |
|---|---|---|
| 00.000.000/0001-91 | Empresa A | ... |
| 33000167000101 | Empresa B | ... |
| 60701190000104 | Empresa C | ... |

O script aceita CNPJs com ou sem pontuação (`/`, `.`, `-`).

---

## Saída gerada

O arquivo `resultado_cnpjs.xlsx` contém todas as colunas originais acrescidas da coluna `status`:

| cnpj | razao_social | ... | status |
|---|---|---|---|
| 00.000.000/0001-91 | Empresa A | ... | Ativa |
| 33000167000101 | Empresa B | ... | Baixada |
| 60701190000104 | Empresa C | ... | Inativa |

> O arquivo de saída é sempre salvo no **diretório de trabalho atual** onde o script é executado.

---

## Mapeamento de status

| Situação Cadastral (API) | Status no resultado |
|---|---|
| `Ativa` | `Ativa` |
| `Baixada` | `Baixada` |
| `Suspensa` | `Inativa` |
| `Inapta` | `Inativa` |
| `Nula` | `Inativa` |
| Qualquer outro valor | `Status Desconhecido: [valor]` |
| Falha na requisição | `Erro na consulta: [mensagem]` |

---

## Política de limites e cadência

A API OpenCNPJ **não exige chave de acesso** e aceita picos pontuais acima de 100 req/s. O bloqueio (`429`) é acionado quando uma mesma origem mantém **volume contínuo por período prolongado** — não pela velocidade individual das chamadas.

Para processar grandes volumes (ex.: 10 mil CNPJs) sem acionar o bloqueio, o script usa uma estratégia de **lotes com pausas**:

| Parâmetro | Valor padrão | Descrição |
|---|---|---|
| `DELAY_ENTRE_REQUESTS` | `0.3s` | Intervalo entre requisições dentro do lote |
| `TAMANHO_LOTE` | `50` | Quantidade de CNPJs por lote |
| `PAUSA_ENTRE_LOTES` | `45s` | Pausa ao final de cada lote |
| `RETRY_DELAYS` | `60s / 120s / 180s` | Backoff progressivo para `429` |

### Por que o backoff é longo?

Quando a API retorna `429`, o bloqueio só é liberado após o tráfego da origem normalizar. Esperas curtas (ex.: 5s) simplesmente re-acionam o bloqueio na próxima chamada. Os valores de 60/120/180s garantem janela suficiente para o bloqueio ser removido antes da próxima tentativa.

### Estimativa de tempo para 10 mil CNPJs

```
200 lotes × (50 req × 0,3s + 45s de pausa) ≈ 2h 40min
```

Para ajustar a cadência, edite as constantes no topo do arquivo `cnpj_status.py`.

---

## Tratamento de erros

| Situação | Mensagem gravada na coluna `status` |
|---|---|
| HTTP 404 | `Erro na consulta: CNPJ não encontrado (404)` |
| HTTP 429 após 3 tentativas | `Erro na consulta: rate limit (429) após 3 tentativas` |
| Timeout | `Erro na consulta: timeout` |
| Falha de conexão | `Erro na consulta: falha de conexão` |
| JSON inválido / campo ausente | `Erro na consulta: resposta inesperada (...)` |
| Outro código HTTP | `Erro na consulta: HTTP [código]` |

Erros individuais **não interrompem** o processamento — o script continua para o próximo CNPJ e registra a falha na planilha.

---

## Estrutura do projeto

```
CNPJ-STATUS/
├── cnpj_status.py   # Script principal
├── .gitignore
└── README.md
```

---

## Aviso legal

Este projeto consulta uma API pública para fins de higienização de dados cadastrais. Use de forma responsável e em conformidade com a política de uso do [OpenCNPJ](https://opencnpj.org). Não utilize para varredura de segmentos, montagem de listas ou geração de estatísticas — para esses casos, utilize o [dataset público no BigQuery](https://opencnpj.org).
