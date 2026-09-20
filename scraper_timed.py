"""
Robô de estoque e consumo de uma unidade hospitalar no TIMED (plataforma Vitai).

Lê os códigos de uma aba de uma planilha do Google e grava, para cada item,
as colunas I–N: estoque na unidade, consumo mensal, consumo do mês corrente,
consumo diário dos últimos 10 dias, valor unitário e data da atualização.

REPOSITÓRIO PÚBLICO
    Este arquivo é feito para viver num repositório público, onde o GitHub
    Actions é ilimitado. Por isso:
      - nenhum endereço, id de planilha ou nome de seção está escrito aqui:
        tudo vem de variável de ambiente;
      - os logs não trazem dado item a item, só contagens — em repositório
        público os registros de execução são visíveis para qualquer pessoa.

VARIÁVEIS DE AMBIENTE
    GOOGLE_CREDENTIALS  JSON da conta de serviço                     (secret)
    TIMED_BASE_URL      raiz do sistema, sem barra no fim            (secret)
    TIMED_USER          usuário do TIMED                             (secret)
    TIMED_PASS          senha do TIMED                               (secret)
    SHEET_ID            id da planilha de destino                    (secret)
    ABA_DESTINO         nome da aba de destino
    COLUNA_GRUPO        letra da coluna que diz o grupo do item (ex.: AH)
    SECOES_POR_GRUPO    "Grupo=SEÇÃO1|SEÇÃO2; Outro Grupo=SEÇÃO"

REGRAS
    - a busca usa o código da coluna A exatamente como está na planilha;
    - item que o TIMED não encontrar fica com estoque, consumo e valor ZERO,
      e não com o valor antigo congelado;
    - o estoque soma apenas as seções configuradas para o grupo do item, e
      apenas lotes dentro da validade.
"""

import json
import os
from datetime import datetime, timedelta, timezone

import gspread
import requests
from google.oauth2.service_account import Credentials

FUSO_BR = timezone(timedelta(hours=-3))
TEMPO_LIMITE = 30          # segundos por requisição
LOTE_GRAVACAO = 150        # células por escrita no Sheets


# ─────────────────────────────────────────────────────────────────────────────
# Configuração
# ─────────────────────────────────────────────────────────────────────────────
def obrigatoria(nome):
    valor = os.environ.get(nome)
    if not valor:
        raise ValueError(f"Variável de ambiente ausente: {nome}")
    return valor


def ler_secoes_por_grupo(texto):
    """Texto no formato 'Grupo=SEÇÃO; Outro Grupo=SEÇÃO A|SEÇÃO B'."""
    mapa = {}
    for parte in texto.split(";"):
        parte = parte.strip()
        if not parte:
            continue
        if "=" not in parte:
            raise ValueError(f"SECOES_POR_GRUPO mal formado em: {parte!r}")
        grupo, secoes = parte.split("=", 1)
        mapa[grupo.strip().upper()] = [s.strip().upper() for s in secoes.split("|") if s.strip()]
    if not mapa:
        raise ValueError("SECOES_POR_GRUPO está vazio.")
    return mapa


def letra_para_indice(letra):
    """'A' -> 1, 'AH' -> 34 (base 1, como o gspread espera)."""
    n = 0
    for c in letra.strip().upper():
        if not ("A" <= c <= "Z"):
            raise ValueError(f"Coluna inválida: {letra!r}")
        n = n * 26 + (ord(c) - 64)
    return n


# ─────────────────────────────────────────────────────────────────────────────
# Regras de negócio
# ─────────────────────────────────────────────────────────────────────────────
def consumo_mensal(dados):
    """12 meses, descarta zerados; com 4 ou mais, remove o maior e o menor."""
    consumos = []
    for i in range(1, 13):
        v = dados.get(f"consumoMesPassado{i}")
        if v is not None:
            consumos.append(float(v))
    validos = [c for c in consumos if c > 0]
    if not validos:
        return 0
    if len(validos) >= 4:
        validos.remove(max(validos))
        validos.remove(min(validos))
    return int(round(sum(validos) / len(validos), 0))


def somar_estoque(linhas, secoes, hoje):
    """Soma a quantidade disponível nas seções pedidas, só de lotes no prazo."""
    total = 0.0
    if not isinstance(linhas, list):
        return 0
    for local in linhas:
        if str(local.get("secaoNome", "")).strip().upper() not in secoes:
            continue
        venc = local.get("dataVencimentoAtual")
        if venc:
            try:
                if datetime.strptime(venc, "%d/%m/%Y") < hoje:
                    continue          # lote vencido não conta
            except ValueError:
                pass                  # data em formato inesperado: mantém o lote
        try:
            total += float(local.get("quantidadeDisponivel", 0))
        except (TypeError, ValueError):
            pass
    return int(round(total, 0))


# ─────────────────────────────────────────────────────────────────────────────
# TIMED
# ─────────────────────────────────────────────────────────────────────────────
class Timed:
    """As respostas vêm em ISO-8859-1, não UTF-8."""

    def __init__(self, base_url):
        self.base = base_url.rstrip("/")
        self.s = requests.Session()
        self.s.headers.update({
            "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
            "accept": "application/json, text/javascript, */*; q=0.01",
            "x-requested-with": "XMLHttpRequest",
        })

    def _json(self, resp):
        resp.encoding = "ISO-8859-1"
        try:
            return resp.json()
        except ValueError:
            return None

    def entrar(self, usuario, senha):
        url = f"{self.base}/vitai/pages/login.do"
        self.s.get(url, timeout=TEMPO_LIMITE)
        r = self.s.post(url, timeout=TEMPO_LIMITE, data={
            "perform": "login", "controle": "1", "login": usuario, "senha": senha,
        })
        if 'id="senha"' in r.text or 'name="senha"' in r.text:
            raise ValueError("O TIMED devolveu a tela de login: credenciais recusadas.")

    def buscar(self, codigo):
        """Devolve o id interno do produto, ou None.

        Confere o codigoAuxiliar antes de aceitar, para não casar um item
        parecido quando a busca devolve mais de um resultado.
        """
        r = self.s.get(f"{self.base}/vitai/produto/produto.do", timeout=TEMPO_LIMITE,
                       params={"perform": "procurarPorNome", "pesquisa": codigo, "controle": "18"})
        achados = self._json(r)
        if not isinstance(achados, list) or not achados:
            return None
        for p in achados:
            if str(p.get("codigoAuxiliar", "")).strip() == codigo:
                return p.get("id") or p.get("value")
        return None

    def consumo(self, produto_id):
        r = self.s.post(f"{self.base}/vitai/produto/avaliacaoConsumoCompra.do", timeout=TEMPO_LIMITE,
                        data={"perform": "pesquisar", "controle": "4",
                              "produtoId": str(produto_id), "incluirProdutosSemConsumo": "S"})
        j = self._json(r)
        if isinstance(j, list) and j:
            return j[0]
        return j if isinstance(j, dict) else {}

    def estoque(self, produto_id):
        r = self.s.get(f"{self.base}/vitai/produto/consultaEstoqueProduto.do", timeout=TEMPO_LIMITE,
                       params={"perform": "pesquisar", "produtoId": str(produto_id),
                               "posicaoId": "", "controle": "4"})
        return self._json(r)


# ─────────────────────────────────────────────────────────────────────────────
def main():
    base_url   = obrigatoria("TIMED_BASE_URL")
    usuario    = obrigatoria("TIMED_USER")
    senha      = obrigatoria("TIMED_PASS")
    sheet_id   = obrigatoria("SHEET_ID")
    nome_aba   = obrigatoria("ABA_DESTINO")
    col_grupo  = letra_para_indice(obrigatoria("COLUNA_GRUPO"))
    por_grupo  = ler_secoes_por_grupo(obrigatoria("SECOES_POR_GRUPO"))

    credenciais = Credentials.from_service_account_info(
        json.loads(obrigatoria("GOOGLE_CREDENTIALS")),
        scopes=["https://www.googleapis.com/auth/spreadsheets"],
    )
    aba = gspread.authorize(credenciais).open_by_key(sheet_id).worksheet(nome_aba)

    codigos = aba.col_values(1)          # A — Cód. TIMED
    grupos  = aba.col_values(col_grupo)  # coluna de grupo
    total_linhas = max(len(codigos), len(grupos))
    print(f"Aba '{nome_aba}': {total_linhas - 1} linhas de dados.")

    timed = Timed(base_url)
    timed.entrar(usuario, senha)
    print("Login no TIMED: ok.")

    agora_br = datetime.now(FUSO_BR)
    hoje = agora_br.replace(hour=0, minute=0, second=0, microsecond=0, tzinfo=None)
    agora = agora_br.strftime("%d/%m/%Y %H:%M:%S")

    celulas = []
    n_encontrados = n_zerados = n_sem_codigo = n_grupo_desconhecido = n_erros = 0

    for i in range(1, total_linhas):           # pula o cabeçalho
        linha = i + 1
        codigo = (codigos[i] if i < len(codigos) else "").strip()
        grupo  = (grupos[i] if i < len(grupos) else "").strip().upper()

        if not codigo and not grupo:
            continue

        # Código não numérico ("-", "s/ cadastro"): não adianta consultar
        if not codigo.isdigit():
            n_sem_codigo += 1
            estoque = consumo_ideal = consumo_mes = consumo_10d = 0
            valor = 0.0
        else:
            secoes = por_grupo.get(grupo)
            if secoes is None:
                n_grupo_desconhecido += 1
                continue                        # grupo fora da configuração: não mexe na linha
            try:
                produto_id = timed.buscar(codigo)
                if produto_id is None:
                    # Decisão do usuário: não encontrou, é zero — e não o valor antigo
                    n_zerados += 1
                    estoque = consumo_ideal = consumo_mes = consumo_10d = 0
                    valor = 0.0
                else:
                    dados = timed.consumo(produto_id)
                    estoque       = somar_estoque(timed.estoque(produto_id), secoes, hoje)
                    consumo_ideal = consumo_mensal(dados)
                    consumo_mes   = int(round(float(dados.get("consumoMesCorrente") or 0), 0))
                    consumo_10d   = int(round(float(dados.get("consumoMedioDiario10Dias") or 0), 0))
                    valor         = float(dados.get("valorUnitarioMedio") or 0)
                    n_encontrados += 1
            except Exception as e:
                # Sem dado item a item no log: o repositório é público
                n_erros += 1
                print(f"[linha {linha}] falha na consulta: {type(e).__name__}")
                continue

        for coluna, valor_celula in (
            (9,  estoque), (10, consumo_ideal), (11, consumo_mes),
            (12, consumo_10d), (13, valor), (14, agora),
        ):
            celulas.append(gspread.Cell(row=linha, col=coluna, value=valor_celula))

        if len(celulas) >= LOTE_GRAVACAO:
            aba.update_cells(celulas)
            celulas.clear()

    if celulas:
        aba.update_cells(celulas)

    print("── Resumo ──")
    print(f"  encontrados no TIMED : {n_encontrados}")
    print(f"  zerados (sem retorno): {n_zerados}")
    print(f"  sem código válido    : {n_sem_codigo}")
    print(f"  grupo não configurado: {n_grupo_desconhecido}")
    print(f"  erros de consulta    : {n_erros}")

    if n_encontrados == 0:
        # Falha silenciosa é o pior defeito possível num robô agendado:
        # aqui o job termina em vermelho, para o problema aparecer.
        raise SystemExit("Nenhum item encontrado no TIMED. Verifique credenciais e códigos.")


if __name__ == "__main__":
    main()
