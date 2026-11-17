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
    SIGCS_DB_URL        banco do SIGCS, usuário próprio deste robô    (secret, opcional)
    SIGCS_UNIDADE       id da unidade no banco do SIGCS                (opcional)

BANCO (Fase 4 da migração das planilhas para PostgreSQL)
    Além da planilha, o robô grava as mesmas leituras no banco do SIGCS, de onde o site
    lê. As duas gravações são independentes: se uma falhar, a outra acontece mesmo assim,
    e o job termina em vermelho. Sem SIGCS_DB_URL, grava só a planilha.

LISTA DE ITENS (Fase 5)
    A lista de itens (grupo e código TIMED) vem do banco do SIGCS, onde é editada pela tela de
    cadastro do site. Cada código é consultado uma vez; na planilha, recebem I–N as linhas cuja
    coluna A tem o mesmo código no mesmo grupo. Sem banco (SIGCS_DB_URL ausente) ou com o banco
    fora do ar, o robô usa a coluna A da planilha, como antes.

REGRAS
    - a busca usa o código exatamente como está na lista (sem espaços);
    - item que o TIMED não encontrar fica com estoque, consumo e valor ZERO,
      e não com o valor antigo congelado;
    - o estoque soma apenas as seções configuradas para o grupo do item, e
      apenas lotes dentro da validade.
"""

import json
import os
import time
from datetime import datetime, timedelta, timezone

import gspread
import requests
from google.oauth2.service_account import Credentials

FUSO_BR = timezone(timedelta(hours=-3))
TEMPO_LIMITE = 30          # segundos por requisição
LOTE_GRAVACAO = 150        # células por escrita no Sheets
# Com a lista do banco, a planilha é toda gravada no fim, de uma vez: lotes maiores e um intervalo
# mínimo entre escritas, para não estourar a cota do Sheets (60 escritas por minuto). Em 17/11 os
# ~65 lotes de 150 células saíram em segundos e o Google recusou (APIError).
LOTE_GRAVACAO_BANCO = 1200
INTERVALO_ESCRITAS = 1.2   # segundos

# Grupo da coluna de grupo (sem acento, maiúsculas) -> grupo no banco do SIGCS
GRUPOS_BANCO = {
    "MEDICAMENTOS": "medicamentos",
    "MATERIAL MEDICO HOSPITALAR": "material",
    "ALMOXARIFADO CENTRAL": "almoxarifado",
}


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


def lista_do_banco(banco, unidade):
    """Lista de itens do hospital no banco do SIGCS: [(grupo, código), ...] (Fase 5).

    Desde a Fase 5 a lista é editada na tela de cadastro do SIGCS. Devolve None (e o robô usa
    a coluna A da planilha, como antes) sem banco configurado ou se a leitura falhar.
    """
    if not banco.ativo or not unidade:
        return None
    try:
        with banco._conectar() as con:
            linhas = con.execute("select grupo, codigo from sigcs.itens_para_leitura(%s)", (unidade,)).fetchall()
        return [(g, c or "") for g, c in linhas]
    except Exception as e:
        print(f"[BANCO] Lista de itens indisponível ({type(e).__name__}): usando a coluna A da planilha.")
        return None


def sem_acento(texto):
    import unicodedata
    return "".join(c for c in unicodedata.normalize("NFD", texto) if unicodedata.category(c) != "Mn")


# ─────────────────────────────────────────────────────────────────────────────
def main():
    from sigcs_banco import Banco
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

    unidade = os.environ.get("SIGCS_UNIDADE", "").strip()
    banco = Banco(publico=True)
    if banco.ativo and not unidade:
        print("[BANCO] SIGCS_UNIDADE não configurada: só a planilha será gravada.")
        banco.url = ""
    banco.iniciar()
    leituras = []              # o que vai para o banco: grupo, código e valores de cada linha gravada
    planilha = {"ok": True, "ultima": 0.0}

    def enviar(celulas_lote):
        """Grava um lote na planilha. Cota estourada (429) ou falha do Google (5xx): espera e tenta
        de novo, até 3 vezes. Outra falha, ou as 3 tentativas esgotadas: para de tentar e segue
        (o banco ainda recebe)."""
        if not planilha["ok"]:
            return
        for tentativa in range(1, 4):
            espera = INTERVALO_ESCRITAS - (time.monotonic() - planilha["ultima"])
            if espera > 0:
                time.sleep(espera)
            try:
                aba.update_cells(celulas_lote)
                planilha["ultima"] = time.monotonic()
                return
            except Exception as e:
                planilha["ultima"] = time.monotonic()
                status = getattr(getattr(e, "response", None), "status_code", None)
                if (status == 429 or (status or 0) >= 500) and tentativa < 3:
                    print(f"Planilha recusou a escrita (HTTP {status}); nova tentativa em 65 s.")
                    time.sleep(65)
                    continue
                planilha["ok"] = False
                print(f"Falha ao gravar a planilha: {type(e).__name__} (HTTP {status}). O banco ainda será gravado.")
                return

    codigos = aba.col_values(1)          # A — Cód. TIMED
    grupos  = aba.col_values(col_grupo)  # coluna de grupo
    total_linhas = max(len(codigos), len(grupos))
    print(f"Aba '{nome_aba}': {total_linhas - 1} linhas de dados.")

    # Lista de itens: do banco (Fase 5: a lista é editada na tela de cadastro do SIGCS).
    # Sem banco ou com o banco fora do ar, a coluna A da planilha, como antes.
    lista = lista_do_banco(banco, unidade)
    print(f"Lista de itens: {'banco do SIGCS (' + str(len(lista)) + ' códigos)' if lista is not None else 'coluna A da planilha'}.")

    timed = Timed(base_url)
    timed.entrar(usuario, senha)
    print("Login no TIMED: ok.")

    agora_br = datetime.now(FUSO_BR)
    hoje = agora_br.replace(hour=0, minute=0, second=0, microsecond=0, tzinfo=None)
    agora = agora_br.strftime("%d/%m/%Y %H:%M:%S")

    celulas = []
    n = {"encontrados": 0, "zerados": 0, "sem_codigo": 0, "grupo_desconhecido": 0, "erros": 0, "fora_da_lista": 0}

    def consultar(codigo, secoes, rotulo):
        """Valores de um código: (estoque, consumo ideal, consumo do mês, consumo 10 dias, valor), ou None se falhou."""
        if not codigo.isdigit():
            # Código não numérico ("-", "s/ cadastro"): não adianta consultar
            n["sem_codigo"] += 1
            return (0, 0, 0, 0, 0.0)
        try:
            produto_id = timed.buscar(codigo)
            if produto_id is None:
                # Decisão do usuário: não encontrou, é zero — e não o valor antigo
                n["zerados"] += 1
                return (0, 0, 0, 0, 0.0)
            dados = timed.consumo(produto_id)
            valores = (somar_estoque(timed.estoque(produto_id), secoes, hoje),
                       consumo_mensal(dados),
                       int(round(float(dados.get("consumoMesCorrente") or 0), 0)),
                       int(round(float(dados.get("consumoMedioDiario10Dias") or 0), 0)),
                       float(dados.get("valorUnitarioMedio") or 0))
            n["encontrados"] += 1
            return valores
        except Exception as e:
            # Sem dado item a item no log: o repositório é público
            n["erros"] += 1
            print(f"[{rotulo}] falha na consulta: {type(e).__name__}")
            return None

    def gravar_linha(linha, valores, lote=LOTE_GRAVACAO):
        nonlocal celulas
        for coluna, valor_celula in zip((9, 10, 11, 12, 13), valores):
            celulas.append(gspread.Cell(row=linha, col=coluna, value=valor_celula))
        celulas.append(gspread.Cell(row=linha, col=14, value=agora))
        if len(celulas) >= lote:
            enviar(celulas)
            celulas = []

    def leitura(grupo_banco, codigo, valores):
        estoque, consumo_ideal, consumo_mes, consumo_10d, valor = valores
        leituras.append({
            "grupo": grupo_banco, "codigo": codigo, "estoque": estoque,
            "consumo_mensal": consumo_ideal, "consumo_mes_corrente": consumo_mes,
            "consumo_10_dias": consumo_10d, "valor_unitario": valor, "lido_em": agora_br.isoformat(),
        })

    if lista is not None:
        # Cada código do banco é consultado uma vez; a planilha recebe I–N nas linhas cuja coluna A
        # tem o mesmo código no mesmo grupo (até a Fase 6 ela continua sendo gravada).
        secoes_banco = {GRUPOS_BANCO.get(sem_acento(nome)): secoes for nome, secoes in por_grupo.items()}
        resultados = {}
        for k, (grupo_banco, codigo) in enumerate(lista, start=1):
            codigo = " ".join((codigo or "").split())
            secoes = secoes_banco.get(grupo_banco)
            if secoes is None:
                n["grupo_desconhecido"] += 1
                continue
            valores = consultar(codigo, secoes, f"item {k}")
            if valores is None:
                continue
            resultados[(grupo_banco, codigo)] = valores
            leitura(grupo_banco, codigo, valores)
        for i in range(1, total_linhas):           # pula o cabeçalho
            codigo = " ".join((codigos[i] if i < len(codigos) else "").split())
            grupo  = (grupos[i] if i < len(grupos) else "").strip().upper()
            if not codigo and not grupo:
                continue
            valores = resultados.get((GRUPOS_BANCO.get(sem_acento(grupo)), codigo))
            if valores is None:
                n["fora_da_lista"] += 1           # linha da planilha que não está na lista do banco: não mexe
                continue
            gravar_linha(i + 1, valores, LOTE_GRAVACAO_BANCO)
    else:
        for i in range(1, total_linhas):           # pula o cabeçalho
            linha = i + 1
            codigo = (codigos[i] if i < len(codigos) else "").strip()
            grupo  = (grupos[i] if i < len(grupos) else "").strip().upper()
            if not codigo and not grupo:
                continue
            secoes = por_grupo.get(grupo)
            if codigo.isdigit() and secoes is None:
                n["grupo_desconhecido"] += 1
                continue                        # grupo fora da configuração: não mexe na linha
            valores = consultar(codigo, secoes, f"linha {linha}")
            if valores is None:
                continue
            gravar_linha(linha, valores)
            grupo_banco = GRUPOS_BANCO.get(sem_acento(grupo))
            if grupo_banco:
                leitura(grupo_banco, " ".join(codigo.split()), valores)

    if celulas:
        enviar(celulas)

    # Banco: com nenhum item encontrado (credencial recusada, TIMED fora do ar), os zeros não
    # vão para o banco — lá fica a leitura anterior, e o job termina em vermelho logo abaixo.
    resultado = None
    if n["encontrados"] > 0:
        def gravar(con, execucao):
            from psycopg.types.json import Jsonb
            return con.execute("select sigcs.gravar_leituras(%s, %s, %s)",
                               (unidade, Jsonb(leituras), execucao)).fetchone()[0]
        resultado = banco.gravar(gravar)
    mensagem = (f"lista: {'banco' if lista is not None else 'planilha'}; {n['encontrados']} encontrados, {n['zerados']} zerados, "
                f"{n['sem_codigo']} sem código, {n['erros']} erros"
                + (f", {n['fora_da_lista']} linhas da planilha fora da lista" if lista is not None else "")
                + (f"; banco: {resultado['leituras']} leituras, {resultado['sem_par']} códigos sem par" if resultado else ""))
    banco_ok = banco.finalizar(planilha["ok"], linhas=resultado["leituras"] if resultado else None, mensagem=mensagem)

    print("── Resumo ──")
    print(f"  lista de itens       : {'banco do SIGCS' if lista is not None else 'coluna A da planilha'}")
    print(f"  encontrados no TIMED : {n['encontrados']}")
    print(f"  zerados (sem retorno): {n['zerados']}")
    print(f"  sem código válido    : {n['sem_codigo']}")
    print(f"  grupo não configurado: {n['grupo_desconhecido']}")
    print(f"  erros de consulta    : {n['erros']}")
    if lista is not None:
        print(f"  linhas da planilha fora da lista do banco (não gravadas): {n['fora_da_lista']}")
    if resultado:
        print(f"  banco: {resultado['leituras']} leituras gravadas, {resultado['sem_par']} códigos sem par no banco")

    if n["encontrados"] == 0:
        # Falha silenciosa é o pior defeito possível num robô agendado:
        # aqui o job termina em vermelho, para o problema aparecer.
        raise SystemExit("Nenhum item encontrado no TIMED. Verifique credenciais e códigos.")
    if not banco_ok:
        raise SystemExit("A planilha ou o banco não foram gravados (ver as mensagens acima).")


if __name__ == "__main__":
    main()
