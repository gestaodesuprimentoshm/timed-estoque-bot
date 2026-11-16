"""
Gravação no banco do SIGCS (PostgreSQL no Supabase) — Fase 4 da migração, gravação dupla.

O robô continua gravando a planilha como sempre e, além disso, grava as mesmas leituras no
banco, que é de onde o site lê. As duas gravações são independentes: se uma falhar, a outra
acontece mesmo assim, e o robô termina com erro para a falha aparecer no GitHub Actions.

Cópia idêntica em cada repositório de robô. Original: scripts/postgres/sigcs_banco.py da pasta
central do projeto (docs/13 §9.6). Mudou aqui, copie para os outros.

Variável de ambiente:
  SIGCS_DB_URL   endereço de conexão do usuário de banco do robô (Secret do GitHub).
                 Ausente: o robô avisa e grava só a planilha (ex.: execução local sem banco).

Cada execução fica registrada em sigcs.execucao (início, fim, status, linhas, mensagem),
com o próprio usuário do robô como componente. Status: "ok" (banco e planilha gravados),
"parcial" (banco gravado, planilha falhou) e "erro" (banco não gravado).
"""
import os

VARIAVEL = "SIGCS_DB_URL"


def _resumo(erro, publico):
    """Mensagem curta do erro. Em repositório público, só o tipo: o log é visível para todos."""
    if publico:
        return type(erro).__name__
    texto = " ".join(str(erro).split())
    return f"{type(erro).__name__}: {texto[:300]}"


class Banco:
    def __init__(self, publico=False):
        self.url = os.environ.get(VARIAVEL, "").strip()
        self.publico = publico
        self.execucao = None
        self.gravado = False      # o banco recebeu os dados desta execução
        self.erro = None          # resumo do erro do banco, se houve

    @property
    def ativo(self):
        return bool(self.url)

    def _conectar(self):
        import psycopg
        return psycopg.connect(self.url, connect_timeout=30, application_name="sigcs_robo")

    def iniciar(self):
        """Registra o início da execução. Falhar aqui não impede o robô de seguir."""
        if not self.ativo:
            print(f"[BANCO] {VARIAVEL} não configurada: só a planilha será gravada.")
            return
        try:
            with self._conectar() as con:
                self.execucao = con.execute("select sigcs.execucao_inicio()").fetchone()[0]
            print(f"[BANCO] Execução nº {self.execucao} registrada.")
        except Exception as e:
            print(f"[BANCO] Não foi possível registrar o início ({_resumo(e, self.publico)}).")

    def gravar(self, funcao, tentativas=3):
        """Roda funcao(con, execucao) numa transação: ou grava tudo, ou nada.

        Devolve o que a função devolver (um texto de resumo, por exemplo) ou None se falhar.
        Erros de conexão passageiros têm nova tentativa.
        """
        if not self.ativo:
            return None
        import time
        import psycopg
        for tentativa in range(1, tentativas + 1):
            try:
                with self._conectar() as con:
                    with con.transaction():
                        resultado = funcao(con, self.execucao)
                self.gravado = True
                self.erro = None
                print(f"[BANCO] Gravado: {resultado}")
                return resultado
            except psycopg.OperationalError as e:
                self.erro = _resumo(e, self.publico)
                if tentativa < tentativas:
                    print(f"[BANCO] Falha de conexão ({self.erro}). Nova tentativa em {15 * tentativa} s...")
                    time.sleep(15 * tentativa)
            except Exception as e:
                self.erro = _resumo(e, self.publico)
                break
        print(f"[BANCO] ERRO: os dados não foram gravados no banco ({self.erro}).")
        return None

    def finalizar(self, planilha_ok, linhas=None, mensagem=None):
        """Fecha o registro da execução. Devolve True se banco e planilha deram certo."""
        if not self.ativo:
            return planilha_ok
        if self.gravado:
            status = "ok" if planilha_ok else "parcial"
        else:
            status = "erro"
        partes = [mensagem] if mensagem else []
        if not planilha_ok:
            partes.append("planilha não gravada")
        if self.erro:
            partes.append(f"banco: {self.erro}")
        if self.execucao is not None:
            try:
                with self._conectar() as con:
                    con.execute("select sigcs.execucao_fim(%s, %s, %s, %s)",
                                (self.execucao, status, linhas, "; ".join(partes) or None))
            except Exception as e:
                print(f"[BANCO] Não foi possível registrar o fim ({_resumo(e, self.publico)}).")
        return planilha_ok and self.gravado
