#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
fca_consolidar.py — Consolida as 10 planilhas de FCA (Forms) e gera o dashboard.

Uso:
    python fca_consolidar.py                      # lê ./planilhas e gera ./index.html
    python fca_consolidar.py "C:\\caminho\\FCAS"    # pasta com as planilhas
    python fca_consolidar.py --publicar           # além de gerar, faz git add/commit/push

Regras (vindas do histórico do projeto — NÃO alteram a base original):
  * Cada linha da planilha é 1 FCA aberto. Nada é descartado, nem duplicatas.
  * Turno em branco é inferido pela "Hora de conclusão":
        A 07:01–15:00 | B 15:01–22:00 | C 22:01–07:00
  * Operador só vale se for nome de pessoa. Senão:
        QRZ -> "Operador COI – Turno X"   |   CLE -> "Não identificado"
  * Última ação = árvore de decisão: a coluna que recebeu "Finalizar FCA";
    se não houver, a última coluna preenchida da direita. Se o operador
    descreveu a ação em texto livre ("Outra"), o texto livre prevalece.
  * Sub-processo vazio recebe o nome da área (nome da pasta/planilha).
  * Fato vazio recebe o texto de "Descreva a anomalia" (quando existir).

Requer: Python 3.8+ e openpyxl  (pip install openpyxl)
"""
import csv
import json
import os
import re
import subprocess
import sys
import unicodedata
from datetime import datetime
from pathlib import Path

try:
    import openpyxl
except ImportError:
    sys.exit("Falta instalar o openpyxl:  pip install openpyxl")

AQUI = Path(__file__).resolve().parent
LINKS_URL = os.environ.get("FCA_LINKS_URL")  # modo automático (GitHub Actions)
PASTA_PLANILHAS = Path(os.environ.get("FCA_PASTA", AQUI / "planilhas"))
TEMPLATE = AQUI / "template_fca.html"
SAIDA_HTML = AQUI / "index.html"
SAIDA_CSV = AQUI / "fca_base_consolidada.csv"

AREAS = {
    "extracao": "Extração de Caldo",
    "tratamento": "Tratamento de Caldo",
    "acucar": "Fábrica de Açúcar",
    "etanol": "Fábrica de Etanol",
    "vapor": "Geração de Vapor",
}
UNIDADES = {"clementina": ("CLE", "Clementina"), "queiroz": ("QRZ", "Queiroz")}
PARTICULAS = {"de", "da", "do", "das", "dos", "e"}


# ----------------------------------------------------------------- utilidades
def limpar(v):
    """Texto sem NBSP e espaços duplicados; None se vazio."""
    if v is None:
        return None
    s = re.sub(r"\s+", " ", str(v).replace("\xa0", " ")).strip()
    return s or None


def norm(s):
    """minúsculo, sem acento, sem NBSP — só para comparar."""
    s = unicodedata.normalize("NFKD", str(s or "").replace("\xa0", " "))
    s = "".join(c for c in s if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", s).strip().lower()


def identificar(nome_arquivo):
    """Descobre (área, unidade) pelo nome do arquivo; tolera acentos quebrados."""
    n = norm(nome_arquivo).replace("_", "")
    uni = next((v for k, v in UNIDADES.items() if k in n), None)
    if "extra" in n and "caldo" in n:
        area = "extracao"
    elif "tratamento" in n:
        area = "tratamento"
    elif "etanol" in n:
        area = "etanol"
    elif "vapor" in n:
        area = "vapor"
    elif "acucar" in n or "acar" in n:
        area = "acucar"
    else:
        area = None
    return (area, uni) if area and uni else None


def para_datetime(v):
    if isinstance(v, datetime):
        return v
    if isinstance(v, (int, float)):  # serial do Excel
        return openpyxl.utils.datetime.from_excel(v)
    s = limpar(v)
    if not s:
        return None
    for fmt in ("%Y-%m-%d %H:%M:%S", "%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M", "%Y-%m-%d %H:%M"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            pass
    return None


def turno_por_hora(dt):
    """A 07:01–15:00 | B 15:01–22:00 | C 22:01–07:00 (minuto do dia)."""
    if dt is None:
        return None
    m = dt.hour * 60 + dt.minute
    if 7 * 60 + 1 <= m <= 15 * 60:
        return "Turno A"
    if 15 * 60 + 1 <= m <= 22 * 60:
        return "Turno B"
    return "Turno C"


def nome_proprio(txt):
    s = limpar(txt)
    if not s:
        return None
    partes = []
    for i, p in enumerate(s.split(" ")):
        p = p.lower()
        partes.append(p if (i > 0 and p in PARTICULAS) else p.capitalize())
    return " ".join(partes)


def eh_nome_de_pessoa(txt):
    s = limpar(txt)
    if not s or re.match(r"(?i)^operador\b", s):
        return False
    return len(s.split(" ")) >= 2 and not re.search(r"\d", s)


def operador_final(bruto, unidade_sigla, turno):
    s = limpar(bruto)
    if eh_nome_de_pessoa(s):
        return nome_proprio(s)
    letra = (turno or "Turno ?")[-1]
    if unidade_sigla == "QRZ":
        return f"Operador COI – Turno {letra}"
    return "Não identificado"


SUB_CORRECOES = {"decantancao": "Decantação", "centrifugacao": "Centrifugação"}


def sub_limpo(v):
    s = limpar(v)
    if not s:
        return None
    return SUB_CORRECOES.get(norm(s), s[0].upper() + s[1:])


def limpar_titulo_acao(h):
    """Cabeçalho da árvore -> nome da ação (tira sufixo '1' duplicado do Forms)."""
    s = limpar(h) or ""
    s = re.sub(r"(?<=[a-zà-ú\)\s])1$", "", s).strip().rstrip(":.").strip()
    s = s[:1].upper() + s[1:]
    if re.match(r"(?i)^comunicar (o |ao )?l[ií]der e aguardar orienta", s):
        s = "Comunicar o líder e aguardar orientação"
    return s


# ----------------------------------------------------- classificação de colunas
def papel(idx, h):
    n = norm(h)
    if idx <= 4:
        return "meta"
    if n in ("turno", "selecione o turno"):
        return "turno"
    if "responsavel pelo fca" in n:
        return "operador"
    if n in ("escolha o subprocesso", "escolha a area", "sub-processo", "subprocessos"):
        return "sub"
    if n.startswith("qual produto da geracao de vapor"):
        return "sub_produto"
    if n.startswith("fato") or n.endswith("(fato)"):
        return "fato"
    if n.startswith("qual caracteristica") or n.startswith("como a pressao do vapor de baixa foi afetada"):
        return "fato_carac"
    if re.search(r"(acao (foi )?tomada|o que foi feito|tratativa)", n) and re.match(r"(descreva|qual)", n):
        return "livre_acao"
    if n.startswith("descreva") and "anomalia" in n:
        return "livre_anomalia"
    if n == "finalizar fca":
        return "marcador"
    return "arvore"


ACAO_LIKE = re.compile(r"^(\w+(ar|er|ir|or)|verifique|caso|se)$")


def acao_arvore(linha, cols_arvore, hdr, produto_qrz_vapor):
    """Retorna (acao, resultado) segundo a árvore de decisão."""
    preenchidas = [(i, limpar(linha[i])) for i in cols_arvore if limpar(linha[i])]
    if not preenchidas:
        return None, None
    # 1) coluna que recebeu "Finalizar FCA" = última ação executada
    term = [(i, v) for i, v in preenchidas if norm(v).rstrip(".") == "finalizar fca"]
    if term:
        i, v = term[-1]
        return limpar_titulo_acao(hdr[i]), v
    if produto_qrz_vapor:
        # neste Forms o valor da célula é a própria instrução
        for i, v in reversed(preenchidas):
            if ACAO_LIKE.match(norm(v).split(" ")[0]):
                return v.rstrip(".").strip(), None
        i, v = preenchidas[-1]
        return v.rstrip(".").strip(), None
    # 2) senão, a última coluna preenchida à direita
    i, v = preenchidas[-1]
    return limpar_titulo_acao(hdr[i]), v


# --------------------------------------------------------------- leitura
def ler_planilha(caminho, area_key, un_sigla, un_nome, proximo_id):
    ws = openpyxl.load_workbook(caminho, data_only=True).active
    linhas = list(ws.iter_rows(values_only=True))
    if len(linhas) < 2:
        return []
    hdr = [limpar(h) or "" for h in linhas[0]]
    pap = [papel(i, h) for i, h in enumerate(hdr)]
    col = lambda p: [i for i, x in enumerate(pap) if x == p]
    c_turno, c_op, c_sub = col("turno"), col("operador"), col("sub")
    c_prod, c_fato, c_car = col("sub_produto"), col("fato"), col("fato_carac")
    c_lacao, c_lanom = col("livre_acao"), col("livre_anomalia")
    c_arv = col("arvore")
    area_nome = AREAS[area_key]
    setor = f"{area_nome} {un_sigla}"
    saida = []

    for linha in linhas[1:]:
        linha = list(linha) + [None] * (len(hdr) - len(linha))
        if all(v in (None, "") for v in linha):
            continue  # linha totalmente vazia não é FCA
        concl = para_datetime(linha[2]) or para_datetime(linha[1])
        t_raw = limpar(linha[c_turno[0]]) if c_turno else None
        m = re.search(r"(?i)turno\s*([abc])", t_raw or "")
        turno = f"Turno {m.group(1).upper()}" if m else turno_por_hora(concl)

        # operador (QRZ Vapor tem um campo por turno)
        ops = [limpar(linha[i]) for i in c_op if limpar(linha[i])]
        op_bruto = None
        if ops:
            letra = turno[-1] if turno else ""
            pref = [limpar(linha[i]) for i in c_op if limpar(linha[i]) and f'"{letra.lower()}"' in norm(hdr[i])]
            op_bruto = (pref or ops)[0]
        operador = operador_final(op_bruto, un_sigla, turno)

        # sub-processo
        sub = sub_limpo(linha[c_sub[0]]) if c_sub else None
        if c_prod and limpar(linha[c_prod[0]]):
            sub = limpar(linha[c_prod[0]])
        sub = sub or area_nome

        # anomalia / ação em texto livre
        anom = next((limpar(linha[i]) for i in c_lanom if limpar(linha[i])), None)
        acao_livre = next((limpar(linha[i]) for i in c_lacao if limpar(linha[i])), None)

        # fato
        fato = next((limpar(linha[i]) for i in c_fato if limpar(linha[i])), None)
        if not fato and c_car:
            partes = [limpar(linha[i]) for i in c_car if limpar(linha[i])]
            fato = " — ".join(p.rstrip(".") for p in partes) or None
        fato = fato or anom or "Não informado"

        # última ação
        a_arv, resultado = acao_arvore(linha, c_arv, hdr, bool(c_prod))
        acao = acao_livre or a_arv or "Sem ação registrada"
        acao = acao[:1].upper() + acao[1:]

        saida.append({
            "i": proximo_id + len(saida),
            "idf": linha[0],
            "u": un_sigla, "un": un_nome, "a": area_nome, "s": setor,
            "d": concl.strftime("%Y-%m-%d") if concl else "",
            "h": concl.strftime("%H:%M") if concl else "",
            "t": turno or "Turno ?",
            "o": operador, "sp": sub, "f": fato,
            "ac": acao, "at": a_arv, "an": anom,
            "x": 1 if acao_livre else 0,
            "r": resultado,
        })
    return saida


def consolidar(pasta):
    arquivos = sorted(p for p in Path(pasta).glob("*.xlsx") if not p.name.startswith("~$"))
    registros, resumo, vistos = [], [], {}
    for p in arquivos:
        ident = identificar(p.name)
        if not ident:
            print(f"  (ignorado, nome não reconhecido) {p.name}")
            continue
        area_key, (sigla, nome_un) = ident
        if (area_key, sigla) in vistos:
            print(f"  (ignorado, já lido: {vistos[(area_key, sigla)]}) {p.name}")
            continue
        vistos[(area_key, sigla)] = p.name
        recs = ler_planilha(p, area_key, sigla, nome_un, len(registros) + 1)
        registros += recs
        resumo.append((AREAS[area_key], sigla, len(recs), p.name))
    faltam = [f"{AREAS[a]} {u}" for a in AREAS for u in ("CLE", "QRZ")
              if (a, u) not in {(k, s) for (k, s) in vistos}]
    return registros, resumo, faltam


def gravar_csv(registros):
    cols = ["i", "idf", "un", "a", "s", "d", "h", "t", "o", "sp", "f", "ac", "at", "an", "x", "r"]
    with open(SAIDA_CSV, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.writer(fh, delimiter=";")
        w.writerow(["ID", "Id no Forms", "Unidade", "Área", "Setor", "Data", "Hora", "Turno", "Operador",
                    "Sub-processo", "Fato", "Ação (final)", "Ação (árvore)", "Anomalia descrita", "Ação em texto livre?", "Resultado"])
        for r in registros:
            w.writerow([r.get(c, "") if r.get(c) is not None else "" for c in cols])


def gerar_html(registros):
    if TEMPLATE.exists():
        html = TEMPLATE.read_text(encoding="utf-8")
    elif SAIDA_HTML.exists():
        # sem template_fca.html: deriva do próprio index.html (troca os dados e a data por marcadores)
        html = SAIDA_HTML.read_text(encoding="utf-8")
        html, n1 = re.subn(r"const fcaRaw = \[.*?\];\n", "const fcaRaw = /*__FCA_DATA__*/[];\n", html, count=1, flags=re.S)
        html, n2 = re.subn(r"ATUALIZADO = '[^']*';", "ATUALIZADO = '__ATUALIZADO__';", html, count=1)
        if not (n1 and n2):
            sys.exit("Não consegui derivar o template do index.html")
    else:
        sys.exit(f"Não achei o template: {TEMPLATE}")
    dados = json.dumps(registros, ensure_ascii=False, separators=(",", ":"))
    agora = datetime.now().strftime("%d/%m/%Y %H:%M")
    if "/*__FCA_DATA__*/[]" not in html:
        sys.exit("Template sem o marcador /*__FCA_DATA__*/[]")
    html = html.replace("/*__FCA_DATA__*/[]", dados).replace("__ATUALIZADO__", agora)
    SAIDA_HTML.write_text(html, encoding="utf-8")


def publicar():
    pasta = SAIDA_HTML.parent
    def git(*a):
        return subprocess.run(["git", *a], cwd=pasta, capture_output=True, text=True)
    if git("rev-parse", "--is-inside-work-tree").returncode != 0:
        print("  Pasta não é um repositório git — nada publicado.")
        return
    git("add", SAIDA_HTML.name)
    c = git("commit", "-m", f"Atualiza FCA {datetime.now():%d/%m/%Y %H:%M}")
    if "nothing to commit" in (c.stdout + c.stderr):
        print("  Sem mudanças para publicar.")
        return
    p = git("push")
    print("  Publicado no GitHub." if p.returncode == 0 else f"  Falha no push:\n{p.stderr}")


def baixar_planilhas(pasta):
    """Modo automático: baixa as planilhas pelos links anônimos (JSON nome -> link)."""
    import urllib.request
    import requests
    pasta = Path(pasta)
    pasta.mkdir(parents=True, exist_ok=True)
    for f in pasta.glob("*.xlsx"):
        f.unlink()
    links = requests.get(LINKS_URL, timeout=60).json() if LINKS_URL.startswith("http") else json.loads(LINKS_URL)
    falhas = []
    for nome, url in links.items():
        ok = False
        for tentativa in (url, url + ("&" if "?" in url else "?") + "download=1"):
            try:
                r = requests.Session().get(tentativa, timeout=120)
                if r.content[:2] == b"PK":
                    (pasta / nome).write_bytes(r.content)
                    ok = True
                    break
            except Exception:
                pass
        if not ok:
            falhas.append(nome)
    if falhas:
        sys.exit("Falha ao baixar: " + ", ".join(falhas))
    print(f"Baixadas {len(links)} planilhas.")


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    global PASTA_PLANILHAS
    if args:
        PASTA_PLANILHAS = Path(args[0])
    if os.environ.get("GITHUB_ACTIONS") and not LINKS_URL:
        sys.exit("ERRO: o secret FCA_LINKS_URL nao esta cadastrado (Settings > Secrets and variables > Actions).")
    if LINKS_URL:
        PASTA_PLANILHAS = AQUI / "planilhas"
        baixar_planilhas(PASTA_PLANILHAS)
    print(f"Lendo planilhas em: {PASTA_PLANILHAS}")
    registros, resumo, faltam = consolidar(PASTA_PLANILHAS)
    for area, sig, n, nome in resumo:
        print(f"  {area:<22} {sig}  {n:>5} FCAs   <- {nome}")
    print(f"TOTAL: {len(registros)} FCAs")
    if faltam:
        print("  ATENÇÃO — planilhas não encontradas: " + ", ".join(faltam))
    if not registros:
        sys.exit("Nenhum registro lido.")
    gravar_csv(registros)
    gerar_html(registros)
    print(f"Gerado: {SAIDA_HTML}\nBase consolidada (conferência): {SAIDA_CSV}")
    if "--publicar" in sys.argv:
        publicar()


if __name__ == "__main__":
    main()
