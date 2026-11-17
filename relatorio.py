"""Relatório de envios da Oferta de Consórcio (Excel no Drive).

- Devolvidos: lê avisos de não entrega (bounces) no Inbox/Lixo do silvano@, cruza com quem recebeu,
  grava em Controle/devolvidos.txt e move os avisos para a pasta IMAP "Consórcio - Devolvidos".
- Aberturas: planilha Google "aberturas" (preenchida pelo pixel na Vercel).
- Gera "Relatório de Envios.xlsx": Resumo, Corrigir e-mails, e uma aba por disparo.
"""
from __future__ import annotations
import datetime as dt, email, imaplib, io, json, os, re
from collections import defaultdict
from email.header import decode_header
from email.utils import parseaddr, parsedate_to_datetime

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

PASTA_LIXEIRA = "Deleted Items"  # avisos de devolução dos envios vão para a lixeira (ficam só no relatório)
RE_EMAIL_TXT = re.compile(r"[A-Za-z0-9._%+'-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}")
REMETENTES_BOUNCE = ("mailer-daemon", "postmaster", "system administrator", "mail delivery", "microsoftexchange")
ASSUNTOS_BOUNCE = ("undeliver", "delivery", "não entregue", "nao entregue", "returned", "falha na entrega",
                   "retorno ao remetente", "failure", "problema ao entregar", "não foi possível entregar", "devolvid")


def _dec(s):
    return "".join(t.decode(c or "utf-8", "replace") if isinstance(t, bytes) else t for t, c in decode_header(s or ""))


def _imap_utf7(s: str) -> str:
    """Codifica nome de pasta IMAP (modified UTF-7)."""
    out, buf = [], ""
    def flush():
        nonlocal buf
        if buf:
            import base64
            b = base64.b64encode(buf.encode("utf-16-be")).decode().rstrip("=").replace("/", ",")
            out.append("&" + b + "-"); buf = ""
    for ch in s:
        if 0x20 <= ord(ch) <= 0x7e:
            flush(); out.append("&-" if ch == "&" else ch)
        else:
            buf += ch
    flush()
    return "".join(out)


def _motivo(msg) -> str:
    txt = ""
    for p in msg.walk():
        ct = p.get_content_type()
        if ct in ("message/delivery-status",):
            txt += str(p.get_payload())
        elif ct == "text/plain":
            try:
                txt += p.get_payload(decode=True).decode(p.get_content_charset() or "utf-8", "replace")
            except Exception:
                pass
    m = re.search(r"Diagnostic-Code:\s*(.+)", txt, re.I) or re.search(r"\b(5\d\d[ \-]\d\.\d\.\d[^\n]*)", txt) \
        or re.search(r"\b(55\d[^\n]{5,})", txt)
    if m:
        return re.sub(r"\s+", " ", m.group(1)).strip()[:180]
    for k, v in (("user unknown", "usuário inexistente"), ("does not exist", "endereço não existe"),
                 ("mailbox full", "caixa cheia"), ("quota", "caixa cheia"), ("not found", "endereço não encontrado")):
        if k in txt.lower():
            return v
    return "endereço não entregue"


def coletar_devolucoes(host, user, pwd, enviados_por_data: dict[str, set], ja_devolvidos: dict, desde: dt.date, log=print):
    """Retorna novos devolvidos {(data, email): motivo}. Move os avisos para a pasta IMAP de devolvidos."""
    novos = {}
    todos_enviados = defaultdict(list)
    for d, s in enviados_por_data.items():
        for e in s:
            todos_enviados[e].append(d)
    with imaplib.IMAP4_SSL(host, 993) as im:
        im.login(user, pwd)
        pasta_dest = PASTA_LIXEIRA
        for pasta in ("Inbox", "Junk E-Mail", "Junk"):
            if im.select(f'"{pasta}"')[0] != "OK":
                continue
            _, d = im.uid("SEARCH", None, "SINCE", desde.strftime("%d-%b-%Y"))
            mover = []
            for uid in (d[0] or b"").split():
                _, h = im.uid("FETCH", uid, "(BODY.PEEK[HEADER.FIELDS (FROM SUBJECT)])")
                hh = email.message_from_bytes(h[0][1])
                frm, assunto = _dec(hh.get("From", "")).lower(), _dec(hh.get("Subject", "")).lower()
                if not (any(k in frm for k in REMETENTES_BOUNCE) or any(k in assunto for k in ASSUNTOS_BOUNCE)):
                    continue
                _, m = im.uid("FETCH", uid, "(BODY.PEEK[])")
                raw = m[0][1]
                msg = email.message_from_bytes(raw)
                try:
                    quando = parsedate_to_datetime(msg.get("Date")).date()
                except Exception:
                    quando = dt.date.today()
                texto = raw.decode("utf-8", "replace")
                for p in msg.walk():
                    if p.get_content_maintype() == "text":
                        try:
                            texto += p.get_payload(decode=True).decode(p.get_content_charset() or "utf-8", "replace")
                        except Exception:
                            pass
                achados = {e.lower().strip(".") for e in RE_EMAIL_TXT.findall(texto)} & set(todos_enviados)
                achados.discard(user.lower())
                if not achados:
                    continue
                motivo = _motivo(msg)
                for e in achados:
                    datas = sorted(x for x in todos_enviados[e] if x <= quando.isoformat()) or sorted(todos_enviados[e])
                    chave = (datas[-1], e)
                    if chave not in ja_devolvidos:
                        novos[chave] = motivo
                mover.append(uid)
            for uid in mover:
                if im.uid("COPY", uid, f'"{pasta_dest}"')[0] == "OK":
                    im.uid("STORE", uid, "+FLAGS", "(\\Deleted)")
            if mover:
                im.expunge()
                log(f"{len(mover)} avisos de devolução apagados de {pasta} (registrados no relatório)")
    return novos


def gerar_excel(historico: list[list[str]], enviados_por_data: dict[str, list[str]], devolvidos: dict,
                aberturas: dict, nomes: dict, descadastrados: set) -> bytes:
    NAVY, GOLD = "072A3C", "C2A360"
    h_font, h_fill = Font(bold=True, color="FFFFFF"), PatternFill("solid", fgColor=NAVY)
    verde, verm, amar = PatternFill("solid", fgColor="E3F4E8"), PatternFill("solid", fgColor="FBE3E1"), PatternFill("solid", fgColor="FFF4D6")

    def cab(ws, cols, larg):
        ws.append(cols)
        for i, c in enumerate(ws[ws.max_row], 1):
            c.font, c.fill, c.alignment = h_font, h_fill, Alignment(vertical="center", wrap_text=True)
            ws.column_dimensions[get_column_letter(i)].width = larg[i - 1]
        ws.freeze_panes = f"A{ws.max_row + 1}"

    wb = Workbook()
    rs = wb.active; rs.title = "Resumo"
    rs.append(["Relatório de envios — Oferta de Consórcio Piaseg"]); rs["A1"].font = Font(bold=True, size=14, color=NAVY)
    rs.append([f"Atualizado em {dt.datetime.now(dt.timezone(dt.timedelta(hours=-4))):%d/%m/%Y %H:%M}. "
               "Abertura é aproximada: alguns programas bloqueiam imagens (não contam) e o iPhone/Apple Mail pode marcar como aberto sozinho."])
    rs.append([])
    cab(rs, ["Data", "Semana", "Arte", "Enviados", "Entregues", "Devolvidos (endereço errado)", "Abertos", "% abertura (sobre entregues)", "Saíram da lista (total)", "Status"],
        [12, 9, 34, 11, 11, 16, 10, 14, 14, 12])
    for linha in historico:
        d, sem, arte, _md5, status, _tot = (linha + [""] * 6)[:6]
        env = enviados_por_data.get(d, [])
        dev = sum(1 for e in env if (d, e) in devolvidos)
        ab = sum(1 for e in env if (d, e) in aberturas)
        entreg = len(env) - dev
        mede = d >= "2026-11-18"
        rs.append([dt.date.fromisoformat(d).strftime("%d/%m/%Y"), int(sem) if sem.isdigit() else sem, arte, len(env), entreg, dev,
                   ab if mede else "não medido", (f"{ab / entreg:.0%}" if entreg else "-") if mede else "-",
                   sum(1 for e in env if e in descadastrados), status])

    # Corrigir e-mails (todos os devolvidos, sem duplicar)
    cz = wb.create_sheet("Corrigir e-mails")
    cab(cz, ["Nome (da lista)", "E-mail com problema", "Devolvido no envio de", "Motivo informado pelo servidor"], [38, 36, 14, 70])
    vistos = set()
    for (d, e), mot in sorted(devolvidos.items(), key=lambda x: (x[0][1])):
        if e in vistos:
            continue
        vistos.add(e)
        cz.append([nomes.get(e, ""), e, dt.date.fromisoformat(d).strftime("%d/%m/%Y"), mot])
    cz.auto_filter.ref = cz.dimensions

    for linha in reversed(historico):
        d, sem = linha[0], linha[1]
        ws = wb.create_sheet(re.sub(r"[\\/*?:\[\]]", "", f"{dt.date.fromisoformat(d):%d-%m} Semana {sem}")[:31])
        cab(ws, ["Nome", "E-mail", "Situação", "Abriu?", "Primeira abertura", "Nº de aberturas", "Motivo da devolução"],
            [36, 36, 26, 9, 17, 11, 60])
        mede = d >= "2026-11-18"
        for e in enviados_por_data.get(d, []):
            dev = devolvidos.get((d, e))
            ab = aberturas.get((d, e))
            situ = "Devolvido – endereço errado" if dev else ("Entregue – saiu da lista" if e in descadastrados else "Entregue")
            abriu = ("Sim" if ab else "Não") if mede else "não medido"
            ws.append([nomes.get(e, ""), e, situ, "" if dev else abriu,
                       ab[0].strftime("%d/%m/%Y %H:%M") if ab else "", ab[1] if ab else "", dev or ""])
            fill = verm if dev else (verde if ab else None)
            if fill:
                for c in ws[ws.max_row]:
                    c.fill = fill
        ws.auto_filter.ref = ws.dimensions
    buf = io.BytesIO(); wb.save(buf)
    return buf.getvalue()


def ler_aberturas(blob_token: str, log=print) -> dict:
    """{(data, email): (primeira_abertura_local, n)} — lê os caminhos gravados pelo pixel no Vercel Blob."""
    import base64, urllib.request, urllib.parse
    if not blob_token:
        log("aviso: BLOB_READ_WRITE_TOKEN ausente — aberturas não lidas"); return {}
    fuso = dt.timezone(dt.timedelta(hours=-4))
    out, cursor = {}, None
    try:
        while True:
            q = {"prefix": "aberturas/", "limit": "1000"}
            if cursor:
                q["cursor"] = cursor
            req = urllib.request.Request("https://blob.vercel-storage.com?" + urllib.parse.urlencode(q),
                                         headers={"Authorization": f"Bearer {blob_token}", "x-api-version": "11"})
            d = json.loads(urllib.request.urlopen(req, timeout=60).read())
            for b in d.get("blobs", []):
                partes = b["pathname"].split("/")
                if len(partes) < 4:
                    continue
                data, em = partes[1], partes[2]
                try:
                    email_ = base64.urlsafe_b64decode(em + "=" * (-len(em) % 4)).decode().lower()
                    t = dt.datetime.fromisoformat(b["uploadedAt"].replace("Z", "+00:00")).astimezone(fuso).replace(tzinfo=None)
                except Exception:
                    continue
                k = (data, email_)
                out[k] = (min(out[k][0], t), out[k][1] + 1) if k in out else (t, 1)
            cursor = d.get("cursor")
            if not d.get("hasMore") or not cursor:
                break
    except Exception as ex:
        log(f"aviso: não consegui ler aberturas: {ex}")
    return out


def descadastros_web(blob_token: str, log=print) -> set:
    """E-mails que clicaram em "sair da lista" (página de um clique na Vercel)."""
    import base64, urllib.request, urllib.parse
    out, cursor = set(), None
    if not blob_token:
        return out
    try:
        while True:
            q = {"prefix": "descadastros/", "limit": "1000"}
            if cursor:
                q["cursor"] = cursor
            req = urllib.request.Request("https://blob.vercel-storage.com?" + urllib.parse.urlencode(q),
                                         headers={"Authorization": f"Bearer {blob_token}", "x-api-version": "11"})
            d = json.loads(urllib.request.urlopen(req, timeout=60).read())
            for b in d.get("blobs", []):
                p = b["pathname"].split("/")
                if len(p) >= 3:
                    try:
                        out.add(base64.urlsafe_b64decode(p[1] + "=" * (-len(p[1]) % 4)).decode().lower())
                    except Exception:
                        pass
            cursor = d.get("cursor")
            if not d.get("hasMore") or not cursor:
                break
    except Exception as ex:
        log(f"aviso: não consegui ler descadastros da página: {ex}")
    return out
