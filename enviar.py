#!/usr/bin/env python3
"""Oferta de Consórcio — envio semanal (terça) da "Semana N" do mês.

Cada pasta "Semana N" no Drive tem:
  - a arte (PNG/JPG — usa a mais recente)
  - conteudo.txt  (assunto, título, textos curtos, botão)
Cada semana usa um MODELO de e-mail diferente (1 a 5).

Uso:
  python3 enviar.py --preview [--semana N]           # gera preview_semanaN.html, não envia
  python3 enviar.py --teste EMAIL [--semana N]       # envia 1 e-mail de teste
  python3 enviar.py --disparar                       # BLOQUEADO até aprovação
"""
from __future__ import annotations
import argparse, base64, datetime as dt, html, imaplib, mimetypes, os, smtplib, ssl, sys, urllib.parse
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.utils import formataddr, make_msgid, formatdate
from pathlib import Path

BASE = Path(__file__).resolve().parent
for _l in (BASE / ".env").read_text().splitlines() if (BASE / ".env").exists() else []:
    if "=" in _l:
        k, v = _l.split("=", 1); os.environ.setdefault(k.strip(), v.strip())

PASTA_DRIVE = Path(os.getenv("PASTA_OFERTA", str(Path.home() /
    "Library/CloudStorage/GoogleDrive-silvanopiacentine1@gmail.com/Meu Drive/Oferta de Consórcio")))
SMTP_HOST = "email02.webplusidc.com.br"
SMTP_USER = os.getenv("SMTP_USER", "silvano@piaseg.com.br")
SMTP_PASS = os.getenv("SMTP_PASS", "")
REMETENTE_NOME = "Silvano Piacentine | Piaseg Seguros"
WHATSAPP = "5567984011034"
WHATS_FMT = "(67) 98401-1034"
MSG_WHATS = "Olá! Vi a oferta de consórcio da Piaseg por e-mail e quero saber mais."
ASSINATURA = BASE / "assinatura.png"
IMG_EXT = {".png", ".jpg", ".jpeg", ".gif"}

NAVY, GOLD, GOLD_L, BG, TXT, MUTED = "#072a3c", "#c2a360", "#f3ecdd", "#eef1f4", "#1d2b33", "#5b6b75"
FONT = "Arial,Helvetica,sans-serif"


def semana_do_mes(d: dt.date) -> int:
    return (d.day - 1) // 7 + 1


def ler_conteudo(pasta: Path) -> dict:
    c = {"ASSUNTO": "Consórcio Piaseg: condições especiais desta semana", "TITULO": "", "SUBTITULO": "",
         "TEXTO": "", "ITENS": [], "FECHAMENTO": "", "BOTAO": "Quero simular no WhatsApp", "FAIXA": ""}
    f = pasta / "conteudo.txt"
    if not f.exists():
        return c
    for linha in f.read_text(encoding="utf-8").splitlines():
        s = linha.strip()
        if not s or s.startswith("#"):
            continue
        if s.startswith("- "):
            c["ITENS"].append(s[2:].strip())
        elif ":" in s and s.split(":", 1)[0].strip().upper() in c:
            k, v = s.split(":", 1)
            c[k.strip().upper()] = v.strip()
    return c


def arte_da_semana(n: int, arte_reserva: Path | None = None):
    pasta = PASTA_DRIVE / f"Semana {n}"
    imgs = sorted([p for p in pasta.glob("*") if p.suffix.lower() in IMG_EXT],
                  key=lambda p: p.stat().st_mtime, reverse=True)
    if not imgs and not arte_reserva:
        raise SystemExit(f"Sem arte em {pasta} — nada enviado.")
    return (imgs[0] if imgs else arte_reserva), ler_conteudo(pasta)


# ---------- blocos reutilizáveis ----------
def e(s): return html.escape(s or "")


def botao(c, link, cor="#25D366"):
    return f"""<table role="presentation" cellpadding="0" cellspacing="0" align="center"><tr>
<td align="center" bgcolor="{cor}" style="border-radius:999px;">
<a href="{link}" target="_blank" style="display:inline-block;padding:16px 34px;font-family:{FONT};font-size:18px;font-weight:bold;color:#ffffff;text-decoration:none;border-radius:999px;">💬&nbsp; {e(c['BOTAO'])}</a>
</td></tr></table>
<p style="margin:12px 0 0;font-family:{FONT};font-size:14px;color:{MUTED};text-align:center;">WhatsApp {WHATS_FMT}</p>"""


def arte_img(link, alt):
    return f'<a href="{link}" target="_blank"><img src="cid:arte" width="600" alt="{e(alt)}" style="display:block;width:100%;max-width:600px;height:auto;border:0;"></a>'


def itens_check(itens, cor=GOLD):
    linhas = "".join(f"""<tr><td valign="top" style="padding:6px 10px 6px 0;font-family:{FONT};font-size:18px;color:{cor};font-weight:bold;">✔</td>
<td style="padding:6px 0;font-family:{FONT};font-size:16px;line-height:1.45;color:{TXT};">{e(i)}</td></tr>""" for i in itens)
    return f'<table role="presentation" cellpadding="0" cellspacing="0" style="margin:0 auto;">{linhas}</table>' if itens else ""


def p(texto, size=16, cor=TXT, peso="normal", align="center", mb=14):
    return f'<p style="margin:0 0 {mb}px;font-family:{FONT};font-size:{size}px;line-height:1.5;color:{cor};font-weight:{peso};text-align:{align};">{e(texto)}</p>' if texto else ""


# ---------- 5 modelos ----------
def modelo_1(c, link):  # DESTAQUE — faixa navy com título dourado
    return f"""
<tr><td bgcolor="{NAVY}" align="center" style="padding:34px 28px 28px;">
 {p(c['FAIXA'], 13, GOLD, 'bold', mb=10)}
 <h1 style="margin:0 0 10px;font-family:{FONT};font-size:30px;line-height:1.2;color:{GOLD};">{e(c['TITULO'])}</h1>
 {p(c['SUBTITULO'], 17, '#ffffff', mb=0)}
</td></tr>
<tr><td>{arte_img(link, c['ASSUNTO'])}</td></tr>
<tr><td style="padding:28px 32px 6px;">{p(c['TEXTO'])}{itens_check(c['ITENS'])}</td></tr>
<tr><td align="center" style="padding:14px 28px 6px;">{p(c['FECHAMENTO'], 16, NAVY, 'bold')}{botao(c, link)}</td></tr>"""


def modelo_2(c, link):  # CARTA — texto pessoal primeiro, arte depois
    paragrafos = "".join(p(i, 16, TXT, align="left") for i in c["ITENS"])
    return f"""
<tr><td style="padding:34px 34px 10px;">
 {p(c['TITULO'], 22, NAVY, 'bold', 'left', 16)}
 {p(c['SUBTITULO'], 16, TXT, align='left')}
 {p(c['TEXTO'], 16, TXT, align='left')}
 {paragrafos}
 {p(c['FECHAMENTO'], 16, NAVY, 'bold', 'left', 4)}
</td></tr>
<tr><td style="padding:10px 0 0;">{arte_img(link, c['ASSUNTO'])}</td></tr>
<tr><td align="center" style="padding:26px 28px 6px;">{botao(c, link)}</td></tr>"""


def modelo_3(c, link):  # URGÊNCIA — faixa dourada + contador de prazo
    return f"""
<tr><td bgcolor="{GOLD}" align="center" style="padding:12px 20px;font-family:{FONT};font-size:14px;font-weight:bold;letter-spacing:1px;color:{NAVY};">⏰ {e(c['FAIXA'] or 'OFERTA POR TEMPO LIMITADO')}</td></tr>
<tr><td align="center" style="padding:30px 28px 18px;">
 <h1 style="margin:0 0 10px;font-family:{FONT};font-size:32px;line-height:1.15;color:{NAVY};">{e(c['TITULO'])}</h1>
 {p(c['SUBTITULO'], 17, MUTED, mb=0)}
</td></tr>
<tr><td>{arte_img(link, c['ASSUNTO'])}</td></tr>
<tr><td style="padding:26px 32px 4px;">{p(c['TEXTO'])}{itens_check(c['ITENS'], '#c0392b')}</td></tr>
<tr><td align="center" style="padding:10px 28px 6px;">
 <table role="presentation" cellpadding="0" cellspacing="0" align="center"><tr><td bgcolor="#fdecea" style="border-radius:10px;padding:12px 18px;font-family:{FONT};font-size:16px;font-weight:bold;color:#c0392b;">{e(c['FECHAMENTO'])}</td></tr></table>
 <div style="height:18px;"></div>{botao(c, link)}</td></tr>"""


def modelo_4(c, link):  # BENEFÍCIOS — cards antes da arte
    cards = ""
    for item in c["ITENS"][:3]:
        tit, _, desc = item.partition("|")
        cards += f"""<td width="33%" valign="top" style="padding:6px;">
<table role="presentation" width="100%" cellpadding="0" cellspacing="0"><tr><td bgcolor="{GOLD_L}" style="border-radius:12px;padding:18px 12px;text-align:center;font-family:{FONT};">
<div style="font-size:17px;font-weight:bold;color:{NAVY};margin-bottom:6px;">{e(tit.strip())}</div>
<div style="font-size:14px;line-height:1.4;color:{TXT};">{e(desc.strip())}</div></td></tr></table></td>"""
    return f"""
<tr><td align="center" style="padding:34px 28px 10px;">
 <h1 style="margin:0 0 10px;font-family:{FONT};font-size:28px;line-height:1.2;color:{NAVY};">{e(c['TITULO'])}</h1>
 {p(c['SUBTITULO'], 17, MUTED, mb=0)}
</td></tr>
<tr><td style="padding:12px 18px 22px;"><table role="presentation" width="100%" cellpadding="0" cellspacing="0"><tr>{cards}</tr></table></td></tr>
<tr><td>{arte_img(link, c['ASSUNTO'])}</td></tr>
<tr><td align="center" style="padding:26px 30px 6px;">{p(c['TEXTO'])}{p(c['FECHAMENTO'], 16, NAVY, 'bold')}{botao(c, link)}</td></tr>"""


def modelo_5(c, link):  # PERGUNTA — pergunta grande + respostas curtas
    respostas = "".join(f"""<tr><td style="padding:0 0 10px;"><table role="presentation" width="100%" cellpadding="0" cellspacing="0"><tr>
<td width="4" bgcolor="{GOLD}"></td><td style="padding:10px 14px;font-family:{FONT};font-size:16px;line-height:1.45;color:{TXT};background:#f7f9fa;">{e(i)}</td></tr></table></td></tr>""" for i in c["ITENS"])
    return f"""
<tr><td bgcolor="{NAVY}" align="center" style="padding:40px 30px;">
 <h1 style="margin:0;font-family:{FONT};font-size:30px;line-height:1.25;color:#ffffff;">{e(c['TITULO'])}</h1>
 {('<p style="margin:12px 0 0;font-family:%s;font-size:17px;color:%s;">%s</p>' % (FONT, GOLD, e(c['SUBTITULO']))) if c['SUBTITULO'] else ''}
</td></tr>
<tr><td style="padding:28px 32px 10px;">{p(c['TEXTO'], 17, NAVY, 'bold')}<table role="presentation" width="100%" cellpadding="0" cellspacing="0">{respostas}</table></td></tr>
<tr><td>{arte_img(link, c['ASSUNTO'])}</td></tr>
<tr><td align="center" style="padding:26px 28px 6px;">{p(c['FECHAMENTO'], 16, NAVY, 'bold')}{botao(c, link)}</td></tr>"""


MODELOS = {1: modelo_1, 2: modelo_2, 3: modelo_3, 4: modelo_4, 5: modelo_5}


def montar(dest, arte: Path, c: dict, n: int, prefixo=""):
    link = f"https://wa.me/{WHATSAPP}?text={urllib.parse.quote(MSG_WHATS)}"
    sair = f"mailto:{SMTP_USER}?subject=SAIR&body={urllib.parse.quote('Não quero mais receber ofertas de consórcio.')}"
    assunto = prefixo + c["ASSUNTO"]
    preheader = c["SUBTITULO"] or c["TITULO"] or c["ASSUNTO"]
    ass = ('<tr><td style="padding:26px 24px 24px;"><img src="cid:assinatura" width="552" alt="Silvano Piacentine — Piaseg Seguros" '
           'style="display:block;width:100%;max-width:552px;height:auto;border:0;"></td></tr>') if ASSINATURA.exists() else ""
    corpo = f"""<!doctype html><html lang="pt-BR"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1"><title>{e(assunto)}</title></head>
<body style="margin:0;padding:0;background:{BG};">
<div style="display:none;max-height:0;overflow:hidden;opacity:0;">{e(preheader)}</div>
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:{BG};">
<tr><td align="center" style="padding:24px 10px;">
<table role="presentation" width="600" cellpadding="0" cellspacing="0" style="max-width:600px;width:100%;background:#ffffff;border-radius:12px;overflow:hidden;">
{MODELOS[n](c, link)}
{ass}
</table>
<p style="font-family:{FONT};font-size:11px;color:#8a979f;max-width:560px;margin:16px auto 0;line-height:1.5;text-align:center;">
Você recebe este e-mail por ter relacionamento com a Piaseg Seguros. Não quer mais receber ofertas de consórcio?
<a href="{sair}" style="color:#8a979f;">Clique aqui para sair da lista</a>.</p>
</td></tr></table></body></html>"""
    texto = "\n".join(x for x in [c["TITULO"], c["SUBTITULO"], c["TEXTO"], *[f"- {i.replace('|', ':')}" for i in c["ITENS"]],
                                    c["FECHAMENTO"], f"\nFale no WhatsApp: {link}", WHATS_FMT,
                                    "\nSilvano Piacentine — Piaseg Seguros", "Para sair da lista, responda com SAIR."] if x)

    msg = MIMEMultipart("related")
    msg["Subject"] = assunto
    msg["From"] = formataddr((REMETENTE_NOME, SMTP_USER))
    msg["To"] = dest
    msg["Date"] = formatdate(localtime=True)
    msg["Message-ID"] = make_msgid(domain="piaseg.com.br")
    msg["List-Unsubscribe"] = f"<mailto:{SMTP_USER}?subject=SAIR>"
    alt = MIMEMultipart("alternative")
    alt.attach(MIMEText(texto, "plain", "utf-8"))
    alt.attach(MIMEText(corpo, "html", "utf-8"))
    msg.attach(alt)
    sub = (mimetypes.guess_type(arte.name)[0] or "image/jpeg").split("/")[1]
    img = MIMEImage(arte.read_bytes(), _subtype=sub)
    img.add_header("Content-ID", "<arte>")
    img.add_header("Content-Disposition", "inline", filename="oferta-consorcio" + arte.suffix.lower())
    msg.attach(img)
    if ASSINATURA.exists():
        a = MIMEImage(ASSINATURA.read_bytes(), _subtype="png")
        a.add_header("Content-ID", "<assinatura>")
        a.add_header("Content-Disposition", "inline", filename="assinatura.png")
        msg.attach(a)
    return msg, corpo


def preview_html(corpo, arte: Path):
    mime = mimetypes.guess_type(arte.name)[0] or "image/jpeg"
    corpo = corpo.replace("cid:arte", f"data:{mime};base64," + base64.b64encode(arte.read_bytes()).decode())
    if ASSINATURA.exists():
        corpo = corpo.replace("cid:assinatura", "data:image/png;base64," + base64.b64encode(ASSINATURA.read_bytes()).decode())
    return corpo


def enviar(msgs):
    with smtplib.SMTP_SSL(SMTP_HOST, 465, context=ssl.create_default_context()) as s:
        s.login(SMTP_USER, SMTP_PASS)
        for m in msgs:
            s.send_message(m)
            print("enviado ->", m["To"], "|", m["Subject"])
    try:
        with imaplib.IMAP4_SSL(SMTP_HOST, 993) as im:
            im.login(SMTP_USER, SMTP_PASS)
            for m in msgs:
                im.append('"Sent Items"', "(\\Seen)", None, m.as_bytes())
    except Exception as ex:
        print("aviso: não salvou cópia em Sent Items:", ex)


def main():
    ap = argparse.ArgumentParser()
    g = ap.add_mutually_exclusive_group(required=True)
    g.add_argument("--teste", metavar="EMAIL")
    g.add_argument("--preview", action="store_true")
    g.add_argument("--disparar", action="store_true")
    ap.add_argument("--semana", type=int, nargs="*", help="Semana(s) N; padrão = semana do mês de hoje")
    ap.add_argument("--como-cliente", action="store_true", help="teste sem o prefixo [TESTE] no assunto")
    ap.add_argument("--arte-reserva", help="arte usada se a pasta da semana ainda não tiver arte (só teste/preview)")
    a = ap.parse_args()
    semanas = a.semana or [semana_do_mes(dt.date.today())]
    reserva = Path(a.arte_reserva) if a.arte_reserva else None

    if a.disparar:
        sys.exit("Disparo para a lista ainda não liberado — aguardando aprovação do teste.")
    msgs = []
    for n in semanas:
        arte, c = arte_da_semana(n, reserva)
        print(f"Semana {n} | modelo {n} | arte: {arte.name} | assunto: {c['ASSUNTO']}")
        if a.preview:
            _, corpo = montar("preview@exemplo.com", arte, c, n)
            out = BASE / f"preview_semana{n}.html"
            out.write_text(preview_html(corpo, arte), encoding="utf-8")
            print("  preview:", out)
        else:
            msg, _ = montar(a.teste, arte, c, n, prefixo="" if a.como_cliente else f"[TESTE Semana {n}] ")
            msgs.append(msg)
    if msgs:
        if not SMTP_PASS:
            sys.exit("Defina SMTP_PASS no .env")
        enviar(msgs)


if __name__ == "__main__":
    main()
