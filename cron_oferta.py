#!/usr/bin/env python3
"""Oferta de Consórcio — robô na nuvem (Render cron, toda terça 13:00 UTC = 9h MS).

Lê tudo da pasta "Oferta de Consórcio" no Google Drive (compartilhada com a conta de serviço):
  Semana N/   -> arte mais recente + conteudo.txt
  Lista de e-mail.XLS
  Controle (não mexer)/ agenda.txt, envios.txt, historico.txt, descadastrados.txt, log_ultimo_disparo.txt, assinatura.png
Os arquivos de controle precisam existir (a conta de serviço só ATUALIZA, não cria arquivos).

Env: GOOGLE_SERVICE_ACCOUNT_JSON, SMTP_USER, SMTP_PASS
     TESTE_PARA="a@b,c@d"  -> ensaio: envia só para esses, não grava envios/histórico
     FORCAR_DATA=AAAA-MM-DD -> simula outro dia (testes)
"""
from __future__ import annotations
import datetime as dt, email, hashlib, imaplib, io, json, os, re, smtplib, ssl, sys, tempfile, time, traceback
from email.header import decode_header
from email.mime.text import MIMEText
from email.utils import parseaddr, formataddr, formatdate
from pathlib import Path

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload

import enviar as E
import relatorio as R

PASTA_RAIZ = "Oferta de Consórcio"
PASTA_CTRL = "Controle (não mexer)"
FUSO = dt.timezone(dt.timedelta(hours=-4))  # Mato Grosso do Sul, sem horário de verão
INTERVALO = 24.0           # 150 e-mails/hora
MAX_FALHAS_SEGUIDAS = 8
SALVAR_A_CADA = 25
AVISAR = os.getenv("AVISAR", "silvano@piaseg.com.br")
RE_EMAIL = re.compile(r"^[^@\s;,]+@[^@\s;,]+\.[a-z]{2,}$")
IMG = ("image/jpeg", "image/png", "image/gif")

TESTE_PARA = [x.strip().lower() for x in os.getenv("TESTE_PARA", "").split(",") if x.strip()]
LOG: list[str] = []
RESUMO_ULTIMO = ""


def log(t):
    linha = f"{dt.datetime.now(FUSO):%d/%m %H:%M:%S} {t}"
    print(linha, flush=True)
    LOG.append(linha)


# ---------------- Drive ----------------
class Drive:
    def __init__(self):
        info = json.loads(os.environ["GOOGLE_SERVICE_ACCOUNT_JSON"])
        cr = service_account.Credentials.from_service_account_info(info, scopes=["https://www.googleapis.com/auth/drive"])
        self.s = build("drive", "v3", credentials=cr, cache_discovery=False)

    def listar(self, q):
        out, tok = [], None
        while True:
            r = self.s.files().list(q=q + " and trashed=false", pageToken=tok, pageSize=200,
                                    fields="nextPageToken,files(id,name,mimeType,modifiedTime,md5Checksum)").execute()
            out += r["files"]; tok = r.get("nextPageToken")
            if not tok:
                return out

    def pasta(self, nome, pai=None):
        q = f"name='{nome}' and mimeType='application/vnd.google-apps.folder'" + (f" and '{pai}' in parents" if pai else "")
        r = self.listar(q)
        if not r:
            raise RuntimeError(f"Pasta não encontrada no Drive: {nome}")
        return r[0]["id"]

    def filhos(self, pai):
        return {f["name"]: f for f in self.listar(f"'{pai}' in parents")}

    def baixar(self, fid) -> bytes:
        buf = io.BytesIO()
        dl = MediaIoBaseDownload(buf, self.s.files().get_media(fileId=fid))
        done = False
        while not done:
            _, done = dl.next_chunk()
        return buf.getvalue()

    def texto(self, fid) -> str:
        return self.baixar(fid).decode("utf-8", "replace")

    def gravar_bin(self, fid, dados: bytes, mime: str):
        media = MediaIoBaseUpload(io.BytesIO(dados), mimetype=mime, resumable=False)
        self.s.files().update(fileId=fid, media_body=media).execute()

    def gravar(self, fid, conteudo: str, tentativas=3):
        for t in range(tentativas):
            try:
                media = MediaIoBaseUpload(io.BytesIO(conteudo.encode("utf-8")), mimetype="text/plain", resumable=False)
                self.s.files().update(fileId=fid, media_body=media).execute()
                return
            except Exception as ex:
                if t == tentativas - 1:
                    print("aviso: falha ao gravar no Drive:", ex)
                time.sleep(5)


# ---------------- e-mail ----------------
def avisar(assunto, corpo):
    try:
        m = MIMEText(corpo, "plain", "utf-8")
        m["Subject"], m["From"], m["To"], m["Date"] = assunto, formataddr(("Oferta de Consórcio (robô)", E.SMTP_USER)), AVISAR, formatdate(localtime=True)
        with smtplib.SMTP_SSL(E.SMTP_HOST, 465, context=ssl.create_default_context(), timeout=60) as s:
            s.login(E.SMTP_USER, E.SMTP_PASS); s.send_message(m)
    except Exception as ex:
        print("aviso: não consegui enviar o aviso:", ex)


def descadastros_imap() -> set[str]:
    sair = set()
    try:
        with imaplib.IMAP4_SSL(E.SMTP_HOST, 993) as im:
            im.login(E.SMTP_USER, E.SMTP_PASS)
            im.select('"Inbox"', readonly=True)
            _, d = im.search(None, "SUBJECT", "SAIR")
            for i in (d[0] or b"").split():
                _, m = im.fetch(i, "(BODY.PEEK[HEADER.FIELDS (FROM SUBJECT)])")
                h = email.message_from_bytes(m[0][1])
                assunto = "".join(t.decode(cs or "utf-8", "replace") if isinstance(t, bytes) else t
                                  for t, cs in decode_header(h.get("Subject", "")))
                if re.search(r"\bSAIR\b", assunto, re.I):
                    sair.add(parseaddr(h.get("From", ""))[1].lower())
    except Exception as ex:
        log(f"aviso: não consegui ler descadastros no Inbox: {ex}")
    sair.discard(E.SMTP_USER.lower())
    return sair


def linhas(txt):
    return [l for l in txt.splitlines() if l.strip() and not l.lstrip().startswith("#")]


# ---------------- relatório ----------------
def ler_lista(dr, itens_raiz):
    import pandas as pd
    nome_lista = next((n for n in itens_raiz if n.lower().startswith("lista de e-mail")), None)
    if not nome_lista:
        raise RuntimeError("Arquivo 'Lista de e-mail' não encontrado")
    df = pd.read_excel(io.BytesIO(dr.baixar(itens_raiz[nome_lista]["id"])))
    col = next(cn for cn in df.columns if "mail" in str(cn).lower())
    coln = next((cn for cn in df.columns if "nome" in str(cn).lower()), None)
    lista, nomes = [], {}
    for _, r in df.iterrows():
        v = r[col]
        if not isinstance(v, str):
            continue
        m = v.strip().lower()
        if RE_EMAIL.match(m) and m not in nomes:
            nomes[m] = str(r[coln]).strip() if coln is not None and isinstance(r[coln], str) else ""
            lista.append(m)
    return lista, nomes


def ler_devolvidos(txt):
    out = {}
    for l in linhas(txt):
        p = l.split("\t")
        if len(p) >= 2:
            out[(p[0], p[1].strip().lower())] = p[2] if len(p) > 2 else ""
    return out


def atualizar_relatorio(dr, ctrl, itens_raiz):
    """Lê devoluções + aberturas e regrava 'Relatório de Envios.xlsx'. Nunca derruba o envio."""
    try:
        nome_rel = next((n for n in ctrl if n.lower().startswith("relatório de envios")), None)
        if not nome_rel or "devolvidos.txt" not in ctrl:
            log("aviso: relatório ou devolvidos.txt ausente na pasta Controle — relatório não gerado"); return set()
        envios = {}
        for l in linhas(dr.texto(ctrl["envios.txt"]["id"])):
            d, _, e = l.partition("\t")
            envios.setdefault(d, {})[e.strip().lower()] = None
        enviados = {d: list(v) for d, v in envios.items()}
        historico = [l.split("\t") for l in linhas(dr.texto(ctrl["historico.txt"]["id"]))]
        datas_hist = {h[0] for h in historico}
        agenda = {k.strip(): v.strip() for k, _, v in (l.partition("=") for l in linhas(dr.texto(ctrl["agenda.txt"]["id"])))}
        for d in sorted(enviados):  # disparo sem linha no histórico (ex.: em andamento)
            if d not in datas_hist:
                sem = agenda.get(d, "")
                sem = sem if sem.isdigit() else str(E.semana_do_mes(dt.date.fromisoformat(d)))
                historico.append([d, sem, "", "", "em andamento", str(len(enviados[d]))])
        historico.sort(key=lambda h: h[0])
        f_dev = ctrl["devolvidos.txt"]["id"]
        dev_txt = dr.texto(f_dev)
        devolvidos = ler_devolvidos(dev_txt)
        if enviados:
            desde = max(dt.date.fromisoformat(min(enviados)), dt.date.today() - dt.timedelta(days=45))
            novos = R.coletar_devolucoes(E.SMTP_HOST, E.SMTP_USER, E.SMTP_PASS,
                                         {d: set(v) for d, v in enviados.items()}, devolvidos, desde, log)
            if novos:
                dev_txt = dev_txt.rstrip("\n") + "\n" + "\n".join(f"{d}\t{e}\t{m}" for (d, e), m in sorted(novos.items())) + "\n"
                dr.gravar(f_dev, dev_txt)
                devolvidos.update(novos)
                log(f"{len(novos)} novos e-mails devolvidos registrados")
        aberturas = R.ler_aberturas(os.getenv("BLOB_READ_WRITE_TOKEN", ""), log)
        _, nomes = ler_lista(dr, itens_raiz)
        f_desc = ctrl["descadastrados.txt"]["id"]
        desc_txt = dr.texto(f_desc)
        sair = {l.strip().lower() for l in linhas(desc_txt)}
        novos_sair = (descadastros_imap() | R.descadastros_web(os.getenv("BLOB_READ_WRITE_TOKEN", ""), log)) - sair
        if novos_sair:
            dr.gravar(f_desc, desc_txt.rstrip("\n") + "\n" + "\n".join(sorted(novos_sair)) + "\n")
            sair |= novos_sair
            log(f"{len(novos_sair)} novos descadastros registrados")
        xlsx = R.gerar_excel(historico, enviados, devolvidos, aberturas, nomes, sair)
        dr.gravar_bin(ctrl[nome_rel]["id"], xlsx, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        log(f"Relatório atualizado: {len(enviados)} disparo(s), {len(devolvidos)} devolvidos, {len(aberturas)} aberturas")
        global RESUMO_ULTIMO
        if enviados:
            d = max(enviados)
            env = enviados[d]
            dev = sum(1 for e in env if (d, e) in devolvidos)
            ab = sum(1 for e in env if (d, e) in aberturas)
            entreg = len(env) - dev
            sem = next((h[1] for h in historico if h[0] == d), "?")
            abertos = f"{ab} ({ab / entreg:.0%} dos entregues)" if d >= "2026-11-18" and entreg else "não medido neste envio"
            RESUMO_ULTIMO = (f"Envio de {dt.date.fromisoformat(d):%d/%m/%Y} (Semana {sem})\n\n"
                             f"Enviados: {len(env)}\nEntregues: {entreg}\nAbertos: {abertos}\n"
                             f"Devolvidos por endereço errado: {dev} (lista para corrigir na aba 'Corrigir e-mails')\n")
        return {e for (_, e) in devolvidos}
    except Exception:
        log("aviso: falha ao gerar relatório:\n" + traceback.format_exc())
        return set()


# ---------------- principal ----------------
def main():
    hoje = dt.date.fromisoformat(os.environ["FORCAR_DATA"]) if os.getenv("FORCAR_DATA") else dt.datetime.now(FUSO).date()
    hj = hoje.isoformat()
    teste = bool(TESTE_PARA)
    log(f"Início {'ENSAIO ' if teste else ''}— dia {hj}")

    dr = Drive()
    raiz = dr.pasta(PASTA_RAIZ)
    itens_raiz = dr.filhos(raiz)
    ctrl = dr.filhos(dr.pasta(PASTA_CTRL, raiz))
    for nec in ("agenda.txt", "envios.txt", "historico.txt", "descadastrados.txt", "log_ultimo_disparo.txt"):
        if nec not in ctrl:
            raise RuntimeError(f"Arquivo de controle ausente: {PASTA_CTRL}/{nec}")
    f_envios, f_hist, f_desc, f_log = (ctrl[n]["id"] for n in ("envios.txt", "historico.txt", "descadastrados.txt", "log_ultimo_disparo.txt"))

    def salvar_log():
        if not teste:
            dr.gravar(f_log, "\n".join(LOG) + "\n")

    try:
        # 1) qual semana
        agenda = {}
        for l in linhas(dr.texto(ctrl["agenda.txt"]["id"])):
            if "=" in l:
                k, v = l.split("=", 1); agenda[k.strip()] = v.strip().lower()
        devolvidos_todos = set() if teste else atualizar_relatorio(dr, ctrl, itens_raiz)
        semana = E.semana_do_mes(hoje)
        if hj in agenda:
            if agenda[hj] == "pular":
                log("Agenda: hoje está marcado para PULAR."); salvar_log(); return
            semana = int(agenda[hj])
        elif hoje.weekday() != 1 and not teste:
            log("Hoje não é dia de envio (só terças ou datas da agenda) — só relatório.")
            if hoje.weekday() == 2 and RESUMO_ULTIMO:
                avisar(f"📊 Relatório semanal — Oferta de Consórcio ({hoje:%d/%m})",
                       "Bom dia! O relatório da semana está pronto no Drive:\n"
                       "Oferta de Consórcio > Controle (não mexer) > Relatório de Envios.xlsx\n\n" + RESUMO_ULTIMO +
                       "\nAbertura é aproximada (alguns programas bloqueiam imagens; o iPhone pode marcar como aberto sozinho).")
            salvar_log(); return
        log(f"Semana {semana}{' (definida na agenda)' if hj in agenda else ' (semana do mês)'}")

        # 2) arte + textos
        if f"Semana {semana}" not in itens_raiz:
            raise RuntimeError(f"Pasta 'Semana {semana}' não existe")
        fs = dr.filhos(itens_raiz[f"Semana {semana}"]["id"])
        artes = sorted([f for f in fs.values() if f["mimeType"] in IMG], key=lambda f: f["modifiedTime"], reverse=True)
        if not artes:
            log(f"Sem arte na pasta Semana {semana} — nada enviado.")
            if not teste:
                avisar(f"⚠️ Oferta de Consórcio {hoje:%d/%m}: sem arte na Semana {semana}",
                       f"A pasta 'Semana {semana}' está sem arte, então nenhum e-mail foi enviado hoje.")
            salvar_log(); return
        arte = artes[0]
        md5 = arte.get("md5Checksum", "")
        hist_txt = dr.texto(f_hist)
        ja_usada = [l for l in linhas(hist_txt) if l.split("\t")[3:4] == [md5] and l.split("\t")[0] != hj]
        if ja_usada and not teste:
            log(f"Arte '{arte['name']}' já foi enviada em {ja_usada[-1].split(chr(9))[0]} — não reenvio.")
            avisar(f"⚠️ Oferta de Consórcio {hoje:%d/%m}: Semana {semana} sem arte nova",
                   f"A arte '{arte['name']}' da pasta Semana {semana} já foi enviada em {ja_usada[-1].split(chr(9))[0]}.\n"
                   "Para não repetir, nada foi enviado hoje. Coloque uma arte nova na pasta (ou ajuste a agenda.txt).")
            salvar_log(); return

        tmp = Path(tempfile.mkdtemp())
        p_arte = tmp / ("arte" + Path(arte["name"]).suffix.lower())
        p_arte.write_bytes(dr.baixar(arte["id"]))
        (tmp / "conteudo.txt").write_text(dr.texto(fs["conteudo.txt"]["id"]) if "conteudo.txt" in fs else "", encoding="utf-8")
        c = E.ler_conteudo(tmp)
        if "assinatura.png" in ctrl:
            (tmp / "assinatura.png").write_bytes(dr.baixar(ctrl["assinatura.png"]["id"]))
            E.ASSINATURA = tmp / "assinatura.png"
        log(f"Arte: {arte['name']} | assunto: {c['ASSUNTO']}")

        # 3) lista
        if teste:
            lista = TESTE_PARA
        else:
            lista, _ = ler_lista(dr, itens_raiz)

        # 4) descadastros
        desc_txt = dr.texto(f_desc)
        sair = {l.strip().lower() for l in linhas(desc_txt)}
        novos = (descadastros_imap() | R.descadastros_web(os.getenv("BLOB_READ_WRITE_TOKEN", ""), log)) - sair
        if novos and not teste:
            desc_txt = desc_txt.rstrip("\n") + "\n" + "\n".join(sorted(novos)) + "\n"
            dr.gravar(f_desc, desc_txt)
        sair |= novos

        # 5) quem já recebeu HOJE (retomada sem duplicar)
        envios_txt = dr.texto(f_envios)
        ja = {l.split("\t")[1] for l in linhas(envios_txt) if l.startswith(hj + "\t")} if not teste else set()
        fila = [m for m in lista if m not in ja and m not in sair and m not in devolvidos_todos]
        n_sair = sum(1 for m in lista if m in sair)
        n_dev = sum(1 for m in lista if m in devolvidos_todos and m not in sair)
        log(f"Lista {len(lista)} | já receberam hoje {len(ja)} | descadastrados {n_sair} | endereço errado (fora) {n_dev} | a enviar {len(fila)}")
        salvar_log()

        # 6) envio
        enviados, falhas, seguidas, parada = 0, [], 0, ""
        novos_envios: list[str] = []
        recusados: dict[str, str] = {}
        smtp = None
        t_ini = time.time()

        def persistir():
            nonlocal envios_txt, novos_envios
            if novos_envios and not teste:
                envios_txt = envios_txt.rstrip("\n") + "\n" + "\n".join(novos_envios) + "\n"
                dr.gravar(f_envios, envios_txt)
                novos_envios = []
            salvar_log()

        for idx, dest in enumerate(fila):
            t0 = time.time()
            try:
                if smtp is None:
                    smtp = smtplib.SMTP_SSL(E.SMTP_HOST, 465, context=ssl.create_default_context(), timeout=60)
                    smtp.login(E.SMTP_USER, E.SMTP_PASS)
                msg, _ = E.montar(dest, p_arte, c, semana, prefixo="[TESTE] " if teste else "", data_disparo=hj)
                smtp.send_message(msg)
                enviados += 1; seguidas = 0
                novos_envios.append(f"{hj}\t{dest}")
                if enviados == 1:
                    try:
                        with imaplib.IMAP4_SSL(E.SMTP_HOST, 993) as im:
                            im.login(E.SMTP_USER, E.SMTP_PASS); im.append('"Sent Items"', "(\\Seen)", None, msg.as_bytes())
                    except Exception:
                        pass
            except smtplib.SMTPAuthenticationError as ex:
                parada = f"login recusado pelo servidor ({ex})"; break
            except smtplib.SMTPRecipientsRefused as ex:
                falhas.append((dest, str(ex)[:160])); seguidas = 0
                recusados[dest] = "recusado na hora pelo servidor: " + re.sub(r"\s+", " ", str(list(ex.recipients.values())[0][1:]))[:150]
            except Exception as ex:
                falhas.append((dest, str(ex)[:160])); seguidas += 1
                try:
                    smtp and smtp.close()
                except Exception:
                    pass
                smtp = None
                if seguidas >= MAX_FALHAS_SEGUIDAS:
                    parada = f"{seguidas} falhas seguidas (último erro: {str(ex)[:200]}) — possível limite/bloqueio do servidor"; break
                time.sleep(60)
            if (idx + 1) % SALVAR_A_CADA == 0:
                log(f"progresso: {enviados} enviados, {len(falhas)} falhas, {len(fila) - idx - 1} restantes")
                persistir()
            if idx + 1 < len(fila):
                time.sleep(max(0.0, INTERVALO - (time.time() - t0)))
        try:
            smtp and smtp.quit()
        except Exception:
            pass

        total = len(ja) + enviados
        status = "PARADO ⚠️" if parada else "CONCLUÍDO ✅"
        dur = (time.time() - t_ini) / 60
        log(f"FIM {status} | enviados agora {enviados} | total hoje {total}/{len(lista)} | falhas {len(falhas)} | {dur:.0f} min {parada}")
        persistir()
        if recusados and not teste and "devolvidos.txt" in ctrl:
            f_dev = ctrl["devolvidos.txt"]["id"]
            dev_txt = dr.texto(f_dev)
            ja_dev = {tuple(l.split("\t")[:2]) for l in linhas(dev_txt)}
            add = [f"{hj}\t{e}\t{m}" for e, m in recusados.items() if (hj, e) not in ja_dev]
            if add:
                dr.gravar(f_dev, dev_txt.rstrip("\n") + "\n" + "\n".join(add) + "\n")
        if not teste:
            hist_linhas = [l for l in hist_txt.splitlines() if not l.startswith(hj + "\t")]
            hist_linhas.append(f"{hj}\t{semana}\t{arte['name']}\t{md5}\t{'parado' if parada else 'concluido'}\t{total}")
            dr.gravar(f_hist, "\n".join(hist_linhas) + "\n")
            ctrl = dr.filhos(dr.pasta(PASTA_CTRL, raiz))
            atualizar_relatorio(dr, ctrl, itens_raiz)
        corpo = (f"Oferta de Consórcio — {hoje:%d/%m/%Y} — Semana {semana}\nAssunto: {c['ASSUNTO']}\n\n"
                 f"Status: {status}\nEnviados: {total} de {len(lista)}\nDescadastrados (não receberam): {n_sair}\n"
                 f"Fora por endereço errado em envio anterior: {n_dev}\n"
                 f"Relatório completo (entregues, abertos, devolvidos): pasta 'Controle (não mexer)' > Relatório de Envios.xlsx\n"
                 f"Falhas: {len(falhas)}\nDuração: {dur:.0f} min\n")
        if parada:
            corpo += f"\nMotivo da parada: {parada}\nAo rodar de novo no mesmo dia, continua de onde parou (ninguém recebe em dobro).\n"
        if falhas:
            corpo += "\nE-mails com falha:\n" + "\n".join(f"- {d}: {m}" for d, m in falhas[:150])
        avisar(f"{'[ENSAIO] ' if teste else ''}Oferta de Consórcio {hoje:%d/%m}: {status} ({total}/{len(lista)})", corpo)
    except Exception:
        tb = traceback.format_exc()
        log("ERRO:\n" + tb)
        salvar_log()
        avisar(f"❌ Oferta de Consórcio {hoje:%d/%m}: erro no robô", tb)
        sys.exit(1)


if __name__ == "__main__":
    main()
