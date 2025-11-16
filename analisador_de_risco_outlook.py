"""
analisador_de_risco_outlook.py
Requer: pywin32  (pip install pywin32)
Roda em Windows com Outlook Desktop configurado (MAPI).
"""

import re
import csv
import os
import html
import datetime
from collections import Counter
from win32com import client as win32
from urllib.parse import urlparse
import difflib

# ---------- Configurações ----------
FOLDER_NAME = "Inbox"   # pasta a analisar (ex: "Inbox", "Caixa de Entrada")
MARK_SUBJECT = True     # adiciona prefixo no assunto com classificação
MARK_PREFIX = "[Risco:{}]"  # exemplo: [Risco:ALTO]
CSV_OUTPUT = "relatorio_risco_emails.csv"
MAX_EMAILS = 500        # limite para processar por execução (evitar loop muito grande)

# Palavras/expressões indicadoras de phishing/urgência/suspensão
SUSPICIOUS_KEYWORDS = [
    "senha", "reset", "alterar", "confirmar", "verifique", "verificação",
    "fatura", "pagamento", "atualizar", "urgente", "imediato", "clique aqui",
    "account", "verify", "bank", "confirme", "recuperar", "security",
    "suspend", "limited time", "doc", "invoice"
]

# tipos de anexos perigosos (executáveis / macros)
SUSPICIOUS_EXT = {".exe", ".scr", ".pif", ".bat", ".cmd", ".js", ".vbs", ".wsf", ".docm", ".xlsm", ".hta", ".msi"}

# pesos heurísticos (ajustáveis)
WEIGHTS = {
    "keywords": 30,
    "urls": 25,
    "attachment": 25,
    "sender_mismatch": 30,
    "many_recipients": 10,
    "html_form": 20,
    "domain_age_proxy": 0  # deixei 0 porque não consulto a web aqui
}

# thresholds (pontuação 0..100)
TH_LOW = 25
TH_MED = 55

# ---------- Helpers ----------
url_regex = re.compile(r'(https?://[^\s<>"]+|www\.[^\s<>"]+)', re.IGNORECASE)

def extract_urls_from_text(text):
    if not text:
        return []
    return url_regex.findall(text)

def get_domain(url):
    try:
        if not re.match(r'https?://', url):
            url = 'http://' + url
        parsed = urlparse(url)
        domain = parsed.hostname or ""
        return domain.lower().lstrip('www.')
    except Exception:
        return ""

def text_contains_keyword(text, keywords):
    if not text:
        return False
    low = text.lower()
    for k in keywords:
        if k.lower() in low:
            return True
    return False

def attachment_risk_score(attachments):
    score = 0
    for att in attachments:
        name = (att.FileName or "").lower()
        _, ext = os.path.splitext(name)
        if ext in SUSPICIOUS_EXT:
            score += 100  # considerável
    return min(score, 100)

def domain_similarity(a, b):
    # ratio de similaridade simples entre dois domínios
    if not a or not b:
        return 0.0
    return difflib.SequenceMatcher(None, a, b).ratio()

def suspicious_link_text_check(html_body):
    """
    Detecta links cujo texto visível é um domínio mas o href aponta para outro domínio.
    Retorna 100 se encontrar mismatch suspeito (alto risco), 0 caso contrário.
    Simples: procura padrões <a ...>texto</a> com texto contendo domínio.
    """
    mismatches = 0
    try:
        # procura tags <a ...>...</a>
        for m in re.finditer(r'<a[^>]+href=["\']([^"\']+)["\'][^>]*>(.*?)</a>', html_body or "", flags=re.I|re.S):
            href = m.group(1).strip()
            text = re.sub(r'<.*?>', '', m.group(2) or "").strip()  # limpa tags internas
            domain_href = get_domain(href)
            domain_text = get_domain(text)
            # se o texto parece um domínio e os domínios diferem muito -> suspeito
            if domain_text and domain_href and domain_similarity(domain_href, domain_text) < 0.7:
                mismatches += 1
    except Exception:
        return 0
    return 100 if mismatches > 0 else 0

def html_has_form(html_body):
    if not html_body:
        return False
    return bool(re.search(r'<form\b', html_body, flags=re.I))

# ---------- Scoring ----------
def score_email(message):
    """
    Entrada: objeto MailItem do Outlook (win32com)
    Retorna: dict com 'score' (0..100), 'classification' (Baixo/Médio/Alto), 'details'
    """
    score = 0
    details = {}

    # Remetente
    sender = getattr(message, "SenderEmailAddress", "") or ""
    # Display name (remetente visível)
    sender_name = getattr(message, "SenderName", "") or ""

    # Assunto e corpo
    subject = getattr(message, "Subject", "") or ""
    # Tenta pegar corpo em HTML, senão TextBody
    try:
        body_html = getattr(message, "HTMLBody", None)
    except Exception:
        body_html = None
    try:
        body_text = getattr(message, "Body", "") or ""
    except Exception:
        body_text = ""

    # Destinatários
    try:
        to_count = len([r for r in (message.Recipients or [])])
    except Exception:
        to_count = 1

    # Anexos
    try:
        attachments = list(message.Attachments or [])
    except Exception:
        attachments = []

    # 1) Palavras-chave suspeitas no assunto ou corpo
    kw_flag = text_contains_keyword(subject + "\n" + (body_text or ""), SUSPICIOUS_KEYWORDS)
    kw_score = WEIGHTS["keywords"] if kw_flag else 0
    score += kw_score
    details['keywords'] = kw_score

    # 2) URLs suspeitas / mismatches
    urls = set(extract_urls_from_text((body_html or "") + "\n" + body_text))
    url_score = 0
    if urls:
        # penaliza se muitos links ou links com IPs, ou mismatch entre texto e href
        for u in urls:
            dom = get_domain(u)
            if re.match(r'^\d{1,3}(\.\d{1,3}){3}$', dom or ""):
                url_score += 40
            # domínio estranho (placeholder: se contém '-' em início ou usa IP) -> pontua
            if dom.count('-') > 2:
                url_score += 10
        # detectar mismatches em html
        if body_html:
            url_score = max(url_score, suspicious_link_text_check(body_html))
        # manter dentro de peso
        url_score = min(url_score, WEIGHTS["urls"])
    score += url_score
    details['urls'] = url_score

    # 3) Anexos com extensões perigosas
    att_score = 0
    if attachments:
        att_score = attachment_risk_score(attachments)
        # escala para peso máximo definido
        att_score = min(att_score, WEIGHTS["attachment"])
    score += att_score
    details['attachments'] = att_score

    # 4) Sender display name mismatch (ex.: nome "Banco X" mas domínio diferente)
    sender_mismatch_score = 0
    try:
        # tenta inferir domínio do email
        sender_domain = ""
        if sender:
            sender_domain = sender.split('@')[-1].lower() if '@' in sender else sender.lower()
        # verifica se display name contém domínio/empresa e se diverge
        probable_org_in_name = None
        m = re.search(r'@?([A-Za-z0-9\-\.]{3,})', sender_name or "")
        if m:
            probable_org_in_name = m.group(1).lower()
        # se existir e divergir fortemente -> suspeita
        if probable_org_in_name and sender_domain and domain_similarity(probable_org_in_name, sender_domain) < 0.5:
            sender_mismatch_score = WEIGHTS["sender_mismatch"]
    except Exception:
        sender_mismatch_score = 0
    score += sender_mismatch_score
    details['sender_mismatch'] = sender_mismatch_score

    # 5) Muitos destinatários (mass mail) -> leve risco
    many_recip_score = WEIGHTS["many_recipients"] if to_count > 10 else 0
    score += many_recip_score
    details['many_recipients'] = many_recip_score

    # 6) HTML contendo form (phishing tende a perguntar informações)
    html_form_score = WEIGHTS["html_form"] if html_has_form(body_html) else 0
    score += html_form_score
    details['html_form'] = html_form_score

    # escala final para máximo 100
    if score > 100:
        score = 100

    # classificação
    if score < TH_LOW:
        cls = "Baixo"
    elif score < TH_MED:
        cls = "Médio"
    else:
        cls = "Alto"

    return {"score": int(score), "classification": cls, "details": details,
            "subject": subject, "sender": sender, "sender_name": sender_name,
            "to_count": to_count, "has_attachments": len(attachments) > 0, "urls": list(urls)}

# ---------- Integração com Outlook (MAPI) ----------
def analyze_inbox(limit=MAX_EMAILS, mark_subject=MARK_SUBJECT):
    outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
    # 6 é a Inbox em muitas instalações; também pode buscar por nome
    inbox = outlook.GetDefaultFolder(6)
    # se quiser uma subpasta: inbox.Folders['NomeDaPasta']
    messages = inbox.Items
    # filtrar por não lidos primeiro
    try:
        filtered = messages.Restrict("[UnRead] = True")
    except Exception:
        filtered = messages
    count = min(len(filtered), limit) if hasattr(filtered, "__len__") else limit

    results = []
    processed = 0

    # iterar (o objeto pode ser indexado 1..n)
    # Safety: iterar via while para evitar problemas com coleção dinâmica
    i = 1
    while processed < limit:
        try:
            mail = filtered.Item(i)
        except Exception:
            break
        try:
            if mail.Class == 43:  # olMailItem
                res = score_email(mail)
                res['entry_id'] = mail.EntryID
                res['received_time'] = str(getattr(mail, "ReceivedTime", ""))
                results.append(res)

                # marcar assunto (opcional)
                if mark_subject:
                    try:
                        new_subject = MARK_PREFIX.format(res['classification']) + " " + res['subject']
                        mail.Subject = new_subject
                        mail.Save()
                    except Exception:
                        pass

                processed += 1
        except Exception as e:
            # pular mensagens problemáticas
            print("Erro processando item:", e)
        i += 1

    # salvar CSV
    with open(CSV_OUTPUT, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["received_time", "sender_name", "sender_email", "subject", "score", "classification", "to_count", "has_attachments", "urls"])
        for r in results:
            writer.writerow([r.get('received_time'), r.get('sender_name'), r.get('sender'), r.get('subject'), r.get('score'), r.get('classification'), r.get('to_count'), r.get('has_attachments'), ";".join(r.get('urls', []))])

    return results

# ---------- Execução ----------
if __name__ == "__main__":
    print("Iniciando análise de e-mails no Outlook...")
    start = datetime.datetime.now()
    results = analyze_inbox()
    end = datetime.datetime.now()
    print(f"Processados: {len(results)} mensagens em {(end-start).total_seconds():.1f}s")
    # resumo
    counter = Counter([r['classification'] for r in results])
    print("Resumo:", dict(counter))
    print(f"Relatório salvo em: {CSV_OUTPUT}")
