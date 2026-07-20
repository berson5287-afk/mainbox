"""
Hybrid RfqScorer for MaINbox -- keyword pre-filter + local Ollama for the
ambiguous middle. Scores how likely an email is a CUSTOMER asking us to quote
(0..1); the classifier separately combines this with the sender's role.

Three tiers, cheapest first:
  1. high-precision NEGATIVE filter (order confs, ASNs, shipping, invoices,
     auto-replies, 2FA, quarantine, PO acks) -> 0.0, no AI call.
  2. strong RFQ phrasing in the subject/body -> high, no AI call.
  3. everything else -> ask local Ollama a DIRECTIONAL yes/no: is the sender
     requesting a quote FROM us (vs replying with prices, confirming an order,
     etc.). On any AI failure, fall back to the deterministic keyword score.

The directional framing is what keywords can't do: it separates "a contractor
asking us to quote" from "a vendor replying to our P&A" even when subjects look
alike. Limitation: the model only sees subject + body text, so a request whose
items live ONLY in an attachment may read as NO -- attachment parsing is
SmartScan's job, wired separately.

Self-contained: talks to your Ollama (default http://localhost:11434/api/generate,
llama3.2:3b) over urllib, no extra deps and no app import. You can also inject
ai_complete(prompt)->str to route through your own AI host instead.
"""
from __future__ import annotations

import json
import re
import urllib.request
from typing import Callable, Optional

from inbound_router import InboundMail, KeywordRfqScorer

# match MaINbox's Ollama defaults; override per-instance if needed
DEFAULT_OLLAMA_URL = "http://localhost:11434/api/generate"
DEFAULT_OLLAMA_MODEL = "llama3.2:3b"
DEFAULT_KEEP_ALIVE = "30m"

# Tier-1 negatives: phrasings that are NEVER a customer asking us to quote.
# Kept high-precision on purpose -- anything not caught here still goes to the AI.
_NEG_SUBJECT = (
    "automatic reply", "auto-reply", "out of office", "undeliverable",
    "delivery has failed", "read:", "accepted:", "declined:", "tentative:",
    "has shipped", "package has been delivered", "out for delivery",
    "tracking number", "shipment notification", "asn for", "asn -",
    "order confirmation", "order acknowledg", "shipping confirmation",
    "invoice", "remittance", "statement of account",
    "two factor", "2fa", "verification code", "security code", "one-time pass",
    "quarantine", "unsubscribe", "newsletter",
)
_NEG_SENDER = ("donotreply", "do-not-reply", "noreply", "no-reply")

_STRONG = ("request for quote", "request for quotation", "rfq",
           "please quote", "please provide a quote", "requesting a quote",
           "request a quote", "quote request")


def _looks_like_non_rfq(mail: InboundMail) -> bool:
    subj = (mail.subject or "").lower()
    sender = (mail.sender or "").lower()
    if any(p in subj for p in _NEG_SUBJECT):
        return True
    if any(p in sender for p in _NEG_SENDER):
        return True
    return False


def _has_strong_phrase(mail: InboundMail) -> bool:
    text = f"{mail.subject}\n{mail.body}".lower()
    return any(p in text for p in _STRONG)


def _build_prompt(mail: InboundMail) -> str:
    att = "yes" if mail.has_attachments else "no"
    return (
        "You classify inbound email for an electrical-supply distributor.\n"
        "Question: is the SENDER asking US to PROVIDE a price quote -- a customer "
        "request for quote / price-and-availability request directed AT us?\n\n"
        "Answer YES only if the sender wants us to quote prices/availability for them.\n"
        "Answer NO if the sender is any of:\n"
        "- a vendor/supplier REPLYING with prices or availability we requested\n"
        "- confirming, acknowledging, or shipping an order; sending a PO or invoice\n"
        "- an automatic reply, delivery/tracking notice, or internal order alert\n"
        "A subject beginning 'Re:' usually means a reply in an existing thread -- weigh "
        "that toward NO, but judge by the body: a FORWARDED customer request is still YES.\n\n"
        f"From: {mail.sender}\n"
        f"Subject: {mail.subject}\n"
        f"Has attachment: {att}\n"
        f"Body:\n{(mail.body or '')[:1500]}\n\n"
        "Reply with exactly one word: YES or NO."
    )


def _parse_yes_no(text: str, yes: float, no: float) -> Optional[float]:
    # word-boundary match on the FIRST yes/no token, so "not"/"note" don't read as
    # "no" and "yesterday" doesn't read as "yes".
    m = re.search(r"\b(yes|no)\b", (text or "").strip().lower())
    if not m:
        return None
    return yes if m.group(1) == "yes" else no


def _ollama_generate(prompt: str, url: str, model: str, keep_alive: str, timeout: float) -> str:
    payload = {"model": model, "prompt": prompt, "stream": False, "keep_alive": keep_alive,
               "options": {"temperature": 0, "num_predict": 5}}
    data = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(url, data=data, headers={"Content-Type": "application/json"}, method="POST")
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        out = json.loads(resp.read().decode("utf-8", errors="replace"))
    return str(out.get("response", ""))


class HybridRfqScorer:
    def __init__(self, keyword_scorer=None, ai_complete: Optional[Callable[[str], str]] = None,
                 ollama_url: str = DEFAULT_OLLAMA_URL, model: str = DEFAULT_OLLAMA_MODEL,
                 keep_alive: str = DEFAULT_KEEP_ALIVE, timeout: float = 20.0,
                 strong_score: float = 0.85, ai_yes: float = 0.85, ai_no: float = 0.1):
        self.kw = keyword_scorer or KeywordRfqScorer()
        self.ai_complete = ai_complete          # optional: route through your own AI host
        self.ollama_url = ollama_url
        self.model = model
        self.keep_alive = keep_alive
        self.timeout = timeout
        self.strong_score = strong_score
        self.ai_yes = ai_yes
        self.ai_no = ai_no
        self.last_used = None                    # "neg" | "strong" | "ai" | "fallback" (for visibility)

    def score(self, mail: InboundMail) -> float:
        if _looks_like_non_rfq(mail):
            self.last_used = "neg"
            return 0.0
        if _has_strong_phrase(mail):
            self.last_used = "strong"
            return self.strong_score
        verdict = self._ai_verdict(mail)
        if verdict is not None:
            self.last_used = "ai"
            return verdict
        self.last_used = "fallback"
        return self.kw.score(mail)

    def _ai_verdict(self, mail: InboundMail) -> Optional[float]:
        prompt = _build_prompt(mail)
        try:
            if self.ai_complete is not None:
                text = self.ai_complete(prompt)
            else:
                text = _ollama_generate(prompt, self.ollama_url, self.model,
                                        self.keep_alive, self.timeout)
        except Exception:
            return None
        return _parse_yes_no(text, self.ai_yes, self.ai_no)
