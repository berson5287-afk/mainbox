"""
Tests for HybridRfqScorer -- tier routing and fallback, with a recording fake AI
so nothing hits Ollama. The live Ollama path is exercised by you on your machine.
Run: python test_mainbox_rfq_scorer.py
"""
from mainbox_rfq_scorer import HybridRfqScorer, _looks_like_non_rfq, _parse_yes_no
from inbound_router import InboundMail

PASS, FAIL = "  ok  ", "  FAIL"
_results = []


def check(name, cond):
    _results.append(bool(cond))
    print(f"[{PASS if cond else FAIL}] {name}")


def mk(subject="", body="", sender="buyer@acme.test", att=False):
    return InboundMail(message_id="m", sender=sender, subject=subject, body=body, has_attachments=att)


class RecAI:
    """Records calls; returns a fixed string or raises."""
    def __init__(self, ret="", raise_=False):
        self.calls = 0
        self.ret = ret
        self.raise_ = raise_
    def __call__(self, prompt):
        self.calls += 1
        if self.raise_:
            raise RuntimeError("ollama unavailable")
        return self.ret


# ----------------------------- tier 1: negatives -----------------------------
def test_negative_filter_skips_ai():
    ai = RecAI(ret="YES")            # would say YES, but must never be called
    s = HybridRfqScorer(ai_complete=ai)
    for subj in ["Order Confirmation S100", "WDK ASN for PO - P000020221",
                 "Your Package Has Been Delivered!", "Automatic reply: BBJ Welsbach",
                 "SV3 Two Factor Verification", "Invoice 12345"]:
        check(f"negative -> 0.0 ({subj[:24]!r})", s.score(mk(subject=subj)) == 0.0)
    check("negative sender -> 0.0", s.score(mk(subject="hello", sender="donotreply@hubbell.com")) == 0.0)
    check("AI never called for negatives", ai.calls == 0)
    check("last_used recorded as 'neg'", s.last_used == "neg")


# --------------------------- tier 2: strong keyword --------------------------
def test_strong_phrase_skips_ai():
    ai = RecAI(ret="NO")             # would say NO, but must never be called
    s = HybridRfqScorer(ai_complete=ai, strong_score=0.85)
    check("'Request for Quote' -> strong score", s.score(mk(subject="Request for Quote - Job 12")) == 0.85)
    check("'RFQ - ...' -> strong score", s.score(mk(subject="RFQ - Solar Electric / WMC")) == 0.85)
    check("AI never called for strong-keyword mail", ai.calls == 0)
    check("last_used recorded as 'strong'", s.last_used == "strong")


# ------------------------------ tier 3: the AI -------------------------------
def test_ambiguous_goes_to_ai():
    ai_yes = RecAI(ret="YES")
    s = HybridRfqScorer(ai_complete=ai_yes, ai_yes=0.85, ai_no=0.1)
    m = mk(subject="1980 Lexington ave", body="we have a job at this site, can you help us out")
    check("ambiguous + AI YES -> ai_yes", s.score(m) == 0.85)
    check("AI was called once", ai_yes.calls == 1)
    check("last_used recorded as 'ai'", s.last_used == "ai")

    ai_no = RecAI(ret="NO")
    s2 = HybridRfqScorer(ai_complete=ai_no, ai_no=0.1)
    check("ambiguous + AI NO -> ai_no", s2.score(m) == 0.1)


def test_ai_failure_falls_back_to_keyword():
    down = RecAI(raise_=True)
    s = HybridRfqScorer(ai_complete=down)
    # body has a MEDIUM keyword so the keyword fallback returns > 0 (proves it ran)
    m = mk(subject="project materials", body="please send price and availability when you can")
    score = s.score(m)
    check("AI failure -> falls back to keyword score (>0 here)", score > 0.0)
    check("last_used recorded as 'fallback'", s.last_used == "fallback")

    garbage = RecAI(ret="I am not sure about this one")
    s2 = HybridRfqScorer(ai_complete=garbage)
    check("unparseable AI -> fallback too", s2.score(mk(subject="hi", body="random")) == 0.0
          and s2.last_used == "fallback")


# ------------------------------ parse helper ---------------------------------
def test_parse_yes_no():
    check("'YES' -> yes", _parse_yes_no("YES", 0.85, 0.1) == 0.85)
    check("'no.' -> no", _parse_yes_no("no.", 0.85, 0.1) == 0.1)
    check("'I think yes' -> yes", _parse_yes_no("I think yes", 0.85, 0.1) == 0.85)
    check("'unsure' -> None", _parse_yes_no("unsure", 0.85, 0.1) is None)


if __name__ == "__main__":
    test_negative_filter_skips_ai()
    test_strong_phrase_skips_ai()
    test_ambiguous_goes_to_ai()
    test_ai_failure_falls_back_to_keyword()
    test_parse_yes_no()
    print("-" * 56)
    total, ok = len(_results), sum(_results)
    print(f"{ok}/{total} checks passed")
    raise SystemExit(0 if ok == total else 1)
