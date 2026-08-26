# MaINbox

MaINbox runs the quote desk at American Power Electrical Supply, a
Philadelphia-area electrical distributor serving the NY metro market. It sits
on top of desktop Outlook and manages the whole RFQ cycle: a customer request
comes in, its line items get extracted and tracked, RFQs go out to vendors, and
the quotes that come back land on the right job with per-line status. I built
it to replace a purely manual email quoting process and it has been in daily
production use since.

## What it does

Customer requests are picked up the moment they arrive. Typed lists,
sentence-form asks ("can you quote 25 of X and 500ft of Y?"), spreadsheet and
PDF attachments, even screenshots pasted into a body all parse into line items.
Every job keeps a coverage ledger: what was requested, which vendors were
asked, who quoted what at what price and lead time, and what is still open.

Sent mail is tracked, so an RFQ nobody answered surfaces on its own instead of
being remembered. Threads group themselves by job. Follow-ups and dated
commitments (ship dates, quote expirations, callbacks) get queued from the mail
itself. When a request or a reply can't be placed with confidence, the app puts
up a short picker and asks, rather than guessing and filing it wrong.

Triage runs through a local Ollama model when one is reachable. The heuristics
carry the full load on their own when it isn't — the LLM is a refinement, not a
dependency.

## How it's built

Python and Tkinter, roughly 56K lines, deliberately in one file. For a
one-developer production tool that runs a live sales desk, that buys real
things: no deployment step, no dependency drift, and rollback is literally
launching the previous file. Every release is a copy of the last one plus the
change, so the version history doubles as the rollback chain.

A few rules the app lives by, all learned the hard way:

Outlook COM never runs on the UI thread. Every COM operation happens on a
worker with its own `pythoncom.CoInitialize()`, and results marshal back to
Tkinter through the event loop. Breaking this rule freezes the app. Work drifts
back onto the UI thread as features get added, so `mainbox_thread_audit.py`
walks the call graph and reports anything that has; it runs before a release.

Messages are identified by their Internet Message-ID, not Outlook's EntryID —
EntryID drifts under Cached Exchange Mode and will quietly orphan your
tracking.

Matching is gated, not fuzzy-friendly. A vendor line only marks a requested
item quoted when it survives a distinctive-word check, because mapping money
onto the wrong line is worse than asking. The looser "is this the same item?"
test is reserved for questions where a false match is cheap.

State is small JSON stores mirrored into SQLite — no ORM, no framework. A
buffered ops log records every decision point, and that telemetry is what
drives the tuning.

There's also a localhost-only diagnostic bridge (off unless you launch with
`MAINBOX_DIAG=1`): read-only state over HTTP plus a synthetic-email injector
that exercises the real import pipeline, so behavior can be tested without
touching live mail. It has no endpoint that sends anything or deletes real
data, on purpose.

## Status

Active production use, developed continuously. The business data the app
manages stays out of this repo; what's here is the application and its bench of
standalone modules and tests.

## Companion projects

- [`mainbox-brain`](https://github.com/berson5287-afk/mainbox-brain) — procurement intelligence backend, vendor email mining, voice interface
- [`omniscan`](https://github.com/berson5287-afk/omniscan) — universal document scanner and parser engine
