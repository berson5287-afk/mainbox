# MaINbox

> RFQ and email workflow automation for electrical supply procurement — built with Python, Tkinter, and Outlook COM automation.

MaINbox is a production desktop application I built and maintain for **American Power Electrical Supply**, a Philadelphia-area electrical distributor serving the NY metro market. It replaced a manual, email-driven quoting process and is used daily to manage vendor RFQs, track quote coverage, and automate Outlook-based workflows.

---

## What it does

- **RFQ Management** — Create, track, and manage Request for Quote workflows across multiple vendors. Items are cross-referenced so coverage from one vendor can be mapped against another.
- **Quote Coverage System** — Bidirectional cross-reference graph that tracks which line items are covered, partially covered, or missing across all open RFQs. Coverage can be extracted automatically from PDF attachments.
- **Outlook COM Integration** — Deep integration with Microsoft Outlook: reads/writes emails, creates draft replies, monitors inbox threads, auto-archives resolved threads, and manages follow-up queues — all without leaving the app.
- **AutoPilot** — Background thread that monitors Outlook activity, applies auto-grouping rules, and handles thread lifecycle management without user intervention.
- **AI Triage Queue** — Vendor emails are routed through a local Ollama model for classification and priority scoring before hitting the main workflow.
- **Cloud Backup/Restore** — Bridge system (`mainbox_cloud_bridge`) for backup and restore of app state across machines.
- **Document Parsing** — Integrated with OmniScan (see companion repo) for extracting line items from vendor PDFs and handwritten RFQ documents.

---

## Architecture

| Layer | Technology |
|---|---|
| UI | Python / Tkinter (single `.pyw` file, ~34K lines) |
| Email integration | `win32com` Outlook COM automation |
| Local database | SQLite via `sqlite3` |
| Document parsing | `pdfplumber`, Tesseract OCR, Docling |
| AI triage | Ollama (local LLM, gemma3 models) |
| Threading | Python `threading` + `queue` (COM-safe worker pattern) |

The entire application is intentionally a single-file architecture. This is a deliberate choice for a one-developer, one-machine production tool: zero deployment complexity, instant rollback by file replacement, and no dependency management at runtime.

---

## Key engineering decisions

**COM thread safety** — Outlook COM requires `pythoncom.CoInitialize()` per worker thread. All COM operations run on dedicated worker threads, never on the Tkinter main thread. A `run_outlook_worker` / `fresh_outlook` pattern is used consistently to prevent cross-thread COM access.

**Internet Message-ID for thread matching** — Outlook's `EntryID` drifts under Cached Exchange Mode. All thread matching uses the RFC-standard `Internet Message-ID` header, which is stable across sync cycles.

**Bidirectional xref graph** — Quote coverage is tracked as a graph structure with transitive closure, so covering item A on RFQ-1 automatically reflects on RFQ-2 if they share a cross-reference chain.

**No ORM, no framework** — Direct SQLite with hand-written queries. At this scale and access pattern, an ORM adds complexity without benefit.

---

## Status

Active production use. Not currently open-sourced (proprietary business logic). This repo represents the architecture and scope of the project for portfolio purposes.

---

## Companion projects

- [`mainbox-brain`](https://github.com/berson5287-afk/mainbox-brain) — Procurement intelligence backend, vendor email mining, voice interface
- [`omniscan`](https://github.com/berson5287-afk/omniscan) — Universal document scanner and parser engine
