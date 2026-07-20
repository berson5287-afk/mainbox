# ============================================================
# MaINbox Cloud Bridge - Phase 2.5 Backup + Mirror + Restore (Lock-safe restore)
# Copyright (c) 2026 Stephen Berson
# Product: MaINbox
# Version: v3 / MaINbox v3 compatible
# ============================================================
r"""
Phase 2.5 Cloud Bridge for MaINbox + Supabase.

What this bridge does:
- Leaves MaINbox local-first during normal operation. Backups never change MaINbox data.
- Creates consistent timestamped backups of mainbox_data.db using SQLite backup.
- Packages the snapshot into a .zip file with safe metadata and optional JSON exports.
- Uploads the zip to Supabase Storage.
- Uploads a latest/read-only JSON mirror for dashboards or future cloud viewers.
- Watches MaINbox's request file for manual backup requests.
- On an explicit, opt-in restore (--restore-latest), downloads the latest cloud backup
  and replaces the local mainbox_data.db as a whole-file restore while MaINbox is closed.
  v3.8.31: after the replace, it re-stamps app_state.updated_at = now and writes a
  db_meta restore marker so MaINbox treats the restored rows as the newest copy (see
  stamp_restored_db_freshness for the full rationale).

What this bridge does NOT do:
- It does not merge records between computers (restore is whole-file, not a two-way merge).
- It does not touch MaINbox data except during an explicit restore. Backups and the
  read-only cloud mirror never write back; the local DB stays the source of truth unless
  you deliberately run a restore.

Setup:
1. Run MaINbox v3 once so it creates:
   %LOCALAPPDATA%\MaINbox\mainbox_cloud_bridge_config.json
2. Edit that config file and fill in:
   supabase_url
   supabase_service_role_key
   supabase_bucket
   enabled: true
3. Run this bridge:
   python mainbox_cloud_bridge_v4_3.py

Optional:
   python mainbox_cloud_bridge_v4_3.py --once
   python mainbox_cloud_bridge_v4_3.py --config "C:\\path\\to\\mainbox_cloud_bridge_config.json"
"""

from __future__ import annotations

import argparse
import json
import os
import shutil
import socket
import sqlite3
import subprocess
import sys
import time
import traceback
import uuid
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, Optional
from urllib import error, request

BRIDGE_VERSION = "mainbox_cloud_bridge_v4_4"
DEFAULT_BUCKET = "mainbox-backups"


def now_iso() -> str:
    return datetime.now().isoformat(timespec="seconds")


def utc_stamp() -> str:
    return datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S_UTC")


def default_app_data_dir() -> Path:
    base = os.environ.get("LOCALAPPDATA")
    if not base:
        base = os.path.join(os.path.expanduser("~"), "AppData", "Local")
    return Path(base) / "MaINbox"


def default_config_path() -> Path:
    return default_app_data_dir() / "mainbox_cloud_bridge_config.json"


def load_json_file(path: Path, fallback: Any) -> Any:
    try:
        if path.exists():
            return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        pass
    return fallback


def write_json_file(path: Path, data: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, indent=2, default=str), encoding="utf-8")


def sanitize_config_for_status(config: Dict[str, Any]) -> Dict[str, Any]:
    safe = dict(config or {})
    if safe.get("supabase_service_role_key"):
        safe["supabase_service_role_key"] = "***hidden***"
    return safe


def load_config(path: Path) -> Dict[str, Any]:
    cfg = load_json_file(path, {})
    if not isinstance(cfg, dict):
        cfg = {}

    app_dir = Path(cfg.get("app_data_dir") or default_app_data_dir())
    defaults = {
        "enabled": False,
        "phase": "phase_2_5_backup_mirror_restore",
        "app_name": "MaINbox",
        "app_version": "v3",
        "device_id": str(uuid.uuid4()),
        "app_data_dir": str(app_dir),
        "main_db_file": str(app_dir / "mainbox_data.db"),
        "snapshot_dir": str(app_dir / "cloud_bridge_snapshots"),
        "restore_dir": str(app_dir / "cloud_bridge_restores"),
        "request_file": str(app_dir / "mainbox_cloud_backup_request.json"),
        "status_file": str(app_dir / "mainbox_cloud_bridge_status.json"),
        "supabase_url": "",
        "supabase_service_role_key": "",
        "supabase_bucket": DEFAULT_BUCKET,
        "cloud_prefix": "mainbox",
        "auto_backup_minutes": 30,
        "retention_local_snapshots": 12,
        "include_json_exports": True,
        "create_bucket_if_missing": True,
        "upload_latest_json_mirror": True,
        "latest_mirror_prefix": "latest",
        "allow_cloud_restore": True,
        # v3.8.32: leave blank to restore THIS computer's own latest backup. To restore a
        # DIFFERENT computer's backup (e.g. moving to a new machine), set this to that
        # computer's device_id -- find it in that machine's mainbox_cloud_bridge_config.json
        # or in any snapshot_metadata.json it uploaded. Backups still upload under the local
        # device_id; only the restore-side manifest lookup uses this.
        "restore_source_device_id": "",
    }
    changed = False
    for k, v in defaults.items():
        if k not in cfg:
            cfg[k] = v
            changed = True
    if changed or not path.exists():
        write_json_file(path, cfg)
    return cfg


def write_status(config: Dict[str, Any], status: str, message: str = "", **extra: Any) -> None:
    status_file = Path(config.get("status_file") or default_app_data_dir() / "mainbox_cloud_bridge_status.json")
    payload = {
        "status": status,
        "message": message,
        "last_run_at": now_iso(),
        "bridge_version": BRIDGE_VERSION,
        "config": sanitize_config_for_status(config),
    }
    payload.update(extra)
    write_json_file(status_file, payload)


def sqlite_backup(src_db: Path, dst_db: Path) -> None:
    dst_db.parent.mkdir(parents=True, exist_ok=True)
    if not src_db.exists():
        raise FileNotFoundError(f"Main MaINbox database not found: {src_db}")

    # SQLite backup API gives a consistent snapshot even if MaINbox is open.
    src = sqlite3.connect(f"file:{src_db}?mode=ro", uri=True, timeout=30)
    try:
        dst = sqlite3.connect(str(dst_db), timeout=30)
        try:
            src.backup(dst)
            dst.commit()
        finally:
            dst.close()
    finally:
        src.close()


def export_app_state_json(snapshot_dir: Path, snapshot_db: Path) -> Optional[Path]:
    export_dir = snapshot_dir / "json_exports"
    export_dir.mkdir(parents=True, exist_ok=True)
    try:
        conn = sqlite3.connect(str(snapshot_db))
        try:
            rows = conn.execute("SELECT state_key, json_value FROM app_state ORDER BY state_key").fetchall()
        finally:
            conn.close()
        for state_key, json_value in rows:
            safe_name = "".join(ch if ch.isalnum() or ch in "_-" else "_" for ch in str(state_key))
            out = export_dir / f"{safe_name}.json"
            try:
                parsed = json.loads(json_value)
                out.write_text(json.dumps(parsed, indent=2, default=str), encoding="utf-8")
            except Exception:
                out.write_text(str(json_value), encoding="utf-8")
        return export_dir
    except Exception:
        traceback.print_exc()
        return None


def make_snapshot(config: Dict[str, Any], reason: str = "scheduled") -> Path:
    app_dir = Path(config.get("app_data_dir") or default_app_data_dir())
    src_db = Path(config.get("main_db_file") or app_dir / "mainbox_data.db")
    snapshot_root = Path(config.get("snapshot_dir") or app_dir / "cloud_bridge_snapshots")
    device_id = str(config.get("device_id") or "mainbox-device")
    stamp = utc_stamp()

    work_dir = snapshot_root / f"snapshot_{stamp}"
    work_dir.mkdir(parents=True, exist_ok=True)
    snapshot_db = work_dir / "mainbox_data.db"
    sqlite_backup(src_db, snapshot_db)

    metadata = {
        "created_at": now_iso(),
        "created_at_utc_stamp": stamp,
        "reason": reason,
        "bridge_version": BRIDGE_VERSION,
        "phase": "phase_2_5_backup_mirror_restore",
        "device_id": device_id,
        "source_db": str(src_db),
        "app_version": config.get("app_version", "v3"),
        "note": "Backup-only snapshot. Safe to restore manually; bridge does not write this back automatically.",
    }
    write_json_file(work_dir / "snapshot_metadata.json", metadata)

    if bool(config.get("include_json_exports", True)):
        export_app_state_json(work_dir, snapshot_db)

    zip_name = f"mainbox_backup_{device_id}_{stamp}.zip"
    zip_path = snapshot_root / zip_name
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for file_path in work_dir.rglob("*"):
            if file_path.is_file():
                zf.write(file_path, arcname=str(file_path.relative_to(work_dir)))

    shutil.rmtree(work_dir, ignore_errors=True)
    return zip_path


def cleanup_old_snapshots(config: Dict[str, Any]) -> None:
    try:
        snapshot_root = Path(config.get("snapshot_dir") or default_app_data_dir() / "cloud_bridge_snapshots")
        keep = int(config.get("retention_local_snapshots", 12) or 12)
        zips = sorted(snapshot_root.glob("mainbox_backup_*.zip"), key=lambda p: p.stat().st_mtime, reverse=True)
        for old in zips[keep:]:
            try:
                old.unlink()
            except Exception:
                pass
    except Exception:
        pass


def supabase_request(config: Dict[str, Any], method: str, path: str, body: Optional[bytes] = None, content_type: str = "application/json") -> bytes:
    supabase_url = str(config.get("supabase_url") or "").rstrip("/")
    key = str(config.get("supabase_service_role_key") or "")
    if not supabase_url or not key:
        raise RuntimeError("Supabase URL/key not configured. Edit mainbox_cloud_bridge_config.json first.")

    url = f"{supabase_url}{path}"
    headers = {
        "apikey": key,
        "Authorization": f"Bearer {key}",
        "Content-Type": content_type,
    }
    req = request.Request(url, data=body, headers=headers, method=method)
    with request.urlopen(req, timeout=120) as resp:
        return resp.read()


def ensure_supabase_bucket(config: Dict[str, Any]) -> None:
    if not bool(config.get("create_bucket_if_missing", True)):
        return
    bucket = str(config.get("supabase_bucket") or DEFAULT_BUCKET)
    payload = json.dumps({"id": bucket, "name": bucket, "public": False}).encode("utf-8")
    try:
        supabase_request(config, "POST", "/storage/v1/bucket", payload, "application/json")
    except error.HTTPError as exc:
        # 400/409 commonly mean already exists or not allowed. Existing bucket is OK.
        if exc.code not in (400, 409):
            raise
    except Exception:
        # If bucket creation is blocked by policy, upload may still work if bucket already exists.
        pass


def supabase_upload_bytes(config: Dict[str, Any], object_path: str, data: bytes, content_type: str) -> str:
    """Upload one object to Supabase Storage and return bucket/object path."""
    bucket = str(config.get("supabase_bucket") or DEFAULT_BUCKET).strip("/")
    object_path = str(object_path or "").strip("/")
    if not object_path:
        raise ValueError("Supabase object_path cannot be blank.")
    storage_path = f"/storage/v1/object/{bucket}/{object_path}"

    ensure_supabase_bucket(config)
    supabase_url = str(config.get("supabase_url") or "").rstrip("/")
    key = str(config.get("supabase_service_role_key") or "")
    if not supabase_url or not key:
        raise RuntimeError("Supabase URL/key not configured. Edit mainbox_cloud_bridge_config.json first.")

    headers = {
        "apikey": key,
        "Authorization": f"Bearer {key}",
        "Content-Type": content_type,
        "x-upsert": "true",
    }
    url = f"{supabase_url}{storage_path}"
    req = request.Request(url, data=data, headers=headers, method="POST")
    try:
        with request.urlopen(req, timeout=180) as resp:
            resp.read()
    except error.HTTPError as exc:
        # Some Supabase deployments accept PUT for upsert better than POST.
        if exc.code in (400, 405, 409):
            req = request.Request(url, data=data, headers=headers, method="PUT")
            with request.urlopen(req, timeout=180) as resp:
                resp.read()
        else:
            raise
    return f"{bucket}/{object_path}"



def supabase_download_bytes(config: Dict[str, Any], object_or_cloud_path: str) -> bytes:
    """Download one object from Supabase Storage.

    Accepts either an object path like mainbox/device/file.zip or a full
    bucket/object path like mainbox-backups/mainbox/device/file.zip.
    """
    bucket = str(config.get("supabase_bucket") or DEFAULT_BUCKET).strip("/")
    object_path = str(object_or_cloud_path or "").strip("/")
    if object_path.startswith(bucket + "/"):
        object_path = object_path[len(bucket) + 1:]
    if not object_path:
        raise ValueError("Supabase object path cannot be blank.")
    return supabase_request(config, "GET", f"/storage/v1/object/{bucket}/{object_path}", None, "application/octet-stream")


def latest_manifest_object_path(config: Dict[str, Any]) -> str:
    prefix = str(config.get("cloud_prefix") or "mainbox").strip("/")
    # v3.8.32: this path is used by the RESTORE/download read side only (uploads build
    # their own path under the local device_id). If restore_source_device_id is set, look
    # up another computer's latest manifest so "Restore Latest Cloud Backup" works when
    # moving to a new machine; otherwise fall back to this computer's own device_id.
    device_id = str(
        config.get("restore_source_device_id")
        or config.get("device_id")
        or "mainbox-device"
    ).strip("/")
    latest_prefix = str(config.get("latest_mirror_prefix") or "latest").strip("/")
    return f"{prefix}/{device_id}/{latest_prefix}/manifest.json"


def read_latest_cloud_manifest(config: Dict[str, Any]) -> Dict[str, Any]:
    # v4.4: when restore_source_device_id is blank, "Restore Latest" now means the
    # newest backup across ALL of this account's devices -- not this computer's own
    # (usually older) backup. Backing up at work and restoring at home previously
    # restored the HOME PC's own last backup, silently undoing the whole point of
    # the round trip. Pinning restore_source_device_id still restores exactly that
    # device; any listing/parsing problem falls back to the old own-device path.
    pinned = str(config.get("restore_source_device_id") or "").strip()
    if not pinned:
        try:
            best = find_newest_manifest_across_devices(config)
            if best:
                return best
        except Exception:
            traceback.print_exc()
    data = supabase_download_bytes(config, latest_manifest_object_path(config))
    manifest = json.loads(data.decode("utf-8"))
    if not isinstance(manifest, dict):
        raise RuntimeError("Latest cloud manifest was not a JSON object.")
    return manifest


def find_newest_manifest_across_devices(config: Dict[str, Any]) -> Dict[str, Any]:
    """v4.4: list every device folder under the cloud prefix and return the manifest
    with the newest uploaded_at. Returns {} when nothing readable is found."""
    bucket = str(config.get("supabase_bucket") or DEFAULT_BUCKET).strip("/")
    prefix = str(config.get("cloud_prefix") or "mainbox").strip("/")
    latest_prefix = str(config.get("latest_mirror_prefix") or "latest").strip("/")
    body = json.dumps({
        "prefix": f"{prefix}/",
        "limit": 200,
        "offset": 0,
        "sortBy": {"column": "name", "order": "asc"},
    }).encode("utf-8")
    raw = supabase_request(config, "POST", f"/storage/v1/object/list/{bucket}", body, "application/json")
    entries = json.loads(raw.decode("utf-8"))
    device_ids = []
    if isinstance(entries, list):
        for it in entries:
            name = str((it or {}).get("name") or "").strip().strip("/")
            # Supabase returns FOLDERS as entries with id=None; files carry an id.
            if name and (it or {}).get("id") is None and name not in device_ids:
                device_ids.append(name)
    best: Dict[str, Any] = {}
    best_ts = ""
    for dev in device_ids:
        try:
            data = supabase_download_bytes(config, f"{prefix}/{dev}/{latest_prefix}/manifest.json")
            manifest = json.loads(data.decode("utf-8"))
            if not isinstance(manifest, dict) or not str(manifest.get("backup_cloud_path") or "").strip():
                continue
            ts = str(manifest.get("uploaded_at") or "")
            if ts > best_ts:
                best, best_ts = manifest, ts
        except Exception:
            continue
    return best


def download_latest_backup(config: Dict[str, Any]) -> Dict[str, Any]:
    """Download the latest backup zip referenced by the cloud manifest."""
    if not bool(config.get("enabled", False)):
        raise RuntimeError("Cloud bridge is not enabled. Set enabled=true in mainbox_cloud_bridge_config.json.")
    manifest = read_latest_cloud_manifest(config)
    backup_cloud_path = str(manifest.get("backup_cloud_path") or "").strip()
    if not backup_cloud_path:
        raise RuntimeError("Latest cloud manifest does not include backup_cloud_path.")

    restore_dir = Path(config.get("restore_dir") or default_app_data_dir() / "cloud_bridge_restores")
    restore_dir.mkdir(parents=True, exist_ok=True)
    filename = os.path.basename(backup_cloud_path.rstrip("/")) or f"mainbox_cloud_backup_{utc_stamp()}.zip"
    out_path = restore_dir / filename
    out_path.write_bytes(supabase_download_bytes(config, backup_cloud_path))

    return {
        "manifest": manifest,
        "backup_cloud_path": backup_cloud_path,
        "downloaded_zip": str(out_path),
        "downloaded_size_bytes": out_path.stat().st_size,
    }


def validate_snapshot_zip(zip_path: Path, extract_dir: Path) -> Path:
    if not zip_path.exists():
        raise FileNotFoundError(f"Downloaded backup zip not found: {zip_path}")
    if extract_dir.exists():
        shutil.rmtree(extract_dir, ignore_errors=True)
    extract_dir.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(zip_path, "r") as zf:
        names = zf.namelist()
        if "mainbox_data.db" not in names:
            raise RuntimeError("Backup zip does not contain mainbox_data.db.")
        zf.extractall(extract_dir)
    candidate_db = extract_dir / "mainbox_data.db"
    conn = sqlite3.connect(str(candidate_db), timeout=30)
    try:
        integrity = conn.execute("PRAGMA integrity_check").fetchone()
        if not integrity or str(integrity[0]).lower() != "ok":
            raise RuntimeError(f"Downloaded backup database failed integrity check: {integrity}")
    finally:
        conn.close()
    return candidate_db


def stamp_restored_db_freshness(main_db: Path, source_zip: str = "") -> Dict[str, Any]:
    """v3.8.31: after a whole-file restore, mark every restored app_state row as
    freshly written (updated_at = now) and record a restore marker in db_meta.

    Why this matters: MaINbox treats the SQLite DB as the authoritative runtime source
    and mirrors data between JSON files and the DB. A restored database carries the OLD
    updated_at timestamps from when the backup was taken. If MaINbox compares a synced
    JSON file's modification time against that per-row updated_at, a local JSON file that
    was edited AFTER the backup was taken would look "newer" than the restored row and
    could silently overwrite the restore. Re-stamping updated_at = now makes the restored
    rows unambiguously the newest copy, so the restore wins. This is what lets MaINbox
    move from a coarse whole-database-file mtime check to a precise per-row updated_at
    check without risking that a stale local JSON undoes a cloud restore.

    The timestamp format intentionally matches what MaINbox writes (datetime.now()
    .isoformat(), local naive), so the comparison on the MaINbox side is apples-to-apples.

    The database has already been replaced before this runs, so any failure here is
    non-fatal: the restore itself still succeeded. Very old backups may predate the
    app_state/db_meta tables; MaINbox recreates them on next launch, so a failure to
    stamp simply means MaINbox falls back to rebuilding state from the restored data.
    """
    info: Dict[str, Any] = {"app_state_rows_restamped": 0, "restore_marker_written": False}
    now_iso_local = datetime.now().isoformat()
    try:
        with sqlite3.connect(str(main_db)) as conn:
            try:
                cur = conn.execute("UPDATE app_state SET updated_at = ?", (now_iso_local,))
                rc = cur.rowcount
                info["app_state_rows_restamped"] = int(rc) if isinstance(rc, int) and rc >= 0 else 0
            except Exception:
                # app_state table missing (very old backup) -> nothing to stamp.
                traceback.print_exc()
            try:
                conn.execute(
                    "INSERT OR REPLACE INTO db_meta(meta_key, meta_value, updated_at) VALUES (?, ?, ?)",
                    ("last_cloud_restore_at", now_iso_local, now_iso_local),
                )
                if source_zip:
                    conn.execute(
                        "INSERT OR REPLACE INTO db_meta(meta_key, meta_value, updated_at) VALUES (?, ?, ?)",
                        ("last_cloud_restore_from", str(source_zip), now_iso_local),
                    )
                info["restore_marker_written"] = True
            except Exception:
                # db_meta table missing in a very old backup; the app recreates it on launch.
                traceback.print_exc()
            conn.commit()
    except Exception:
        traceback.print_exc()
    return info


# app_state keys that are truly machine-local and safe to drop on a cross-machine
# restore. Keys listed in _MACHINE_LOCAL_PREFIX_KEYS are matched as prefixes
# (MaINbox stores them as '<key>::<mailbox cache key>').
#
# v4.4: "email_statuses" is NO LONGER deleted. Deleting it wiped every workflow
# status/hidden/completed flag on every restore -- the exact data a cross-PC
# mirror exists to carry. MaINbox v4.0.0 re-links foreign Outlook EntryIDs to the
# local profile on the first refresh (_relink_restored_entry_ids), and no current
# startup path COM-resolves status records in bulk, so the old freeze this
# deletion papered over no longer applies. "scanner_commitment_processed_message_ids"
# is also kept: it is keyed on RFC Message-IDs (machine-independent) and deleting
# it resurrected commitment reminders the user had already cancelled. Only the
# sent-item dedup set remains dropped -- it is keyed on sent-item EntryIDs that
# genuinely do not exist on another machine and rebuilds harmlessly.
# v4.3's prefix matching also never fired (no listed key contained '::'), so the
# prefix list is now explicit.
_MACHINE_LOCAL_APP_STATE_KEYS: list[str] = []
_MACHINE_LOCAL_PREFIX_KEYS: list[str] = [
    "scanner_sent_waiting_processed_sent_ids",
]


def sanitize_machine_local_app_state(main_db: Path) -> Dict[str, Any]:
    """After a whole-file restore, delete app_state rows that contain Outlook
    EntryIDs or other COM-bound, machine-local data.

    These rows are safe to drop because MaINbox rebuilds them from the local
    Outlook profile on first launch.  Carrying them over from a different machine
    causes MaINbox to deadlock on startup while trying to resolve thousands of
    EntryIDs that do not exist in the local COM session.

    Keys in _MACHINE_LOCAL_APP_STATE_KEYS that contain '::' are matched as a
    prefix so scanner keys like 'scanner_commitment_processed_message_ids::SB'
    are caught regardless of the mailbox suffix.

    This runs after stamp_restored_db_freshness so the rest of the restored data
    is already correctly stamped.  Any failure here is non-fatal.
    """
    info: Dict[str, Any] = {"machine_local_rows_deleted": 0, "machine_local_keys_deleted": []}
    try:
        with sqlite3.connect(str(main_db)) as conn:
            # v4.4: explicit exact + prefix lists (the old '::' detection never matched).
            exact: list[str] = list(_MACHINE_LOCAL_APP_STATE_KEYS)
            prefixes: list[str] = [k + "::" for k in _MACHINE_LOCAL_PREFIX_KEYS]

            deleted_keys: list[str] = []

            # Exact matches
            for key in exact:
                cur = conn.execute("DELETE FROM app_state WHERE state_key = ?", (key,))
                if cur.rowcount and cur.rowcount > 0:
                    deleted_keys.append(key)
                    info["machine_local_rows_deleted"] = info["machine_local_rows_deleted"] + int(cur.rowcount)

            # Prefix matches (e.g. scanner_commitment_processed_message_ids::SB)
            for prefix in prefixes:
                cur = conn.execute("DELETE FROM app_state WHERE state_key LIKE ?", (prefix + "%",))
                if cur.rowcount and cur.rowcount > 0:
                    deleted_keys.append(prefix + "*")
                    info["machine_local_rows_deleted"] = info["machine_local_rows_deleted"] + int(cur.rowcount)

            conn.commit()
            info["machine_local_keys_deleted"] = deleted_keys
    except Exception:
        traceback.print_exc()
    return info


def restore_latest_backup(config: Dict[str, Any], delay_seconds: float = 0.0) -> Dict[str, Any]:
    """Restore the latest cloud backup onto this computer.

    This is intentionally whole-file restore, not two-way merge. MaINbox should be
    closed before the database is replaced. The caller may pass a short delay so
    the GUI has time to exit after launching this command.
    """
    if not bool(config.get("allow_cloud_restore", True)):
        raise RuntimeError("Cloud restore is disabled in mainbox_cloud_bridge_config.json.")
    if delay_seconds and delay_seconds > 0:
        write_status(config, "restore_waiting_for_mainbox_to_close", f"Restore will start in {delay_seconds:.0f} seconds.")
        time.sleep(delay_seconds)

    write_status(config, "restore_downloading", "Downloading latest cloud backup.")
    download = download_latest_backup(config)
    zip_path = Path(download["downloaded_zip"])

    restore_dir = Path(config.get("restore_dir") or default_app_data_dir() / "cloud_bridge_restores")
    extract_dir = restore_dir / f"extract_{utc_stamp()}"
    candidate_db = validate_snapshot_zip(zip_path, extract_dir)

    main_db = Path(config.get("main_db_file") or default_app_data_dir() / "mainbox_data.db")
    main_db.parent.mkdir(parents=True, exist_ok=True)

    safety_zip = ""
    if main_db.exists():
        try:
            safety_zip = str(make_snapshot(config, reason="pre_cloud_restore_safety_backup"))
        except Exception:
            # If MaINbox was already closed but SQLite backup still fails, make a raw copy fallback.
            try:
                fallback = restore_dir / f"pre_restore_mainbox_data_{utc_stamp()}.db"
                shutil.copy2(main_db, fallback)
                safety_zip = str(fallback)
            except Exception:
                safety_zip = ""

    backup_existing = ""
    if main_db.exists():
        backup_existing_path = restore_dir / f"replaced_mainbox_data_{utc_stamp()}.db"
        shutil.copy2(main_db, backup_existing_path)
        backup_existing = str(backup_existing_path)

    # Replace the local database atomically where possible.
    # Windows may keep SQLite locked for a few seconds while MaINbox finishes closing,
    # so retry instead of failing immediately.
    temp_target = main_db.with_suffix(".db.restore_tmp")
    shutil.copy2(candidate_db, temp_target)
    replace_deadline = time.time() + float(config.get("restore_replace_retry_seconds", 45) or 45)
    last_replace_error = None
    while True:
        try:
            os.replace(temp_target, main_db)
            break
        except PermissionError as exc:
            last_replace_error = exc
            if time.time() >= replace_deadline:
                try:
                    if temp_target.exists():
                        temp_target.unlink()
                except Exception:
                    pass
                raise RuntimeError(
                    "Could not replace mainbox_data.db because Windows still has it locked. "
                    "Please make sure every MaINbox window is closed, then run Restore Latest Cloud Backup again."
                ) from exc
            write_status(
                config,
                "restore_waiting_for_database_unlock",
                "Waiting for MaINbox to release mainbox_data.db before restoring.",
                restore_waiting_on=str(main_db),
            )
            time.sleep(2)
        except Exception as exc:
            try:
                if temp_target.exists():
                    temp_target.unlink()
            except Exception:
                pass
            raise

    # v3.8.31: the local database is now the restored copy. Its app_state rows still
    # carry the updated_at timestamps from when the backup was taken, which are older
    # than any local JSON file edited after that backup. Re-stamp them to "now" and drop
    # a restore marker so MaINbox treats the restored rows as the newest copy and a
    # per-row freshness check cannot let stale local JSON overwrite the restore. This
    # runs only after the replace succeeded above, so a failure here is non-fatal.
    restore_stamp = stamp_restored_db_freshness(main_db, source_zip=str(zip_path))

    # v4.3: delete machine-local app_state rows (Outlook EntryIDs, COM-bound scanner
    # dedup sets) that must not carry over to a different machine.  MaINbox rebuilds
    # these from the local Outlook profile on first launch.  Without this step,
    # MaINbox deadlocks on startup trying to resolve ~3-4k EntryIDs that do not exist
    # in the local COM session, permanently freezing the UI after ~30 seconds.
    sanitize_info = sanitize_machine_local_app_state(main_db)

    # v4.4: drop MaINbox's own post-restore marker so the next launch deterministically
    # runs its restore sequence (make restored DB authoritative -> Outlook Send/Receive
    # -> refresh -> EntryID re-link) no matter how this restore was started.
    marker_written = False
    try:
        app_dir = Path(config.get("app_data_dir") or default_app_data_dir())
        app_dir.mkdir(parents=True, exist_ok=True)
        write_json_file(app_dir / "mainbox_post_restore_outlook_sync.json", {
            "requested_at": now_iso(),
            "reason": f"cloud bridge restore ({BRIDGE_VERSION})",
        })
        marker_written = True
    except Exception:
        traceback.print_exc()

    src_manifest = download.get("manifest") or {}
    # Keep the extracted files available for inspection, but do not force JSON over DB.
    result = {
        "downloaded_zip": str(zip_path),
        "backup_cloud_path": download.get("backup_cloud_path", ""),
        "safety_backup": safety_zip,
        "replaced_existing_db_copy": backup_existing,
        "restored_db": str(main_db),
        "extract_dir": str(extract_dir),
        "restamp": restore_stamp,
        "sanitize": sanitize_info,
        "restored_from_device_id": str(src_manifest.get("device_id", "")),
        "restored_from_hostname": str(src_manifest.get("hostname", "")),
        "mainbox_marker_written": marker_written,
    }
    write_status(
        config,
        "restore_complete_restart_required",
        "Latest cloud backup restored. Reopen MaINbox to load the restored data.",
        restore_completed_at=now_iso(),
        last_downloaded_backup_zip=str(zip_path),
        last_restore_safety_backup=safety_zip,
        last_restored_from_zip=str(zip_path),
        last_restored_cloud_path=download.get("backup_cloud_path", ""),
        last_restored_from_device_id=str(src_manifest.get("device_id", "")),
        last_restored_from_hostname=str(src_manifest.get("hostname", "")),
        last_restored_backup_uploaded_at=str(src_manifest.get("uploaded_at", "")),
        restored_db=str(main_db),
        restore_extract_dir=str(extract_dir),
        restored_app_state_rows=restore_stamp.get("app_state_rows_restamped", 0),
        restore_marker_written=restore_stamp.get("restore_marker_written", False),
        mainbox_post_restore_marker_written=marker_written,
        machine_local_rows_deleted=sanitize_info.get("machine_local_rows_deleted", 0),
        machine_local_keys_deleted=sanitize_info.get("machine_local_keys_deleted", []),
    )
    return result

def upload_to_supabase(config: Dict[str, Any], zip_path: Path) -> str:
    prefix = str(config.get("cloud_prefix") or "mainbox").strip("/")
    device_id = str(config.get("device_id") or "mainbox-device").strip("/")
    object_path = f"{prefix}/{device_id}/{zip_path.name}"
    return supabase_upload_bytes(config, object_path, zip_path.read_bytes(), "application/zip")


def safe_state_filename(state_key: str) -> str:
    name = "".join(ch if ch.isalnum() or ch in "_-" else "_" for ch in str(state_key or "state"))
    return (name[:120] or "state") + ".json"


def read_app_state_rows(db_file: Path) -> Dict[str, Any]:
    """Read the app_state JSON rows for Phase 2 one-way cloud mirror."""
    rows: Dict[str, Any] = {}
    if not db_file.exists():
        return rows
    conn = sqlite3.connect(f"file:{db_file}?mode=ro", uri=True, timeout=30)
    try:
        for state_key, json_value, updated_at in conn.execute("SELECT state_key, json_value, updated_at FROM app_state ORDER BY state_key"):
            try:
                parsed = json.loads(json_value)
            except Exception:
                parsed = json_value
            rows[str(state_key)] = {
                "state_key": str(state_key),
                "updated_at": str(updated_at or ""),
                "data": parsed,
            }
    finally:
        conn.close()
    return rows


def upload_latest_json_mirror(config: Dict[str, Any], backup_cloud_path: str, reason: str) -> Dict[str, Any]:
    """Upload read-only latest JSON state files plus a manifest. Does not change local data."""
    if not bool(config.get("upload_latest_json_mirror", True)):
        return {"enabled": False, "uploaded_count": 0, "manifest_path": ""}

    prefix = str(config.get("cloud_prefix") or "mainbox").strip("/")
    device_id = str(config.get("device_id") or "mainbox-device").strip("/")
    latest_prefix = str(config.get("latest_mirror_prefix") or "latest").strip("/")
    db_file = Path(config.get("main_db_file") or default_app_data_dir() / "mainbox_data.db")
    app_state = read_app_state_rows(db_file)
    uploaded_files = []

    for state_key, payload in app_state.items():
        file_name = safe_state_filename(state_key)
        object_path = f"{prefix}/{device_id}/{latest_prefix}/app_state/{file_name}"
        body = json.dumps(payload, indent=2, default=str).encode("utf-8")
        cloud_path = supabase_upload_bytes(config, object_path, body, "application/json")
        uploaded_files.append({
            "state_key": state_key,
            "cloud_path": cloud_path,
            "bytes": len(body),
            "updated_at": payload.get("updated_at", ""),
        })

    manifest = {
        "phase": "phase_2_5_backup_mirror_restore",
        "uploaded_at": now_iso(),
        "reason": reason,
        "bridge_version": BRIDGE_VERSION,
        "device_id": device_id,
        "hostname": (socket.gethostname() or ""),  # v4.4: human-readable machine identity
        "app_version": config.get("app_version", "v3"),
        "source_db": str(db_file),
        "backup_cloud_path": backup_cloud_path,
        "app_state_count": len(app_state),
        "app_state_files": uploaded_files,
        "note": "Read-only latest cloud mirror. MaINbox still uses the local SQLite database as source of truth.",
    }
    manifest_object_path = f"{prefix}/{device_id}/{latest_prefix}/manifest.json"
    manifest_path = supabase_upload_bytes(
        config,
        manifest_object_path,
        json.dumps(manifest, indent=2, default=str).encode("utf-8"),
        "application/json",
    )
    return {
        "enabled": True,
        "uploaded_count": len(uploaded_files),
        "manifest_path": manifest_path,
        "files": uploaded_files,
    }


def perform_backup(config: Dict[str, Any], reason: str) -> Dict[str, Any]:
    zip_path = make_snapshot(config, reason=reason)
    cleanup_old_snapshots(config)

    result = {
        "snapshot_zip": str(zip_path),
        "snapshot_size_bytes": zip_path.stat().st_size,
        "uploaded": False,
        "cloud_path": "",
    }

    if bool(config.get("enabled", False)):
        cloud_path = upload_to_supabase(config, zip_path)
        result["uploaded"] = True
        result["cloud_path"] = cloud_path
        mirror = upload_latest_json_mirror(config, cloud_path, reason=reason)
        result["mirror_uploaded"] = bool(mirror.get("enabled"))
        result["mirror_manifest_path"] = mirror.get("manifest_path", "")
        result["mirror_app_state_count"] = mirror.get("uploaded_count", 0)
    return result


def consume_request_if_new(config: Dict[str, Any], last_request_id: Optional[str]) -> tuple[Optional[str], Optional[Dict[str, Any]]]:
    request_file = Path(config.get("request_file") or default_app_data_dir() / "mainbox_cloud_backup_request.json")
    req = load_json_file(request_file, None)
    if not isinstance(req, dict):
        return last_request_id, None
    if req.get("handled_at"):
        return last_request_id, None
    request_id = str(req.get("request_id") or "")
    if not request_id or request_id == last_request_id:
        return last_request_id, None
    return request_id, req


def mark_request_handled(config: Dict[str, Any], request_id: str, status: str) -> None:
    try:
        request_file = Path(config.get("request_file") or default_app_data_dir() / "mainbox_cloud_backup_request.json")
        req = load_json_file(request_file, {})
        if isinstance(req, dict) and str(req.get("request_id") or "") == str(request_id or ""):
            req["handled_at"] = now_iso()
            req["handled_by"] = BRIDGE_VERSION
            req["handled_status"] = status
            write_json_file(request_file, req)
    except Exception:
        pass


def latest_pending_request(config: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    request_file = Path(config.get("request_file") or default_app_data_dir() / "mainbox_cloud_backup_request.json")
    req = load_json_file(request_file, None)
    if isinstance(req, dict) and req.get("request_id") and not req.get("handled_at"):
        return req
    return None


def run_once(config_path: Path, reason: str = "manual_once") -> int:
    config = load_config(config_path)
    pending_req = latest_pending_request(config)
    request_id = ""
    if pending_req:
        request_id = str(pending_req.get("request_id") or "")
        reason = str(pending_req.get("reason") or reason)
    try:
        write_status(config, "running", "One-time backup started.", last_request_id=request_id)
        result = perform_backup(config, reason=reason)
        status = "uploaded" if result.get("uploaded") else "local_snapshot_created_cloud_disabled"
        message = "Backup uploaded to Supabase." if result.get("uploaded") else "Local snapshot created. Cloud upload disabled until enabled=true and Supabase credentials are configured."
        write_status(
            config,
            status,
            message,
            last_snapshot_zip=result.get("snapshot_zip"),
            last_upload_path=result.get("cloud_path"),
            snapshot_size_bytes=result.get("snapshot_size_bytes"),
            last_mirror_manifest_path=result.get("mirror_manifest_path", ""),
            mirror_app_state_count=result.get("mirror_app_state_count", 0),
            last_request_id=request_id,
        )
        if request_id:
            mark_request_handled(config, request_id, status)
        print(f"[{now_iso()}] {message}")
        print(f"Snapshot: {result.get('snapshot_zip')}")
        if result.get("cloud_path"):
            print(f"Cloud path: {result.get('cloud_path')}")
        if result.get("mirror_manifest_path"):
            print(f"Latest mirror manifest: {result.get('mirror_manifest_path')}")
        return 0
    except Exception as exc:
        traceback.print_exc()
        write_status(config, "error", str(exc), error=traceback.format_exc(), last_request_id=request_id)
        if request_id:
            mark_request_handled(config, request_id, "error")
        return 1


def run_loop(config_path: Path) -> int:
    last_request_id = None
    last_auto_backup = 0.0
    print(f"[{now_iso()}] MaINbox Cloud Bridge started.")
    print(f"Config: {config_path}")

    while True:
        config = load_config(config_path)
        interval_minutes = float(config.get("auto_backup_minutes", 30) or 30)
        interval_seconds = max(60.0, interval_minutes * 60.0)
        now = time.time()

        try:
            last_request_id, req = consume_request_if_new(config, last_request_id)
            if req:
                print(f"[{now_iso()}] Backup requested by MaINbox: {req.get('reason', 'manual')}")
                result = perform_backup(config, reason=str(req.get("reason") or "manual_request"))
                status = "uploaded" if result.get("uploaded") else "local_snapshot_created_cloud_disabled"
                message = "Manual backup uploaded to Supabase." if result.get("uploaded") else "Manual local snapshot created. Cloud upload disabled or not configured."
                write_status(
                    config,
                    status,
                    message,
                    last_snapshot_zip=result.get("snapshot_zip"),
                    last_upload_path=result.get("cloud_path"),
                    snapshot_size_bytes=result.get("snapshot_size_bytes"),
                    last_mirror_manifest_path=result.get("mirror_manifest_path", ""),
                    mirror_app_state_count=result.get("mirror_app_state_count", 0),
                    last_request_id=last_request_id,
                )
                mark_request_handled(config, str(last_request_id or ""), status)
                last_auto_backup = now
                print(f"[{now_iso()}] Manual backup complete: {result.get('snapshot_zip')}")
                if result.get("cloud_path"):
                    print(f"[{now_iso()}] Uploaded to Supabase Storage: {result.get('cloud_path')}")
                if result.get("mirror_manifest_path"):
                    print(f"[{now_iso()}] Latest read-only mirror manifest: {result.get('mirror_manifest_path')}")

            elif now - last_auto_backup >= interval_seconds:
                print(f"[{now_iso()}] Scheduled backup starting.")
                result = perform_backup(config, reason="scheduled")
                status = "uploaded" if result.get("uploaded") else "local_snapshot_created_cloud_disabled"
                write_status(
                    config,
                    status,
                    "Scheduled backup completed.",
                    last_snapshot_zip=result.get("snapshot_zip"),
                    last_upload_path=result.get("cloud_path"),
                    snapshot_size_bytes=result.get("snapshot_size_bytes"),
                    last_mirror_manifest_path=result.get("mirror_manifest_path", ""),
                    mirror_app_state_count=result.get("mirror_app_state_count", 0),
                )
                last_auto_backup = now
                print(f"[{now_iso()}] Scheduled backup complete: {result.get('snapshot_zip')}")
                if result.get("cloud_path"):
                    print(f"[{now_iso()}] Uploaded to Supabase Storage: {result.get('cloud_path')}")
                if result.get("mirror_manifest_path"):
                    print(f"[{now_iso()}] Latest read-only mirror manifest: {result.get('mirror_manifest_path')}")

            else:
                write_status(config, "watching", "Bridge is watching for backup requests.")

        except KeyboardInterrupt:
            print("Stopping bridge.")
            return 0
        except Exception as exc:
            traceback.print_exc()
            write_status(config, "error", str(exc), error=traceback.format_exc())

        time.sleep(10)


def run_download_latest(config_path: Path) -> int:
    config = load_config(config_path)
    try:
        write_status(config, "download_running", "Downloading latest cloud backup.")
        result = download_latest_backup(config)
        write_status(
            config,
            "download_complete",
            "Latest cloud backup downloaded. Local MaINbox data was not changed.",
            last_downloaded_backup_zip=result.get("downloaded_zip"),
            last_downloaded_backup_size_bytes=result.get("downloaded_size_bytes"),
            last_downloaded_cloud_path=result.get("backup_cloud_path"),
        )
        print(f"[{now_iso()}] Latest cloud backup downloaded: {result.get('downloaded_zip')}")
        return 0
    except Exception as exc:
        traceback.print_exc()
        write_status(config, "download_error", str(exc), error=traceback.format_exc())
        return 1



def launch_mainbox_from_config(config: Dict[str, Any]) -> None:
    """Launch MaINbox from the path stored by the app before restore."""
    app_path = str(config.get("mainbox_app_path") or "").strip().strip('"')
    if not app_path:
        raise RuntimeError("MaINbox app path is not saved in the cloud bridge config yet.")
    p = Path(app_path)
    if not p.exists():
        raise RuntimeError(f"Saved MaINbox app path was not found:\n{app_path}")
    if os.name == "nt":
        os.startfile(str(p))
        return
    python_exe = str(config.get("bridge_python_executable") or sys.executable or "python").strip().strip('"') or "python"
    subprocess.Popen([python_exe, str(p)], cwd=str(p.parent) or None)


def run_restore_latest_with_ui(config_path: Path, delay_seconds: float = 0.0) -> int:
    """Run restore with a small visible progress window.

    This is used when MaINbox launches the bridge for a restore. The popup remains
    visible while MaINbox closes and the database is replaced, then offers to
    relaunch MaINbox or close the popup.
    """
    try:
        import tkinter as tk
        from tkinter import messagebox, ttk
    except Exception:
        # If Tk is unavailable for any reason, restore still works in normal CLI mode.
        return run_restore_latest(config_path, delay_seconds=delay_seconds)

    config = load_config(config_path)
    root = tk.Tk()
    root.title("MaINbox Cloud Restore")
    root.geometry("560x250")
    root.resizable(False, False)

    done = {"finished": False, "ok": False, "error": ""}

    frm = ttk.Frame(root, padding=18)
    frm.pack(fill="both", expand=True)

    title_var = tk.StringVar(value="Restoring latest MaINbox cloud backup...")
    detail_var = tk.StringVar(value="Please wait. MaINbox is closing and the bridge is preparing the restore.")

    ttk.Label(frm, textvariable=title_var, font=("Segoe UI", 12, "bold")).pack(anchor="w")
    ttk.Label(frm, textvariable=detail_var, wraplength=510, justify="left").pack(anchor="w", pady=(10, 8))

    progress = ttk.Progressbar(frm, mode="indeterminate")
    progress.pack(fill="x", pady=(8, 14))
    progress.start(12)

    button_frame = ttk.Frame(frm)
    button_frame.pack(fill="x", side="bottom", pady=(10, 0))

    def finish_ui(ok: bool, error_text: str = "") -> None:
        done["finished"] = True
        done["ok"] = ok
        done["error"] = error_text
        progress.stop()
        progress.pack_forget()
        for child in button_frame.winfo_children():
            child.destroy()
        if ok:
            title_var.set("Cloud restore complete")
            detail_var.set("The latest cloud backup was restored. Launch MaINbox now to load the restored data.")
            try:
                status_file = Path(config.get("status_file") or default_app_data_dir() / "mainbox_cloud_bridge_status.json")
                # v4.4: this called read_json_file(), which does not exist -- the NameError was
                # swallowed and keep={} wiped every detail field from the final status
                # (the blank 'Last restored from' dialog). The helper is load_json_file().
                prev = load_json_file(status_file, {}) if status_file.exists() else {}
                keep = {
                    k: v for k, v in prev.items()
                    if k not in {"status", "message", "last_run_at", "bridge_version", "config"}
                }
            except Exception:
                keep = {}
            write_status(config, "restore_complete", "Latest cloud backup restored. Ready to launch MaINbox.", **keep)
            ttk.Button(button_frame, text="Launch MaINbox", command=launch_and_close).pack(side="right", padx=(8, 0))
            ttk.Button(button_frame, text="Cancel", command=root.destroy).pack(side="right")
        else:
            title_var.set("Cloud restore could not complete")
            detail_var.set(error_text or "The restore did not complete. Check Cloud Bridge Status for details.")
            ttk.Button(button_frame, text="Close", command=root.destroy).pack(side="right")

    def launch_and_close() -> None:
        try:
            launch_mainbox_from_config(config)
            root.destroy()
        except Exception as exc:
            messagebox.showerror("Launch MaINbox", str(exc))

    def worker() -> None:
        try:
            restore_latest_backup(config, delay_seconds=delay_seconds)
            root.after(0, lambda: finish_ui(True, ""))
        except Exception as exc:
            traceback.print_exc()
            write_status(config, "restore_error", str(exc), error=traceback.format_exc())
            root.after(0, lambda e=str(exc): finish_ui(False, e))

    def poll_status() -> None:
        if done["finished"]:
            return
        try:
            status_file = Path(config.get("status_file") or default_app_data_dir() / "mainbox_cloud_bridge_status.json")
            if status_file.exists():
                data = load_json_file(status_file, {})  # v4.4: was read_json_file (NameError) -- live progress never updated
                st = str(data.get("status") or "")
                msg = str(data.get("message") or "")
                if st:
                    pretty = st.replace("_", " ").capitalize()
                    title_var.set(pretty)
                if msg:
                    detail_var.set(msg)
        except Exception:
            pass
        root.after(750, poll_status)

    import threading
    t = threading.Thread(target=worker, daemon=True)
    t.start()
    root.after(500, poll_status)
    root.mainloop()
    return 0 if done.get("ok") else 1


def run_restore_latest(config_path: Path, delay_seconds: float = 0.0) -> int:
    config = load_config(config_path)
    try:
        result = restore_latest_backup(config, delay_seconds=delay_seconds)
        print(f"[{now_iso()}] Latest cloud backup restored: {result.get('downloaded_zip')}")
        return 0
    except Exception as exc:
        traceback.print_exc()
        write_status(config, "restore_error", str(exc), error=traceback.format_exc())
        return 1


def main(argv: Optional[list[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="MaINbox Cloud Bridge Phase 2.5 backup, one-way mirror, and restore helper for Supabase.")
    parser.add_argument("--config", default=str(default_config_path()), help="Path to mainbox_cloud_bridge_config.json")
    parser.add_argument("--once", action="store_true", help="Run one backup and exit")
    parser.add_argument("--download-latest", action="store_true", help="Download the latest cloud backup and exit without changing local data")
    parser.add_argument("--restore-latest", action="store_true", help="Restore the latest cloud backup to this computer and exit")
    parser.add_argument("--restore-ui", action="store_true", help="Show a small restore progress popup and launch button")
    parser.add_argument("--delay", type=float, default=0.0, help="Seconds to wait before restore, giving MaINbox time to close")
    args = parser.parse_args(argv)

    config_path = Path(args.config).expanduser().resolve()
    config_path.parent.mkdir(parents=True, exist_ok=True)

    if args.download_latest:
        return run_download_latest(config_path)
    if args.restore_latest:
        if args.restore_ui:
            return run_restore_latest_with_ui(config_path, delay_seconds=float(args.delay or 0.0))
        return run_restore_latest(config_path, delay_seconds=float(args.delay or 0.0))
    if args.once:
        return run_once(config_path)
    return run_loop(config_path)


if __name__ == "__main__":
    raise SystemExit(main())
