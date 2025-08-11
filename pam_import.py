import argparse
import os
import re
import time
from getpass import getpass

import pandas as pd
import requests
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


def normalize_name(s: str) -> str:
    s = (s or "").strip().replace(" ", "_")
    return re.sub(r"[^A-Za-z0-9_\-\.]", "_", s)


def ip_suffix(ip: str) -> str:
    ip = (ip or "").strip()
    if ":" in ip:
        parts = [p for p in ip.split(":") if p]
        if len(parts) >= 2:
            return f"{parts[-2]}.{parts[-1]}"
        return ip.replace(":", ".")
    parts = [p for p in ip.split(".") if p]
    if len(parts) >= 2:
        return f"{parts[-2]}.{parts[-1]}"
    return ip.replace(".", ".")


def build_service_payload(service_name: str):
    s = (service_name or "").strip().upper()
    if s == "SSH":
        return {
            "port": 22,
            "protocol": "SSH",
            "subprotocols": [
                "SSH_SHELL_SESSION",
                "SSH_REMOTE_COMMAND",
                "SSH_SCP_UP",
                "SSH_SCP_DOWN",
                "SFTP_SESSION",
            ],
            "service_name": "SSH",
            "connection_policy": "SSH",
        }
    if s == "RDP":
        return {
            "port": 3389,
            "protocol": "RDP",
            "subprotocols": [
                "RDP_CLIPBOARD_UP",
                "RDP_CLIPBOARD_DOWN",
                "RDP_CLIPBOARD_FILE",
                "RDP_PRINTER",
                "RDP_COM_PORT",
                "RDP_DRIVE",
                "RDP_SMARTCARD",
                "RDP_AUDIO_OUTPUT",
                "RDP_AUDIO_INPUT",
            ],
            "service_name": "RDP",
            "connection_policy": "RDP",
        }
    return None


def make_session(host: str, api_version: str, username: str, password: str, insecure: bool):
    s = requests.Session()
    s.verify = not insecure
    s.auth = (username, password)
    s.headers.update({"Content-Type": "application/json"})
    base = f"https://{host}/api/{api_version}"
    return s, base


def iter_devices(session: requests.Session, base_url: str, limit: int = 1000, timeout: float = 60.0):
    url = f"{base_url}/devices"
    offset = 0
    while True:
        r = session.get(url, params={"offset": offset, "limit": limit}, timeout=timeout)
        if r.status_code != 200:
            break
        items = r.json() or []
        if not isinstance(items, list):
            break
        for it in items:
            yield it
        if len(items) < limit:
            break
        offset += limit


def resolve_device_id(session: requests.Session, base_url: str, name: str, ip: str, retries: int = 6, delay: float = 1.2):
    for _ in range(retries):
        for d in iter_devices(session, base_url):
            if d.get("device_name") == name and (not ip or d.get("host") == ip):
                return d.get("id")
        time.sleep(delay)
    return None


def get_service_id(session: requests.Session, base_url: str, device_id: str, service_name_upper: str, timeout: float = 30.0):
    url = f"{base_url}/devices/{device_id}/services"
    r = session.get(url, timeout=timeout)
    if r.status_code == 200:
        for s in r.json() or []:
            if (s.get("service_name") or "").upper() == service_name_upper:
                return s.get("id")
    return None


def get_or_create_target_group(session: requests.Session, base_url: str, group_name: str, timeout: float = 30.0):
    url = f"{base_url}/targetgroups"
    r = session.get(url, timeout=timeout)
    if r.status_code == 200:
        for g in r.json() or []:
            if g.get("group_name") == group_name:
                return g.get("id")
    r = session.post(url, json={"group_name": group_name, "description": group_name}, timeout=timeout)
    if r.status_code in (200, 201):
        body = r.json() if r.content else {}
        return body.get("id")
    if r.status_code == 204:
        r2 = session.get(url, timeout=timeout)
        if r2.status_code == 200:
            for g in r2.json() or []:
                if g.get("group_name") == group_name:
                    return g.get("id")
    return None


def add_interactive_login_to_group_by_names(session: requests.Session, base_url: str, group_id: str, device_name: str, service_name: str, timeout: float = 60.0):
    url = f"{base_url}/targetgroups/{group_id}?force=false"
    body = {
        "session": {
            "interactive_logins": [
                {"device": device_name, "service": service_name, "application": None}
            ]
        }
    }
    r = session.put(url, json=body, timeout=timeout)
    return r.status_code in (200, 204), (r.status_code, r.text)


def main():
    p = argparse.ArgumentParser(prog="bastion_pam_import")
    p.add_argument("--host", default=os.getenv("BASTION_HOST"))
    p.add_argument("--api-version", default=os.getenv("BASTION_API_VERSION", "v3.12"))
    p.add_argument("--username", default=os.getenv("BASTION_USERNAME"))
    p.add_argument("--password", default=os.getenv("BASTION_PASSWORD"))
    p.add_argument("--excel", default="PAM Access 4.xlsx")
    p.add_argument("--sheet", default=None)
    p.add_argument("--ip-column", default="Destination ip")
    p.add_argument("--name-column", default="Server Name")
    p.add_argument("--service-column", default="Service")
    p.add_argument("--group", default="IT Services")
    p.add_argument("--no-group", action="store_true")
    p.add_argument("--description", default="IT Services")
    p.add_argument("--insecure", action="store_true")
    p.add_argument("--timeout", type=float, default=60.0)
    p.add_argument("--limit", type=int, default=1000)
    p.add_argument("--csv-log", default=None)
    args = p.parse_args()

    if not args.host:
        raise SystemExit("--host is required (or set BASTION_HOST)")
    if not args.username:
        raise SystemExit("--username is required (or set BASTION_USERNAME)")
    password = args.password or getpass("Password: ")

    session, base_url = make_session(args.host, args.api_version, args.username, password, args.insecure)

    df = pd.read_excel(args.excel, sheet_name=args.sheet)
    for c in (args.ip_column, args.name_column):
        if c not in df.columns:
            raise SystemExit(f"Excel missing column: {c}. Present: {list(df.columns)}")
    has_service_col = args.service_column in df.columns

    group_id = None
    if not args.no_group and args.group:
        group_id = get_or_create_target_group(session, base_url, args.group)
        if group_id:
            print(f"✅ Using target group '{args.group}' (ID: {group_id})")
        else:
            print(f"⚠️ could not get/create group '{args.group}'")

    csv_rows = []
    for idx, row in df.iterrows():
        ip_raw = str(row[args.ip_column]).strip()
        srv_name_raw = str(row[args.name_column]).strip()
        service_name = (str(row[args.service_column]).strip() if has_service_col else "SSH") or "SSH"
        if not ip_raw or not srv_name_raw:
            print(f"[Row {idx}] ❌ Missing IP/Server Name. Skipped.")
            continue

        device_name = f"{normalize_name(srv_name_raw)}_{ip_suffix(ip_raw)}"
        payload = build_service_payload(service_name)
        if not payload:
            print(f"[{device_name}] ⚠️ Unsupported service '{service_name}'. Skipped.")
            continue

        dev_url = f"{base_url}/devices"
        dev_data = {"device_name": device_name, "alias": device_name, "description": args.description, "host": ip_raw}
        dev_resp = session.post(dev_url, json=dev_data, timeout=args.timeout)
        device_id = None
        device_status = f"{dev_resp.status_code}"
        if dev_resp.status_code in (200, 201):
            body = dev_resp.json() or {}
            device_id = body.get("id")
            print(f"[{device_name}] ✅ Device created: ID={device_id}")
        elif dev_resp.status_code == 204:
            print(f"[{device_name}] ✅ Device created (204). Resolving ID...")
            device_id = resolve_device_id(session, base_url, device_name, ip_raw)
        elif dev_resp.status_code == 409:
            device_id = resolve_device_id(session, base_url, device_name, ip_raw)
            print(f"[{device_name}] ℹ️ Device already exists. ID={device_id}")
        else:
            print(f"[{device_name}] ❌ Device create failed: {dev_resp.status_code} - {dev_resp.text}")
            continue

        if not device_id:
            print(f"[{device_name}] ❌ Could not resolve device ID. Skipped.")
            continue

        svc_url = f"{base_url}/devices/{device_id}/services"
        svc_resp = session.post(svc_url, json=payload, timeout=args.timeout)
        service_id = None
        service_status = f"{svc_resp.status_code}"
        svc_upper = payload["service_name"].upper()
        if svc_resp.status_code in (200, 201):
            body = svc_resp.json() or {}
            service_id = body.get("id")
            print(f"[{device_name}] ✅ Service {svc_upper} created: ID={service_id}")
        elif svc_resp.status_code in (204, 409):
            service_id = get_service_id(session, base_url, device_id, svc_upper, timeout=args.timeout)
            if service_id:
                print(f"[{device_name}] ✅ Service {svc_upper} available: ID={service_id}")
            else:
                print(f"[{device_name}] ❌ Could not resolve {svc_upper} service after {svc_resp.status_code}.")
                continue
        else:
            print(f"[{device_name}] ❌ Service create failed: {svc_resp.status_code} - {svc_resp.text}")
            continue

        added_to_group = False
        if group_id:
            ok, err = add_interactive_login_to_group_by_names(session, base_url, group_id, device_name, svc_upper, timeout=args.timeout)
            if ok:
                print(f"[{device_name}] 🏷️ Added to group '{args.group}'")
                added_to_group = True
            else:
                print(f"[{device_name}] ⚠️ Could not add to group '{args.group}': {err[0]} - {err[1]}")

        if args.csv_log:
            csv_rows.append({
                "row": idx,
                "device_name": device_name,
                "ip": ip_raw,
                "service": svc_upper,
                "device_id": device_id,
                "service_id": service_id,
                "device_status": device_status,
                "service_status": service_status,
                "group": args.group if group_id else "",
                "group_added": added_to_group,
            })

    if args.csv_log and csv_rows:
        import csv
        fieldnames = list(csv_rows[0].keys())
        with open(args.csv_log, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(csv_rows)


if __name__ == "__main__":
    main()
