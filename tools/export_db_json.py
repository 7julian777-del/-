import json
import sqlite3
from datetime import datetime
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parents[1]
DB_PATH = BASE_DIR / "data" / "app.db"
EXPORT_DIR = BASE_DIR / "exports"
EXPORT_DIR.mkdir(parents=True, exist_ok=True)


def fetch_all(conn, table):
    cur = conn.execute(f"SELECT * FROM {table}")
    rows = cur.fetchall()
    return [dict(row) for row in rows]


def main():
    if not DB_PATH.exists():
        raise SystemExit(f"未找到数据库: {DB_PATH}")

    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row

    data = {
        "meta": {
            "app": "kaidan-pwa",
            "exported_at": datetime.now().isoformat(timespec="seconds"),
        },
        "settings": fetch_all(conn, "settings"),
        "customers": fetch_all(conn, "customers"),
        "products": fetch_all(conn, "products"),
        "vehicles": fetch_all(conn, "vehicles"),
        "invoices": fetch_all(conn, "invoices"),
        "invoice_items": fetch_all(conn, "invoice_items"),
        "invoice_audit": fetch_all(conn, "invoice_audit"),
    }

    conn.close()

    out_path = EXPORT_DIR / f"kaidan-export-{datetime.now().strftime('%Y%m%d')}.json"
    out_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"已导出: {out_path}")


if __name__ == "__main__":
    main()
