"""Provision the first Windows application administrator and import the legacy data once."""

import argparse
import os
import re
import sqlite3
from contextlib import closing

import access
import server
import sqlserver

TABLES = ("departments", "people", "product_groups", "products", "variants", "documents",
          "movements", "definition_versions")


def copy_legacy(db, source):
    exists = source.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='products'").fetchone()
    if not exists:
        raise RuntimeError("SQLite kaynak şeması bulunamadı.")
    source_counts = {table: source.execute(f"SELECT COUNT(*) FROM {table}").fetchone()[0] for table in TABLES}
    if any(db.execute(f"SELECT COUNT(*) FROM {table}").fetchone()[0] for table in TABLES):
        raise RuntimeError("SQL Server veritabanında kayıt var. Mevcut kayıtların üzerine aktarım yapılmadı.")
    for table in TABLES:
        if not source_counts[table]:
            continue
        cursor = source.execute(f"SELECT * FROM {table} ORDER BY id")
        columns = [part[0] for part in cursor.description]
        placeholders = ",".join("?" for _ in columns)
        db.raw.execute(f"SET IDENTITY_INSERT dbo.{table} ON")
        try:
            db.raw.cursor().executemany(
                f"INSERT dbo.{table}({','.join('[' + name + ']' for name in columns)}) VALUES ({placeholders})",
                cursor)
        finally:
            db.raw.execute(f"SET IDENTITY_INSERT dbo.{table} OFF")
        if db.execute(f"SELECT COUNT(*) FROM {table}").fetchone()[0] != source_counts[table]:
            raise RuntimeError(f"Aktarım satır sayısı eşleşmedi: {table}")
    return source_counts


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--admin", required=True)
    parser.add_argument("--name", required=True)
    parser.add_argument("--migrate-legacy", action="store_true")
    args = parser.parse_args()
    database = os.environ.get("SARF_SQL_DATABASE", "SarfMalzemeTakip")
    if not re.fullmatch(r"[A-Za-z][A-Za-z0-9_]{0,127}", database):
        raise ValueError("Geçersiz SQL Server veritabanı adı.")
    master = sqlserver.connect(master=True, autocommit=True)
    try:
        master.raw.execute(f"ALTER DATABASE [{database}] SET ALLOW_SNAPSHOT_ISOLATION ON")
    finally:
        master.close()
    sid, account = access.windows_account(args.admin)
    if args.migrate_legacy:
        server.initialize()
    server.SQL_MODE = True
    server.initialize()
    with server.connect() as db:
        db.execute("BEGIN IMMEDIATE")
        existing = db.execute("SELECT * FROM app_users WHERE sid=?", (sid,)).fetchone()
        if existing is None:
            if db.execute("SELECT COUNT(*) FROM app_users").fetchone()[0]:
                raise RuntimeError("İlk kurulum tamamlanmış. Yeni kullanıcıyı yönetim ekranından ekleyin.")
            db.execute("""INSERT INTO app_users(sid,account,display_name,is_admin,active,permissions)
                          VALUES (?,?,?,1,1,?)""", (sid, account, args.name, "[]"))
        elif not existing["is_admin"] or not existing["active"]:
            raise RuntimeError("Hesap zaten var. Yetkileri yönetim ekranından değiştirin.")
        if args.migrate_legacy:
            with closing(sqlite3.connect(server.DB.as_uri() + "?mode=ro", uri=True)) as source:
                source.execute("BEGIN")
                counts = copy_legacy(db, source)
            print("SQLite -> SQL Server:", counts)
    print("Active application administrator:", account)


if __name__ == "__main__":
    main()
