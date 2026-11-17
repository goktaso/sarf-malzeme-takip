"""SQL Server connection adapter for the inventory queries."""

import os
import re
import site
import sqlite3
from functools import lru_cache
from pathlib import Path

ROOT = Path(__file__).resolve().parent
IDENTITIES = {"departments", "people", "product_groups", "products", "variants",
              "documents", "movements", "definition_versions", "stock_counts", "app_period_grants"}


@lru_cache(maxsize=1)
def load_runtime():
    directory = ROOT / "data" / "runtime"
    if directory.is_dir():
        site.addsitedir(str(directory))


class Row:
    def __init__(self, columns, values):
        self.values = tuple(values)
        self.columns = columns
        self.mapping = dict(zip(columns, values))

    def keys(self):
        return self.columns

    def __iter__(self):
        return iter(self.values)

    def __getitem__(self, key):
        return self.mapping[key] if isinstance(key, str) else self.values[key]


class Cursor:
    def __init__(self, cursor, identity=False):
        self.cursor = cursor
        self.columns = [column[0] for column in cursor.description or ()]
        self.lastrowid = int(cursor.fetchone()[0]) if identity else None

    def fetchone(self):
        result = self.cursor.fetchone()
        return Row(self.columns, result) if result is not None else None

    def __iter__(self):
        return (Row(self.columns, row) for row in self.cursor)


def sql_json(match):
    column, field = match.groups()
    value = f"JSON_VALUE({column}, '$.{field}')"
    return f"TRY_CONVERT(bigint, {value})" if field in {"department_id", "active"} else value


class Connection:
    def __init__(self, connection, driver):
        self.raw = connection
        self.driver = driver

    def execute(self, sql, params=()):
        text = sql.strip()
        if text in {"BEGIN", "BEGIN IMMEDIATE"}:
            try:
                self.raw.autocommit = True
                self.raw.execute("SET TRANSACTION ISOLATION LEVEL " +
                                 ("SNAPSHOT" if text == "BEGIN" else "READ COMMITTED"))
                self.raw.autocommit = False
                if text == "BEGIN IMMEDIATE":
                    self.raw.execute("""
                        IF @@TRANCOUNT=0 BEGIN TRANSACTION;
                        DECLARE @result int;
                        EXEC @result=sp_getapplock @Resource='SarfMalzemeTakip:write',
                            @LockMode='Exclusive', @LockOwner='Transaction', @LockTimeout=12000;
                        IF @result < 0 THROW 51000, 'Application write lock timed out', 1;
                    """)
            except self.driver.Error as error:
                raise sqlite3.OperationalError(str(error)) from error
            return
        text = re.sub(r"json_extract\(([\w.]+),'\$\.(\w+)'\)", sql_json, text)
        text = re.sub(r"\bLIMIT 1\b", "OFFSET 0 ROWS FETCH NEXT 1 ROWS ONLY", text)
        match = re.match(r"INSERT INTO (\w+)\s*\(", text, re.IGNORECASE)
        identity = bool(match and match[1].lower() in IDENTITIES)
        if identity:
            text = re.sub(r"\)\s*VALUES\s*\(", ") OUTPUT INSERTED.id VALUES (", text, count=1,
                          flags=re.IGNORECASE)
        try:
            return Cursor(self.raw.execute(text, params), identity)
        except self.driver.IntegrityError as error:
            raise sqlite3.IntegrityError(str(error)) from error
        except self.driver.Error as error:
            raise sqlite3.OperationalError(str(error)) from error

    def commit(self):
        self.raw.commit()

    def rollback(self):
        self.raw.rollback()

    def close(self):
        self.raw.close()


def connect(master=False, autocommit=False):
    load_runtime()
    import pyodbc

    instance = os.environ.get("SARF_SQL_SERVER", "lpc:.")
    database = "master" if master else os.environ.get("SARF_SQL_DATABASE", "SarfMalzemeTakip")
    driver = os.environ.get("SARF_SQL_DRIVER", "ODBC Driver 18 for SQL Server")
    trust = "yes" if instance == "lpc:." else "no"
    try:
        raw = pyodbc.connect(
            f"DRIVER={{{driver}}};SERVER={instance};DATABASE={database};Trusted_Connection=yes;"
            f"Encrypt=yes;TrustServerCertificate={trust};APP=SarfMalzemeTakip;",
            autocommit=autocommit, timeout=5)
    except pyodbc.Error as error:
        raise sqlite3.OperationalError(str(error)) from error
    raw.timeout = 25
    return Connection(raw, pyodbc)
