CREATE TABLE IF NOT EXISTS departments (
    id INTEGER PRIMARY KEY, name TEXT NOT NULL UNIQUE,
    active INTEGER NOT NULL DEFAULT 1 CHECK(active IN (0,1))
);
CREATE TABLE IF NOT EXISTS people (
    id INTEGER PRIMARY KEY, code TEXT NOT NULL UNIQUE, name TEXT NOT NULL,
    department_id INTEGER NOT NULL REFERENCES departments(id),
    active INTEGER NOT NULL DEFAULT 1 CHECK(active IN (0,1))
);
CREATE TABLE IF NOT EXISTS product_groups (
    id INTEGER PRIMARY KEY, name TEXT NOT NULL, prefix TEXT NOT NULL UNIQUE,
    module TEXT NOT NULL CHECK(module IN ('sarf','kkd')),
    next_number INTEGER NOT NULL DEFAULT 1,
    size_kind TEXT NOT NULL DEFAULT 'size' CHECK(size_kind IN ('size','shoe'))
);
CREATE TABLE IF NOT EXISTS products (
    id INTEGER PRIMARY KEY, code TEXT NOT NULL UNIQUE, name TEXT NOT NULL,
    module TEXT NOT NULL CHECK(module IN ('sarf','kkd')), category TEXT NOT NULL,
    unit TEXT NOT NULL, scope TEXT NOT NULL CHECK(scope IN ('person','department','both')),
    group_id INTEGER REFERENCES product_groups(id),
    active INTEGER NOT NULL DEFAULT 1 CHECK(active IN (0,1))
);
CREATE TABLE IF NOT EXISTS variants (
    id INTEGER PRIMARY KEY, product_id INTEGER NOT NULL REFERENCES products(id),
    size TEXT NOT NULL, color TEXT NOT NULL DEFAULT '', minimum INTEGER NOT NULL,
    UNIQUE(product_id, size, color)
);
CREATE TABLE IF NOT EXISTS documents (
    id INTEGER PRIMARY KEY, request_id TEXT NOT NULL UNIQUE, kind TEXT NOT NULL,
    day TEXT NOT NULL, purchase_day TEXT, supplier TEXT, reference TEXT, note TEXT,
    department_id INTEGER REFERENCES departments(id), person_id INTEGER REFERENCES people(id),
    source_id INTEGER REFERENCES documents(id)
);
CREATE TABLE IF NOT EXISTS movements (
    id INTEGER PRIMARY KEY, document_id INTEGER NOT NULL REFERENCES documents(id),
    variant_id INTEGER NOT NULL REFERENCES variants(id),
    qty INTEGER NOT NULL, value INTEGER NOT NULL, unit_price INTEGER,
    UNIQUE(document_id, variant_id)
);
CREATE INDEX IF NOT EXISTS movement_variant ON movements(variant_id);
CREATE TABLE IF NOT EXISTS definition_versions (
    id INTEGER PRIMARY KEY,
    entity TEXT NOT NULL CHECK(entity IN ('products','departments','people','product_groups')),
    entity_id INTEGER NOT NULL,
    effective_day TEXT NOT NULL,
    recorded_at TEXT NOT NULL DEFAULT (strftime('%Y-%m-%dT%H:%M:%fZ','now')),
    state TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS definition_as_of ON definition_versions(entity,entity_id,effective_day DESC,id DESC);

CREATE TABLE IF NOT EXISTS stock_counts (
    id INTEGER PRIMARY KEY, day TEXT NOT NULL, note TEXT NOT NULL,
    status TEXT NOT NULL DEFAULT 'draft' CHECK(status IN ('draft','locked','rejected','approved')),
    revision INTEGER NOT NULL DEFAULT 0, owner TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS stock_count_lines (
    count_id INTEGER NOT NULL REFERENCES stock_counts(id), variant_id INTEGER NOT NULL REFERENCES variants(id),
    code TEXT NOT NULL, name TEXT NOT NULL, category TEXT NOT NULL, color TEXT NOT NULL,
    size TEXT NOT NULL, size_kind TEXT NOT NULL, unit TEXT NOT NULL,
    expected_qty INTEGER NOT NULL, counted_qty INTEGER, unit_price INTEGER,
    PRIMARY KEY(count_id, variant_id)
);
CREATE TABLE IF NOT EXISTS stock_count_documents (
    count_id INTEGER NOT NULL REFERENCES stock_counts(id), document_id INTEGER NOT NULL REFERENCES documents(id),
    PRIMARY KEY(count_id, document_id)
);
