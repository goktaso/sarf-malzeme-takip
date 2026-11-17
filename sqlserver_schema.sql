IF OBJECT_ID('dbo.departments','U') IS NULL
BEGIN
    CREATE TABLE departments (
        id bigint IDENTITY PRIMARY KEY, name nvarchar(250) NOT NULL UNIQUE,
        active int NOT NULL DEFAULT 1 CHECK(active IN (0,1))
    );
    CREATE TABLE people (
        id bigint IDENTITY PRIMARY KEY, code nvarchar(250) NOT NULL UNIQUE, name nvarchar(250) NOT NULL,
        department_id bigint NOT NULL REFERENCES departments(id),
        active int NOT NULL DEFAULT 1 CHECK(active IN (0,1))
    );
    CREATE TABLE product_groups (
        id bigint IDENTITY PRIMARY KEY, name nvarchar(250) NOT NULL, prefix varchar(8) NOT NULL UNIQUE,
        module varchar(4) NOT NULL CHECK(module IN ('sarf','kkd')), next_number bigint NOT NULL DEFAULT 1,
        size_kind varchar(8) NOT NULL DEFAULT 'size' CHECK(size_kind IN ('size','shoe'))
    );
    CREATE TABLE products (
        id bigint IDENTITY PRIMARY KEY, code nvarchar(250) NOT NULL UNIQUE, name nvarchar(250) NOT NULL,
        module varchar(4) NOT NULL CHECK(module IN ('sarf','kkd')), category nvarchar(250) NOT NULL,
        unit nvarchar(250) NOT NULL, scope varchar(10) NOT NULL CHECK(scope IN ('person','department','both')),
        group_id bigint NULL REFERENCES product_groups(id), active int NOT NULL DEFAULT 1 CHECK(active IN (0,1))
    );
    CREATE TABLE variants (
        id bigint IDENTITY PRIMARY KEY, product_id bigint NOT NULL REFERENCES products(id),
        size nvarchar(250) NOT NULL, color nvarchar(250) NOT NULL DEFAULT N'', minimum bigint NOT NULL,
        UNIQUE(product_id,size,color)
    );
    CREATE TABLE documents (
        id bigint IDENTITY PRIMARY KEY, request_id nvarchar(250) NOT NULL UNIQUE, kind varchar(12) NOT NULL,
        day varchar(10) NOT NULL, purchase_day varchar(10) NULL, supplier nvarchar(250), reference nvarchar(250),
        note nvarchar(250), department_id bigint REFERENCES departments(id), person_id bigint REFERENCES people(id),
        source_id bigint REFERENCES documents(id)
    );
    CREATE TABLE movements (
        id bigint IDENTITY PRIMARY KEY, document_id bigint NOT NULL REFERENCES documents(id),
        variant_id bigint NOT NULL REFERENCES variants(id), qty bigint NOT NULL, value bigint NOT NULL,
        unit_price bigint, UNIQUE(document_id,variant_id)
    );
    CREATE INDEX movement_variant ON movements(variant_id);
    CREATE TABLE definition_versions (
        id bigint IDENTITY PRIMARY KEY,
        entity varchar(20) NOT NULL CHECK(entity IN ('products','departments','people','product_groups')),
        entity_id bigint NOT NULL, effective_day varchar(10) NOT NULL,
        recorded_at varchar(30) NOT NULL DEFAULT (CONVERT(varchar(27),SYSUTCDATETIME(),126)+'Z'),
        state nvarchar(max) NOT NULL CHECK(ISJSON(state)=1)
    );
    CREATE INDEX definition_as_of ON definition_versions(entity,entity_id,effective_day DESC,id DESC);
END;
IF OBJECT_ID('dbo.app_users','U') IS NULL
BEGIN
    CREATE TABLE app_users (
        sid varchar(184) PRIMARY KEY, account nvarchar(250) NOT NULL, display_name nvarchar(250) NOT NULL,
        is_admin int NOT NULL CHECK(is_admin IN (0,1)), active int NOT NULL CHECK(active IN (0,1)),
        permissions nvarchar(max) NOT NULL CHECK(ISJSON(permissions)=1),
        department_id bigint NULL REFERENCES departments(id)
    );
    CREATE TABLE app_sessions (
        token_hash char(64) PRIMARY KEY, sid varchar(184) NOT NULL REFERENCES app_users(sid),
        expires_epoch bigint NOT NULL
    );
    CREATE INDEX app_sessions_expiry ON app_sessions(expires_epoch);
    CREATE TABLE app_audit (
        id bigint IDENTITY PRIMARY KEY, sid varchar(184) NOT NULL REFERENCES app_users(sid),
        action varchar(50) NOT NULL, record_ref nvarchar(250),
        recorded_at varchar(30) NOT NULL DEFAULT (CONVERT(varchar(27),SYSUTCDATETIME(),126)+'Z')
    );
END;
IF OBJECT_ID('dbo.app_period_grants','U') IS NULL
BEGIN
    CREATE TABLE app_period_grants (
        id bigint IDENTITY PRIMARY KEY,
        sid varchar(184) NOT NULL REFERENCES app_users(sid),
        start_day varchar(10) NOT NULL,
        end_day varchar(10) NOT NULL,
        reason nvarchar(250) NOT NULL,
        granted_by varchar(184) NOT NULL REFERENCES app_users(sid),
        granted_at varchar(30) NOT NULL,
        revoked_by varchar(184) NULL REFERENCES app_users(sid),
        revoked_at varchar(30) NULL,
        CHECK(start_day<=end_day)
    );
    CREATE INDEX period_grant_user_day ON app_period_grants(sid,start_day,end_day) INCLUDE (revoked_at);
END;
IF OBJECT_ID('dbo.stock_counts','U') IS NULL
BEGIN
    CREATE TABLE stock_counts (
        id bigint IDENTITY PRIMARY KEY, day varchar(10) NOT NULL, note nvarchar(250) NOT NULL,
        status varchar(10) NOT NULL DEFAULT 'draft' CHECK(status IN ('draft','locked','rejected','approved')),
        revision bigint NOT NULL DEFAULT 0, owner nvarchar(250) NOT NULL
    );
    CREATE TABLE stock_count_lines (
        count_id bigint NOT NULL REFERENCES stock_counts(id), variant_id bigint NOT NULL REFERENCES variants(id),
        code nvarchar(250) NOT NULL, name nvarchar(250) NOT NULL, category nvarchar(250) NOT NULL,
        color nvarchar(250) NOT NULL, size nvarchar(250) NOT NULL, size_kind varchar(8) NOT NULL,
        unit nvarchar(250) NOT NULL, expected_qty bigint NOT NULL, counted_qty bigint, unit_price bigint,
        PRIMARY KEY(count_id,variant_id)
    );
    CREATE TABLE stock_count_documents (
        count_id bigint NOT NULL REFERENCES stock_counts(id), document_id bigint NOT NULL REFERENCES documents(id),
        PRIMARY KEY(count_id,document_id)
    );
END;
