-- internal_edo_api: договоры аренды/субаренды и пакетная загрузка из Excel.
-- Общие таблицы documents / participants не меняются.
-- Повторный запуск безопасен (IF NOT EXISTS).

BEGIN;

CREATE TABLE IF NOT EXISTS lease_contracts (
    id              SERIAL PRIMARY KEY,
    category        VARCHAR(16)    NOT NULL,          -- LEASE / SUBLEASE
    contract_number VARCHAR        NOT NULL,
    signed_date     DATE           NOT NULL,
    start_date      DATE           NOT NULL,
    end_date        DATE           NOT NULL,

    seller_tin      VARCHAR(12)    NOT NULL,          -- арендодатель (participants SELLER)
    seller_name     VARCHAR        NOT NULL,
    buyer_tin       VARCHAR(12)    NOT NULL,          -- арендатор (participants BUYER)
    buyer_name      VARCHAR        NOT NULL,
    buyer_phone     VARCHAR,

    floor           VARCHAR,
    room_type       VARCHAR,
    area            NUMERIC(12, 2) NOT NULL,
    rent_per_sqm    NUMERIC(14, 2) NOT NULL,
    rent_total      NUMERIC(14, 2) NOT NULL,
    communal        NUMERIC(14, 2) NOT NULL,

    data            JSONB          NOT NULL,          -- полные данные для шаблона
    docx_path       VARCHAR,                          -- исходный .docx в MinIO (не подписывается)
    pdf_id          UUID REFERENCES documents (id),
    xml_id          UUID REFERENCES documents (id),
    created         TIMESTAMP      NOT NULL DEFAULT now(),

    CONSTRAINT uq_lease_contracts_number UNIQUE (category, seller_tin, contract_number)
);

CREATE INDEX IF NOT EXISTS ix_lease_contracts_buyer_tin ON lease_contracts (buyer_tin);
CREATE INDEX IF NOT EXISTS ix_lease_contracts_pdf_id ON lease_contracts (pdf_id);


CREATE TABLE IF NOT EXISTS lease_batches (
    id          UUID PRIMARY KEY,
    category    VARCHAR(16) NOT NULL,
    filename    VARCHAR,
    source_path VARCHAR,                              -- загруженный Excel в MinIO
    total       INTEGER     NOT NULL DEFAULT 0,
    created_by  VARCHAR,
    created     TIMESTAMP   NOT NULL DEFAULT now()
);

CREATE TABLE IF NOT EXISTS lease_batch_items (
    id                SERIAL PRIMARY KEY,
    batch_id          UUID        NOT NULL REFERENCES lease_batches (id) ON DELETE CASCADE,
    row_number        INTEGER     NOT NULL,           -- строка в исходном Excel
    payload           JSONB,                          -- данные договора, если строка прошла проверку
    status            VARCHAR(16) NOT NULL,           -- rejected / pending / processing / done / error
    error             TEXT,
    attempts          INTEGER     NOT NULL DEFAULT 0,
    lease_contract_id INTEGER REFERENCES lease_contracts (id),
    created           TIMESTAMP   NOT NULL DEFAULT now(),
    updated           TIMESTAMP   NOT NULL DEFAULT now()
);

CREATE INDEX IF NOT EXISTS ix_lease_batch_items_batch_id ON lease_batch_items (batch_id);
-- Очередь: воркер выбирает только pending/processing
CREATE INDEX IF NOT EXISTS ix_lease_batch_items_queue
    ON lease_batch_items (status, id) WHERE status IN ('pending', 'processing');

COMMIT;
