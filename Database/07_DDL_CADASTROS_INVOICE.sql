-- ============================================================================
-- GBS Inventory - Etapa 1
-- Cadastros e controle de invoice (incremental, sem alterar tabelas atuais)
-- ============================================================================

-- ----------------------------------------------------------------------------
-- TBL_CLIENTE
-- ----------------------------------------------------------------------------
CREATE TABLE TBL_CLIENTE (
    ID_CLIENTE          NUMBER(10)        NOT NULL,
    NOME_RAZAO          VARCHAR2(150)     NOT NULL,
    NOME_FANTASIA       VARCHAR2(150),
    DOCUMENTO           VARCHAR2(40),
    EMAIL               VARCHAR2(120),
    TELEFONE            VARCHAR2(40),
    ENDERECO1           VARCHAR2(200),
    ENDERECO2           VARCHAR2(200),
    CIDADE              VARCHAR2(80),
    ESTADO              VARCHAR2(40),
    ZIP_CODE            VARCHAR2(20),
    PAIS                VARCHAR2(40)      DEFAULT 'USA',
    ATIVO               CHAR(1)           DEFAULT 'Y' NOT NULL,
    OBSERVACOES         VARCHAR2(500),
    DATA_CADASTRO       DATE              DEFAULT SYSDATE,
    DATA_ALTERACAO      DATE,
    CONSTRAINT PK_CLIENTE          PRIMARY KEY (ID_CLIENTE),
    CONSTRAINT CK_CLIENTE_ATIVO    CHECK (ATIVO IN ('Y','N'))
);

CREATE INDEX IDX_CLIENTE_NOME     ON TBL_CLIENTE (NOME_RAZAO);
CREATE INDEX IDX_CLIENTE_EMAIL    ON TBL_CLIENTE (EMAIL);

-- ----------------------------------------------------------------------------
-- TBL_INVOICE
-- ----------------------------------------------------------------------------
CREATE TABLE TBL_INVOICE (
    ID_INVOICE          NUMBER(10)        NOT NULL,
    INVOICE_NUMBER      VARCHAR2(40)      NOT NULL,
    ID_CLIENTE          NUMBER(10)        NOT NULL,
    ID_REMESSA          NUMBER(10),
    ISSUE_DATE          DATE              DEFAULT TRUNC(SYSDATE) NOT NULL,
    DUE_DATE            DATE,
    CURRENCY            VARCHAR2(3)       DEFAULT 'USD' NOT NULL,
    SUBTOTAL_USD        NUMBER(12,2)      DEFAULT 0 NOT NULL,
    DISCOUNT_USD        NUMBER(12,2)      DEFAULT 0 NOT NULL,
    SHIPPING_USD        NUMBER(12,2)      DEFAULT 0 NOT NULL,
    TAX_USD             NUMBER(12,2)      DEFAULT 0 NOT NULL,
    TOTAL_USD           NUMBER(12,2)      DEFAULT 0 NOT NULL,
    STATUS_INVOICE      VARCHAR2(20)      DEFAULT 'DRAFT' NOT NULL,
    CLIENT_EMAIL        VARCHAR2(120),
    NOTES               VARCHAR2(1000),
    LOGO_PATH           VARCHAR2(500),
    WORD_FILE_PATH      VARCHAR2(500),
    PDF_FILE_PATH       VARCHAR2(500),
    DATA_CADASTRO       DATE              DEFAULT SYSDATE,
    DATA_ALTERACAO      DATE,
    CONSTRAINT PK_INVOICE            PRIMARY KEY (ID_INVOICE),
    CONSTRAINT UK_INVOICE_NUMBER     UNIQUE (INVOICE_NUMBER),
    CONSTRAINT FK_INVOICE_CLIENTE    FOREIGN KEY (ID_CLIENTE) REFERENCES TBL_CLIENTE (ID_CLIENTE),
    CONSTRAINT FK_INVOICE_REMESSA    FOREIGN KEY (ID_REMESSA) REFERENCES TBL_REMESSA (ID_REMESSA),
    CONSTRAINT CK_INVOICE_STATUS     CHECK (STATUS_INVOICE IN ('DRAFT','SENT','PAID','CANCELLED'))
);

CREATE INDEX IDX_INVOICE_CLIENTE   ON TBL_INVOICE (ID_CLIENTE);
CREATE INDEX IDX_INVOICE_REMESSA   ON TBL_INVOICE (ID_REMESSA);
CREATE INDEX IDX_INVOICE_STATUS    ON TBL_INVOICE (STATUS_INVOICE);

-- ----------------------------------------------------------------------------
-- TBL_INVOICE_ITEM
-- ----------------------------------------------------------------------------
CREATE TABLE TBL_INVOICE_ITEM (
    ID_INVOICE_ITEM     NUMBER(10)        NOT NULL,
    ID_INVOICE          NUMBER(10)        NOT NULL,
    ID_EQUIPAMENTO      NUMBER(10),
    DESCRIPTION         VARCHAR2(200)     NOT NULL,
    QTY                 NUMBER(10,2)      DEFAULT 1 NOT NULL,
    UNIT_PRICE_USD      NUMBER(12,2)      DEFAULT 0 NOT NULL,
    LINE_TOTAL_USD      NUMBER(12,2)      DEFAULT 0 NOT NULL,
    NOTES               VARCHAR2(400),
    CONSTRAINT PK_INVOICE_ITEM           PRIMARY KEY (ID_INVOICE_ITEM),
    CONSTRAINT FK_INV_ITEM_INVOICE       FOREIGN KEY (ID_INVOICE) REFERENCES TBL_INVOICE (ID_INVOICE),
    CONSTRAINT FK_INV_ITEM_EQUIP         FOREIGN KEY (ID_EQUIPAMENTO) REFERENCES TBL_EQUIPAMENTO (ID_EQUIPAMENTO)
);

CREATE INDEX IDX_INV_ITEM_INVOICE  ON TBL_INVOICE_ITEM (ID_INVOICE);
CREATE INDEX IDX_INV_ITEM_EQUIP    ON TBL_INVOICE_ITEM (ID_EQUIPAMENTO);

-- ----------------------------------------------------------------------------
-- Sequences
-- ----------------------------------------------------------------------------
CREATE SEQUENCE SEQ_CLIENTE
    START WITH 1
    INCREMENT BY 1
    NOCACHE
    NOCYCLE;

CREATE SEQUENCE SEQ_INVOICE
    START WITH 1
    INCREMENT BY 1
    NOCACHE
    NOCYCLE;

CREATE SEQUENCE SEQ_INVOICE_ITEM
    START WITH 1
    INCREMENT BY 1
    NOCACHE
    NOCYCLE;

COMMIT;
