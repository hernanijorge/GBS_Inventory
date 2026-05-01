-- ============================================================================
-- GBS (Global Business Solution) - Boston
-- Sistema de Controle de Estoque de Equipamentos
-- Script de criação das tabelas Oracle
-- ============================================================================

-- ----------------------------------------------------------------------------
-- TBL_EQUIPAMENTO - Catálogo principal de equipamentos
-- ----------------------------------------------------------------------------
CREATE TABLE TBL_EQUIPAMENTO (
    ID_EQUIPAMENTO       NUMBER(10)        NOT NULL,
    INTERNAL_UID         VARCHAR2(20)      NOT NULL,
    MARCA                VARCHAR2(50)      NOT NULL,
    MODELO               VARCHAR2(100)     NOT NULL,
    SERIAL_NUMBER        VARCHAR2(60)      NOT NULL,
    CPU_FAMILY           VARCHAR2(50),
    CPU_MODEL            VARCHAR2(30),
    CPU_SPEED_GHZ        NUMBER(4,2),
    RAM_GB               VARCHAR2(20),
    STORAGE_GB           VARCHAR2(40),
    HARD_DRIVE_TYPE      VARCHAR2(80),
    RESOLUTION           VARCHAR2(20),
    GRAPHICS             VARCHAR2(80),
    DEVICE_TYPE          VARCHAR2(20)      DEFAULT 'LAPTOP',
    CONDITION_STATUS     VARCHAR2(20)      DEFAULT 'GOOD',
    NOTES                VARCHAR2(500),
    SOURCE_BATCH         VARCHAR2(50),
    STATUS               VARCHAR2(20)      DEFAULT 'IN_STOCK',
    ID_EMPRESA           NUMBER(10),
    DATA_CADASTRO        DATE              DEFAULT SYSDATE,
    DATA_ALTERACAO       DATE,
    ID_USUARIO_CADASTRO  NUMBER(10),
    CONSTRAINT PK_EQUIPAMENTO   PRIMARY KEY (ID_EQUIPAMENTO),
    CONSTRAINT UK_EQUIP_UID     UNIQUE (INTERNAL_UID),
    CONSTRAINT UK_EQUIP_SERIAL  UNIQUE (SERIAL_NUMBER),
    CONSTRAINT CK_DEVICE_TYPE   CHECK (DEVICE_TYPE IN ('LAPTOP','DESKTOP','TABLET','OTHER')),
    CONSTRAINT CK_CONDITION     CHECK (CONDITION_STATUS IN ('EXCELLENT','GOOD','FAIR','POOR')),
    CONSTRAINT CK_STATUS        CHECK (STATUS IN ('IN_STOCK','LISTED','SOLD','SHIPPED','UPGRADING','BROKEN'))
);

CREATE INDEX IDX_EQUIP_MFR_MODEL ON TBL_EQUIPAMENTO (MARCA, MODELO);
CREATE INDEX IDX_EQUIP_STATUS    ON TBL_EQUIPAMENTO (STATUS);
CREATE INDEX IDX_EQUIP_BATCH     ON TBL_EQUIPAMENTO (SOURCE_BATCH);

-- ----------------------------------------------------------------------------
-- TBL_EQUIPAMENTO_UPGRADE - Histórico de upgrades (RAM, SSD, etc)
-- ----------------------------------------------------------------------------
CREATE TABLE TBL_EQUIPAMENTO_UPGRADE (
    ID_UPGRADE           NUMBER(10)        NOT NULL,
    ID_EQUIPAMENTO       NUMBER(10)        NOT NULL,
    INTERNAL_UID         VARCHAR2(20)      NOT NULL,
    COMPONENT_TYPE       VARCHAR2(20)      NOT NULL,
    VALUE_BEFORE         VARCHAR2(60),
    VALUE_AFTER          VARCHAR2(60)      NOT NULL,
    PART_SERIAL          VARCHAR2(80),
    SOURCE_ORIGEM        VARCHAR2(30)      DEFAULT 'NEW_PURCHASE',
    COST_USD             NUMBER(10,2),
    TECHNICIAN           VARCHAR2(80),
    NOTES                VARCHAR2(500),
    DATA_UPGRADE         DATE              DEFAULT SYSDATE,
    CONSTRAINT PK_UPGRADE      PRIMARY KEY (ID_UPGRADE),
    CONSTRAINT FK_UPGRADE_EQUIP FOREIGN KEY (ID_EQUIPAMENTO) REFERENCES TBL_EQUIPAMENTO (ID_EQUIPAMENTO),
    CONSTRAINT CK_COMPONENT    CHECK (COMPONENT_TYPE IN ('RAM','SSD','HDD','BATTERY','SCREEN','KEYBOARD','COVER','GPU','OTHER')),
    CONSTRAINT CK_SOURCE       CHECK (SOURCE_ORIGEM IN ('SHIPMENT_SURPLUS','NEW_PURCHASE','TRANSFERRED','WARRANTY','OTHER'))
);

CREATE INDEX IDX_UPGRADE_UID    ON TBL_EQUIPAMENTO_UPGRADE (INTERNAL_UID);
CREATE INDEX IDX_UPGRADE_DATA   ON TBL_EQUIPAMENTO_UPGRADE (DATA_UPGRADE);

-- ----------------------------------------------------------------------------
-- TBL_REMESSA - Remessas de envio (outbound) e recebimento (inbound)
-- ----------------------------------------------------------------------------
CREATE TABLE TBL_REMESSA (
    ID_REMESSA           NUMBER(10)        NOT NULL,
    REMESSA_REF          VARCHAR2(40)      NOT NULL,
    DIRECTION            VARCHAR2(10)      DEFAULT 'OUTBOUND',
    CARRIER              VARCHAR2(20)      NOT NULL,
    TRACKING_NUMBER      VARCHAR2(80),
    STATUS_REMESSA       VARCHAR2(30)      DEFAULT 'LABEL_CREATED',
    SENDER_NAME          VARCHAR2(120),
    SENDER_ADDRESS       VARCHAR2(300),
    RECIPIENT_NAME       VARCHAR2(120),
    RECIPIENT_ADDRESS    VARCHAR2(300),
    WEIGHT_LBS           NUMBER(6,2),
    SHIPPING_COST_USD    NUMBER(10,2),
    INSURANCE_USD        NUMBER(10,2),
    SERVICE_LEVEL        VARCHAR2(50),
    ESTIMATED_DELIVERY   DATE,
    ACTUAL_DELIVERY      DATE,
    NOTES                VARCHAR2(500),
    DATA_CADASTRO        DATE              DEFAULT SYSDATE,
    DATA_ALTERACAO       DATE,
    CONSTRAINT PK_REMESSA        PRIMARY KEY (ID_REMESSA),
    CONSTRAINT UK_REMESSA_REF    UNIQUE (REMESSA_REF),
    CONSTRAINT CK_DIRECTION      CHECK (DIRECTION IN ('OUTBOUND','INBOUND')),
    CONSTRAINT CK_CARRIER        CHECK (CARRIER IN ('FEDEX','UPS','USPS','DHL','OTHER')),
    CONSTRAINT CK_STATUS_REM     CHECK (STATUS_REMESSA IN ('LABEL_CREATED','PICKED_UP','IN_TRANSIT','OUT_FOR_DELIVERY','DELIVERED','DELAYED','EXCEPTION','RETURNED'))
);

CREATE INDEX IDX_REMESSA_TRACK   ON TBL_REMESSA (TRACKING_NUMBER);
CREATE INDEX IDX_REMESSA_STATUS  ON TBL_REMESSA (STATUS_REMESSA);

-- ----------------------------------------------------------------------------
-- TBL_REMESSA_ITEM - Itens (equipamentos) dentro de cada remessa
-- ----------------------------------------------------------------------------
CREATE TABLE TBL_REMESSA_ITEM (
    ID_REMESSA_ITEM      NUMBER(10)        NOT NULL,
    ID_REMESSA           NUMBER(10)        NOT NULL,
    ID_EQUIPAMENTO       NUMBER(10)        NOT NULL,
    INTERNAL_UID         VARCHAR2(20)      NOT NULL,
    CONDITION_AT_SHIP    VARCHAR2(20)      DEFAULT 'GOOD',
    SALE_PRICE_USD       NUMBER(10,2),
    NOTES                VARCHAR2(300),
    CONSTRAINT PK_REMESSA_ITEM   PRIMARY KEY (ID_REMESSA_ITEM),
    CONSTRAINT FK_REM_ITEM_REM   FOREIGN KEY (ID_REMESSA)     REFERENCES TBL_REMESSA (ID_REMESSA),
    CONSTRAINT FK_REM_ITEM_EQUIP FOREIGN KEY (ID_EQUIPAMENTO) REFERENCES TBL_EQUIPAMENTO (ID_EQUIPAMENTO),
    CONSTRAINT UK_REM_ITEM       UNIQUE (ID_REMESSA, ID_EQUIPAMENTO)
);

-- ----------------------------------------------------------------------------
-- TBL_IMPORTACAO_LOG - Registro de importações de planilhas
-- ----------------------------------------------------------------------------
CREATE TABLE TBL_IMPORTACAO_LOG (
    ID_IMPORTACAO        NUMBER(10)        NOT NULL,
    FILENAME             VARCHAR2(200)     NOT NULL,
    SHEETS_FOUND         NUMBER(4),
    ROWS_IMPORTED        NUMBER(6),
    ROWS_SKIPPED         NUMBER(6),
    ROWS_UPDATED         NUMBER(6),
    ERROS                CLOB,
    IMPORTED_BY          VARCHAR2(80),
    DATA_IMPORTACAO      DATE              DEFAULT SYSDATE,
    CONSTRAINT PK_IMPORTACAO     PRIMARY KEY (ID_IMPORTACAO)
);

-- ----------------------------------------------------------------------------
-- TBL_WHATSAPP_LOG - Log de mensagens WhatsApp
-- ----------------------------------------------------------------------------
CREATE TABLE TBL_WHATSAPP_LOG (
    ID_WHATSAPP          NUMBER(10)        NOT NULL,
    PHONE_NUMBER         VARCHAR2(30)      NOT NULL,
    COMMAND_RECEIVED     VARCHAR2(200),
    RESPONSE_SENT        CLOB,
    DATA_MENSAGEM        DATE              DEFAULT SYSDATE,
    CONSTRAINT PK_WHATSAPP       PRIMARY KEY (ID_WHATSAPP)
);

COMMIT;
