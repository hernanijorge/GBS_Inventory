-- ============================================================================
-- PACK_REMESSA - Specification
-- Controle de remessas FedEx, UPS, USPS (outbound e inbound)
-- ============================================================================

CREATE OR REPLACE PACKAGE PACK_REMESSA IS

    FUNCTION FUNC_PROXIMO_REMESSA
        RETURN NUMBER;

    FUNCTION FUNC_PROXIMO_REMESSA_REF
        RETURN VARCHAR2;

    FUNCTION FUNC_STATUS_REMESSA_DESC (V_STATUS IN VARCHAR2)
        RETURN VARCHAR2;

    PROCEDURE PROC_INSERT_REMESSA (
        V_DIRECTION          VARCHAR2,
        V_CARRIER            VARCHAR2,
        V_TRACKING_NUMBER    VARCHAR2,
        V_SENDER_NAME        VARCHAR2,
        V_SENDER_ADDRESS     VARCHAR2,
        V_RECIPIENT_NAME     VARCHAR2,
        V_RECIPIENT_ADDRESS  VARCHAR2,
        V_WEIGHT_LBS         NUMBER,
        V_SHIPPING_COST_USD  NUMBER,
        V_INSURANCE_USD      NUMBER,
        V_SERVICE_LEVEL      VARCHAR2,
        V_ESTIMATED_DELIVERY DATE,
        V_NOTES              VARCHAR2,
        V_ID                 OUT NUMBER,
        V_REF                OUT VARCHAR2);

    PROCEDURE PROC_UPDATE_STATUS_REMESSA (
        V_ID            NUMBER,
        V_STATUS        VARCHAR2,
        V_DATA_ENTREGA  DATE);

    PROCEDURE PROC_INSERT_REMESSA_ITEM (
        V_ID_REMESSA       NUMBER,
        V_ID_EQUIPAMENTO   NUMBER,
        V_INTERNAL_UID     VARCHAR2,
        V_CONDITION        VARCHAR2,
        V_SALE_PRICE_USD   NUMBER,
        V_NOTES            VARCHAR2);

    PROCEDURE PROC_DELETE_REMESSA (V_ID IN NUMBER);

    PROCEDURE PROC_SELECT_REMESSAS_ATIVAS (
        V_CURSOR OUT SYS_REFCURSOR);

    PROCEDURE PROC_SELECT_REMESSA (
        V_ID     IN  NUMBER,
        V_CURSOR OUT SYS_REFCURSOR);

    PROCEDURE PROC_SELECT_ITENS_REMESSA (
        V_ID_REMESSA IN  NUMBER,
        V_CURSOR     OUT SYS_REFCURSOR);

END PACK_REMESSA;
/

-- ============================================================================
-- PACK_REMESSA - Body
-- ============================================================================

CREATE OR REPLACE PACKAGE BODY PACK_REMESSA IS

-- ----------------------------------------------------------------------------
FUNCTION FUNC_PROXIMO_REMESSA
RETURN NUMBER

AS

V_ID NUMBER;

BEGIN

       SELECT SEQ_REMESSA.NEXTVAL INTO V_ID FROM DUAL;

       RETURN V_ID;

END;

-- ----------------------------------------------------------------------------
FUNCTION FUNC_PROXIMO_REMESSA_REF
RETURN VARCHAR2

AS

V_REF VARCHAR2(40);

BEGIN

       V_REF := 'GBS-SH-' || TO_CHAR(SYSDATE, 'YYYYMMDDHH24MISS');

       RETURN V_REF;

END;

-- ----------------------------------------------------------------------------
FUNCTION FUNC_STATUS_REMESSA_DESC (V_STATUS IN VARCHAR2)
RETURN VARCHAR2

AS

BEGIN

       RETURN CASE V_STATUS
                WHEN 'LABEL_CREATED'    THEN 'Etiqueta Criada'
                WHEN 'PICKED_UP'        THEN 'Coletado'
                WHEN 'IN_TRANSIT'       THEN 'Em Trânsito'
                WHEN 'OUT_FOR_DELIVERY' THEN 'Saiu para Entrega'
                WHEN 'DELIVERED'        THEN 'Entregue'
                WHEN 'DELAYED'          THEN 'Atrasado'
                WHEN 'EXCEPTION'        THEN 'Exceção'
                WHEN 'RETURNED'         THEN 'Devolvido'
                ELSE 'Desconhecido'
              END;

END;

-- ----------------------------------------------------------------------------
PROCEDURE PROC_INSERT_REMESSA (
    V_DIRECTION          VARCHAR2,
    V_CARRIER            VARCHAR2,
    V_TRACKING_NUMBER    VARCHAR2,
    V_SENDER_NAME        VARCHAR2,
    V_SENDER_ADDRESS     VARCHAR2,
    V_RECIPIENT_NAME     VARCHAR2,
    V_RECIPIENT_ADDRESS  VARCHAR2,
    V_WEIGHT_LBS         NUMBER,
    V_SHIPPING_COST_USD  NUMBER,
    V_INSURANCE_USD      NUMBER,
    V_SERVICE_LEVEL      VARCHAR2,
    V_ESTIMATED_DELIVERY DATE,
    V_NOTES              VARCHAR2,
    V_ID                 OUT NUMBER,
    V_REF                OUT VARCHAR2)

AS

V_IDREMESSA NUMBER;
V_REFREMESSA VARCHAR2(40);

BEGIN
         V_IDREMESSA := FUNC_PROXIMO_REMESSA;
         V_REFREMESSA := FUNC_PROXIMO_REMESSA_REF;

         V_ID  := V_IDREMESSA;
         V_REF := V_REFREMESSA;

         INSERT INTO TBL_REMESSA
                      (ID_REMESSA,
                       REMESSA_REF,
                       DIRECTION,
                       CARRIER,
                       TRACKING_NUMBER,
                       STATUS_REMESSA,
                       SENDER_NAME,
                       SENDER_ADDRESS,
                       RECIPIENT_NAME,
                       RECIPIENT_ADDRESS,
                       WEIGHT_LBS,
                       SHIPPING_COST_USD,
                       INSURANCE_USD,
                       SERVICE_LEVEL,
                       ESTIMATED_DELIVERY,
                       NOTES,
                       DATA_CADASTRO)
               VALUES
                      (V_IDREMESSA,
                       V_REFREMESSA,
                       NVL(UPPER(V_DIRECTION), 'OUTBOUND'),
                       UPPER(V_CARRIER),
                       V_TRACKING_NUMBER,
                       'LABEL_CREATED',
                       V_SENDER_NAME,
                       V_SENDER_ADDRESS,
                       V_RECIPIENT_NAME,
                       V_RECIPIENT_ADDRESS,
                       V_WEIGHT_LBS,
                       V_SHIPPING_COST_USD,
                       V_INSURANCE_USD,
                       V_SERVICE_LEVEL,
                       V_ESTIMATED_DELIVERY,
                       V_NOTES,
                       SYSDATE);

END;

-- ----------------------------------------------------------------------------
PROCEDURE PROC_UPDATE_STATUS_REMESSA (
    V_ID            NUMBER,
    V_STATUS        VARCHAR2,
    V_DATA_ENTREGA  DATE)

AS

BEGIN

       UPDATE TBL_REMESSA SET
                STATUS_REMESSA  = UPPER(V_STATUS),
                ACTUAL_DELIVERY = V_DATA_ENTREGA,
                DATA_ALTERACAO  = SYSDATE
        WHERE   ID_REMESSA = V_ID;

       -- Se entregue, marca equipamentos como SOLD
       IF UPPER(V_STATUS) = 'DELIVERED' THEN
          UPDATE TBL_EQUIPAMENTO SET
                   STATUS = 'SOLD',
                   DATA_ALTERACAO = SYSDATE
           WHERE   ID_EQUIPAMENTO IN (SELECT ID_EQUIPAMENTO
                                      FROM   TBL_REMESSA_ITEM
                                      WHERE  ID_REMESSA = V_ID);
       END IF;

END;

-- ----------------------------------------------------------------------------
PROCEDURE PROC_INSERT_REMESSA_ITEM (
    V_ID_REMESSA       NUMBER,
    V_ID_EQUIPAMENTO   NUMBER,
    V_INTERNAL_UID     VARCHAR2,
    V_CONDITION        VARCHAR2,
    V_SALE_PRICE_USD   NUMBER,
    V_NOTES            VARCHAR2)

AS

BEGIN

       INSERT INTO TBL_REMESSA_ITEM
                   (ID_REMESSA_ITEM,
                    ID_REMESSA,
                    ID_EQUIPAMENTO,
                    INTERNAL_UID,
                    CONDITION_AT_SHIP,
                    SALE_PRICE_USD,
                    NOTES)
            VALUES
                   (SEQ_REMESSA_ITEM.NEXTVAL,
                    V_ID_REMESSA,
                    V_ID_EQUIPAMENTO,
                    V_INTERNAL_UID,
                    NVL(UPPER(V_CONDITION), 'GOOD'),
                    V_SALE_PRICE_USD,
                    V_NOTES);

       -- Marca equipamento como SHIPPED
       UPDATE TBL_EQUIPAMENTO SET
                STATUS         = 'SHIPPED',
                DATA_ALTERACAO = SYSDATE
        WHERE   ID_EQUIPAMENTO = V_ID_EQUIPAMENTO;

END;

-- ----------------------------------------------------------------------------
PROCEDURE PROC_DELETE_REMESSA (V_ID IN NUMBER)

AS

BEGIN

       DELETE FROM TBL_REMESSA_ITEM WHERE ID_REMESSA = V_ID;
       DELETE FROM TBL_REMESSA      WHERE ID_REMESSA = V_ID;

END;

-- ----------------------------------------------------------------------------
PROCEDURE PROC_SELECT_REMESSAS_ATIVAS (V_CURSOR OUT SYS_REFCURSOR)

AS

BEGIN

        OPEN V_CURSOR FOR
             SELECT R.ID_REMESSA,
                    R.REMESSA_REF,
                    R.DIRECTION,
                    R.CARRIER,
                    R.TRACKING_NUMBER,
                    R.STATUS_REMESSA,
                    FUNC_STATUS_REMESSA_DESC(R.STATUS_REMESSA) AS STATUS_DESCRICAO,
                    R.RECIPIENT_NAME,
                    R.WEIGHT_LBS,
                    R.SHIPPING_COST_USD,
                    R.SERVICE_LEVEL,
                    R.ESTIMATED_DELIVERY,
                    R.ACTUAL_DELIVERY,
                    R.DATA_CADASTRO,
                    (SELECT COUNT(*) FROM TBL_REMESSA_ITEM
                     WHERE ID_REMESSA = R.ID_REMESSA) AS TOTAL_ITENS
             FROM   TBL_REMESSA R
             WHERE  R.STATUS_REMESSA NOT IN ('DELIVERED','RETURNED')
             ORDER  BY R.DATA_CADASTRO DESC;

END;

-- ----------------------------------------------------------------------------
PROCEDURE PROC_SELECT_REMESSA (
    V_ID     IN  NUMBER,
    V_CURSOR OUT SYS_REFCURSOR)

AS

BEGIN

        OPEN V_CURSOR FOR
             SELECT R.*,
                    FUNC_STATUS_REMESSA_DESC(R.STATUS_REMESSA) AS STATUS_DESCRICAO
             FROM   TBL_REMESSA R
             WHERE  R.ID_REMESSA = V_ID;

END;

-- ----------------------------------------------------------------------------
PROCEDURE PROC_SELECT_ITENS_REMESSA (
    V_ID_REMESSA IN  NUMBER,
    V_CURSOR     OUT SYS_REFCURSOR)

AS

BEGIN

        OPEN V_CURSOR FOR
             SELECT RI.ID_REMESSA_ITEM,
                    RI.INTERNAL_UID,
                    E.MANUFACTURER,
                    E.MODEL,
                    E.CPU_MODEL,
                    E.RAM_GB,
                    E.STORAGE_GB,
                    RI.CONDITION_AT_SHIP,
                    RI.SALE_PRICE_USD,
                    RI.NOTES
             FROM   TBL_REMESSA_ITEM RI,
                    TBL_EQUIPAMENTO  E
             WHERE  RI.ID_EQUIPAMENTO = E.ID_EQUIPAMENTO
             AND    RI.ID_REMESSA = V_ID_REMESSA
             ORDER  BY RI.INTERNAL_UID;

END;

END PACK_REMESSA;
/
