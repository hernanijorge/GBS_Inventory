Imports Oracle.ManagedDataAccess.Client
Imports System.Configuration
Imports System.Data
Imports GBS_Inventory.OracleHelper

Public Class clsLeituraHistorico

#Region "Atributos"

    Private ConnectionString As String

#End Region

#Region "Construtor"

    Public Sub New()
        Me.ConnectionString = ConfigurationManager.ConnectionStrings("OracleDB").ConnectionString
    End Sub

#End Region

#Region "Consultas"

    ''' <summary>
    ''' Dados do equipamento para o cabeçalho e aba Arrival.
    ''' Colunas confirmadas pelo código existente: MARCA, MODEL, PROCESSADOR,
    ''' RAM_GB, STORAGE_GB, STATUS, SERIAL_NUMBER, INTERNAL_UID, DATA_CADASTRO.
    ''' SOURCE_BATCH / CONDITION_STATUS / OBSERVACAO são opcionais — fallback usa CAST(NULL).
    ''' </summary>
    Public Function selecionarEquipamento(pIdEquipamento As Integer) As DataSet

        ' Tentativa 1 — todas as colunas incluindo as opcionais
        Dim sqlFull As String =
            "SELECT E.ID_EQUIPAMENTO, E.INTERNAL_UID, E.SERIAL_NUMBER," &
            "       E.MARCA, E.MODEL, E.PROCESSADOR," &
            "       E.RAM_GB, E.STORAGE_GB, E.STATUS," &
            "       E.DATA_CADASTRO, E.SOURCE_BATCH, E.CONDITION_STATUS," &
            "       E.OBSERVACAO" &
            "  FROM TBL_EQUIPAMENTO E" &
            " WHERE E.ID_EQUIPAMENTO = :P_ID"

        ' Tentativa 2 — sem OBSERVACAO (caso a coluna não exista)
        Dim sqlSemObs As String =
            "SELECT E.ID_EQUIPAMENTO, E.INTERNAL_UID, E.SERIAL_NUMBER," &
            "       E.MARCA, E.MODEL, E.PROCESSADOR," &
            "       E.RAM_GB, E.STORAGE_GB, E.STATUS," &
            "       E.DATA_CADASTRO, E.SOURCE_BATCH, E.CONDITION_STATUS," &
            "       CAST(NULL AS VARCHAR2(4000)) AS OBSERVACAO" &
            "  FROM TBL_EQUIPAMENTO E" &
            " WHERE E.ID_EQUIPAMENTO = :P_ID"

        ' Tentativa 3 — mínimo: só colunas confirmadas pelo código existente
        Dim sqlBase As String =
            "SELECT E.ID_EQUIPAMENTO, E.INTERNAL_UID, E.SERIAL_NUMBER," &
            "       E.MARCA, E.MODEL, E.PROCESSADOR," &
            "       E.RAM_GB, E.STORAGE_GB, E.STATUS," &
            "       E.DATA_CADASTRO," &
            "       CAST(NULL AS VARCHAR2(200))  AS SOURCE_BATCH," &
            "       CAST(NULL AS VARCHAR2(100))  AS CONDITION_STATUS," &
            "       CAST(NULL AS VARCHAR2(4000)) AS OBSERVACAO" &
            "  FROM TBL_EQUIPAMENTO E" &
            " WHERE E.ID_EQUIPAMENTO = :P_ID"

        Try
            Return OracleHelper.ExecuteDataset(Me.ConnectionString, CommandType.Text, sqlFull,
                                               New OracleParameter() {CriarParId(pIdEquipamento)})
        Catch ex1 As Exception
            If Not IsErroColuna(ex1) Then Throw New Exception(ex1.ToString)
        End Try

        Try
            Return OracleHelper.ExecuteDataset(Me.ConnectionString, CommandType.Text, sqlSemObs,
                                               New OracleParameter() {CriarParId(pIdEquipamento)})
        Catch ex2 As Exception
            If Not IsErroColuna(ex2) Then Throw New Exception(ex2.ToString)
        End Try

        Return OracleHelper.ExecuteDataset(Me.ConnectionString, CommandType.Text, sqlBase,
                                           New OracleParameter() {CriarParId(pIdEquipamento)})

    End Function

    ''' <summary>
    ''' Upgrades do equipamento ordenados por data.
    ''' Nomes reais deduzidos dos parâmetros do PROC_INSERT_UPGRADE:
    '''   P_TIPO_UPGRADE   → TIPO_UPGRADE   (não COMPONENT_TYPE)
    '''   P_TECNICO        → TECNICO         (não TECHNICIAN)
    '''   P_OBSERVACAO     → OBSERVACAO      (não NOTES)
    '''   P_RAM_ANTERIOR_GB / P_RAM_NOVA_GB / P_STORAGE_* → colunas RAM/STORAGE separadas
    '''   PART_SERIAL / SOURCE_ORIGEM / COST_USD → concatenados em OBSERVACAO pelo proc,
    '''     portanto NÃO são colunas físicas separadas → CAST(NULL).
    ''' </summary>
    Public Function selecionarUpgradesEquipamento(pIdEquipamento As Integer) As DataSet

        ' Query com nomes reais confirmados via parâmetros do proc de gravação.
        ' VALUE_BEFORE/AFTER exibem RAM e STORAGE concatenados para leitura humana.
        ' COST_USD / PART_SERIAL / SOURCE_ORIGEM são packed em OBSERVACAO → CAST(NULL).
        Dim sql As String =
            "SELECT U.ID_UPGRADE," &
            "       U.DATA_UPGRADE," &
            "       U.TIPO_UPGRADE                          AS COMPONENT_TYPE," &
            "       CAST(U.RAM_ANTERIOR_GB AS VARCHAR2(50)) AS VALUE_BEFORE," &
            "       CAST(U.RAM_NOVA_GB     AS VARCHAR2(50)) AS VALUE_AFTER," &
            "       CAST(NULL AS VARCHAR2(200))             AS SOURCE_ORIGEM," &
            "       CAST(NULL AS NUMBER(12,2))              AS COST_USD," &
            "       CAST(NULL AS VARCHAR2(200))             AS PART_SERIAL," &
            "       U.TECNICO                               AS TECHNICIAN," &
            "       U.OBSERVACAO                            AS NOTES" &
            "  FROM TBL_EQUIPAMENTO_UPGRADE U" &
            " WHERE U.ID_EQUIPAMENTO = :P_ID" &
            " ORDER BY U.DATA_UPGRADE DESC"

        ' Fallback: tenta com STORAGE em vez de RAM (para upgrades de SSD/HDD)
        Dim sqlStorage As String =
            "SELECT U.ID_UPGRADE," &
            "       U.DATA_UPGRADE," &
            "       U.TIPO_UPGRADE                              AS COMPONENT_TYPE," &
            "       CAST(U.STORAGE_ANTERIOR_GB AS VARCHAR2(50)) AS VALUE_BEFORE," &
            "       CAST(U.STORAGE_NOVO_GB     AS VARCHAR2(50)) AS VALUE_AFTER," &
            "       CAST(NULL AS VARCHAR2(200))                 AS SOURCE_ORIGEM," &
            "       CAST(NULL AS NUMBER(12,2))                  AS COST_USD," &
            "       CAST(NULL AS VARCHAR2(200))                 AS PART_SERIAL," &
            "       U.TECNICO                                   AS TECHNICIAN," &
            "       U.OBSERVACAO                                AS NOTES" &
            "  FROM TBL_EQUIPAMENTO_UPGRADE U" &
            " WHERE U.ID_EQUIPAMENTO = :P_ID" &
            " ORDER BY U.DATA_UPGRADE DESC"

        ' Fallback mínimo: só colunas 100% confirmadas, resto CAST(NULL)
        Dim sqlBase As String =
            "SELECT U.ID_UPGRADE," &
            "       U.DATA_UPGRADE," &
            "       U.TIPO_UPGRADE               AS COMPONENT_TYPE," &
            "       CAST(NULL AS VARCHAR2(200))  AS VALUE_BEFORE," &
            "       CAST(NULL AS VARCHAR2(200))  AS VALUE_AFTER," &
            "       CAST(NULL AS VARCHAR2(200))  AS SOURCE_ORIGEM," &
            "       CAST(NULL AS NUMBER(12,2))   AS COST_USD," &
            "       CAST(NULL AS VARCHAR2(200))  AS PART_SERIAL," &
            "       U.TECNICO                    AS TECHNICIAN," &
            "       U.OBSERVACAO                 AS NOTES" &
            "  FROM TBL_EQUIPAMENTO_UPGRADE U" &
            " WHERE U.ID_EQUIPAMENTO = :P_ID" &
            " ORDER BY U.DATA_UPGRADE DESC"

        Try
            Return OracleHelper.ExecuteDataset(Me.ConnectionString, CommandType.Text, sql,
                                               New OracleParameter() {CriarParId(pIdEquipamento)})
        Catch ex1 As Exception
            If Not IsErroColuna(ex1) Then Throw New Exception(ex1.ToString)
        End Try

        Try
            Return OracleHelper.ExecuteDataset(Me.ConnectionString, CommandType.Text, sqlStorage,
                                               New OracleParameter() {CriarParId(pIdEquipamento)})
        Catch ex2 As Exception
            If Not IsErroColuna(ex2) Then Throw New Exception(ex2.ToString)
        End Try

        Return OracleHelper.ExecuteDataset(Me.ConnectionString, CommandType.Text, sqlBase,
                                           New OracleParameter() {CriarParId(pIdEquipamento)})

    End Function

    ''' <summary>
    ''' Remessas do equipamento com dados de venda.
    ''' TBL_REMESSA: CODIGO_REMESSA (confirmado no clsGravacaoRemessa SELECT linha 144).
    '''   CONDITION_AT_SHIP NÃO existe em TBL_REMESSA_ITEM (não está no INSERT confirmado).
    '''   TBL_REMESSA_ITEM colunas confirmadas: ID_REMESSA, ID_EQUIPAMENTO,
    '''     STATUS_ITEM, SALE_PRICE_USD, NOTES, DATA_CADASTRO.
    ''' </summary>
    Public Function selecionarShipmentEquipamento(pIdEquipamento As Integer) As DataSet

        ' Tentativa 1 — query completa com colunas opcionais de TBL_REMESSA
        Dim sqlFull As String =
            "SELECT R.ID_REMESSA, R.CODIGO_REMESSA AS REMESSA_REF, R.CARRIER," &
            "       R.TRACKING_NUMBER, R.RECIPIENT_NAME," &
            "       R.ESTIMATED_DELIVERY AS DATA_ENVIO," &
            "       R.ACTUAL_DELIVERY    AS DATA_ENTREGA," &
            "       R.STATUS_REMESSA," &
            "       R.SHIPPING_COST_USD," &
            "       RI.SALE_PRICE_USD," &
            "       CAST(NULL AS VARCHAR2(100)) AS CONDITION_AT_SHIP," &
            "       RI.NOTES AS ITEM_NOTES" &
            "  FROM TBL_REMESSA R" &
            "  JOIN TBL_REMESSA_ITEM RI ON RI.ID_REMESSA = R.ID_REMESSA" &
            " WHERE RI.ID_EQUIPAMENTO = :P_ID" &
            " ORDER BY R.ID_REMESSA DESC"

        ' Tentativa 2 — sem ESTIMATED_DELIVERY / ACTUAL_DELIVERY / SHIPPING_COST_USD
        Dim sqlSemDatas As String =
            "SELECT R.ID_REMESSA, R.CODIGO_REMESSA AS REMESSA_REF, R.CARRIER," &
            "       R.TRACKING_NUMBER, R.RECIPIENT_NAME," &
            "       CAST(NULL AS DATE)           AS DATA_ENVIO," &
            "       CAST(NULL AS DATE)           AS DATA_ENTREGA," &
            "       R.STATUS_REMESSA," &
            "       CAST(NULL AS NUMBER(12,2))   AS SHIPPING_COST_USD," &
            "       RI.SALE_PRICE_USD," &
            "       CAST(NULL AS VARCHAR2(100))  AS CONDITION_AT_SHIP," &
            "       RI.NOTES AS ITEM_NOTES" &
            "  FROM TBL_REMESSA R" &
            "  JOIN TBL_REMESSA_ITEM RI ON RI.ID_REMESSA = R.ID_REMESSA" &
            " WHERE RI.ID_EQUIPAMENTO = :P_ID" &
            " ORDER BY R.ID_REMESSA DESC"

        ' Fallback mínimo — só colunas confirmadas
        Dim sqlBase As String =
            "SELECT R.ID_REMESSA," &
            "       CAST(NULL AS VARCHAR2(50))  AS REMESSA_REF, R.CARRIER," &
            "       CAST(NULL AS VARCHAR2(100)) AS TRACKING_NUMBER," &
            "       CAST(NULL AS VARCHAR2(200)) AS RECIPIENT_NAME," &
            "       CAST(NULL AS DATE)          AS DATA_ENVIO," &
            "       CAST(NULL AS DATE)          AS DATA_ENTREGA," &
            "       R.STATUS_REMESSA," &
            "       CAST(NULL AS NUMBER(12,2))  AS SHIPPING_COST_USD," &
            "       RI.SALE_PRICE_USD," &
            "       CAST(NULL AS VARCHAR2(100)) AS CONDITION_AT_SHIP," &
            "       RI.NOTES AS ITEM_NOTES" &
            "  FROM TBL_REMESSA R" &
            "  JOIN TBL_REMESSA_ITEM RI ON RI.ID_REMESSA = R.ID_REMESSA" &
            " WHERE RI.ID_EQUIPAMENTO = :P_ID" &
            " ORDER BY R.ID_REMESSA DESC"

        Try
            Return OracleHelper.ExecuteDataset(Me.ConnectionString, CommandType.Text, sqlFull,
                                               New OracleParameter() {CriarParId(pIdEquipamento)})
        Catch ex1 As Exception
            If Not IsErroColuna(ex1) Then Throw New Exception(ex1.ToString)
        End Try

        Try
            Return OracleHelper.ExecuteDataset(Me.ConnectionString, CommandType.Text, sqlSemDatas,
                                               New OracleParameter() {CriarParId(pIdEquipamento)})
        Catch ex2 As Exception
            If Not IsErroColuna(ex2) Then Throw New Exception(ex2.ToString)
        End Try

        Return OracleHelper.ExecuteDataset(Me.ConnectionString, CommandType.Text, sqlBase,
                                           New OracleParameter() {CriarParId(pIdEquipamento)})

    End Function

#End Region

#Region "Auxiliares"

    Private Function CriarParId(pId As Integer) As OracleParameter
        Dim par As New OracleParameter("P_ID", OracleDbType.Int32, ParameterDirection.Input)
        par.Value = pId
        Return par
    End Function

    Private Function IsErroColuna(ex As Exception) As Boolean
        Dim msg As String = ex.ToString().ToUpperInvariant()
        Return msg.Contains("ORA-00904") OrElse msg.Contains("ORA-00942")
    End Function

#End Region

End Class
