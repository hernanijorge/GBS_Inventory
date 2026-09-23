Imports Oracle.ManagedDataAccess.Client
Imports System.Configuration
Imports System.Data
Imports GBS_Inventory.OracleHelper

''' <summary>
''' Classe de leitura de Remessas
''' </summary>
Public Class clsLeituraRemessa

#Region "Atributos"

    Private ConnectionString As String

#End Region

#Region "Construtor"

    Public Sub New()

        Try

            Me.ConnectionString = ConfigurationManager.ConnectionStrings("OracleDB").ConnectionString

        Catch ex As Exception

            Throw New Exception(ex.ToString)

        End Try

    End Sub

#End Region

#Region "Consultas"

#Region "Relatório por Destinatário"

    ' Shipments in the report's scope, one row per shipment, with:
    '   RECIPIENT  = UPPER(TRIM(DESTINATARIO))  ("MACIR" / "Macir " are the same spelling)
    '   CLIENT_KEY = grouping key for the chosen mode (spelling / first N chars / canonical alias)
    ' Statuses, prefix length and ids come from whitelists/integers (never free text) and are
    ' inlined; dates and recipient are bound.
    Private Function MontarBaseRelatorioDestinatario(pOpts As Models.ShipmentReportOptions,
                                                     pAplicarRecipient As Boolean,
                                                     pAplicarIds As Boolean,
                                                     ByRef pPar As OracleParameter()) As String

        Dim sNome  As String = "UPPER(TRIM(R.DESTINATARIO))"
        Dim sJoin  As String = ""
        Dim sChave As String
        Dim sChaveDoFiltro As String   ' same key applied to :P_DEST (already upper/trimmed)
        Select Case pOpts.Agrupamento
            Case Models.ShipmentReportOptions.AGRUP_PREFIXO
                Dim n As String = Math.Max(Models.ShipmentReportOptions.PREFIXO_MIN,
                                           Math.Min(Models.ShipmentReportOptions.PREFIXO_MAX, pOpts.PrefixLen)).ToString()
                sChave = "SUBSTR(" & sNome & ", 1, " & n & ")"
                sChaveDoFiltro = "SUBSTR(:P_DEST, 1, " & n & ")"
            Case Models.ShipmentReportOptions.AGRUP_ALIAS
                sJoin  = " LEFT JOIN TBL_CLIENT_ALIAS CA ON CA.ALIAS_NAME = " & sNome
                sChave = "NVL(CA.CANONICAL_NAME, " & sNome & ")"
                sChaveDoFiltro = "NVL((SELECT MAX(A2.CANONICAL_NAME) FROM TBL_CLIENT_ALIAS A2 WHERE A2.ALIAS_NAME = :P_DEST), :P_DEST)"
            Case Else
                sChave = sNome
                sChaveDoFiltro = ":P_DEST"
        End Select

        Dim listaStatus As String = String.Join(",", pOpts.StatusIncluidos.Select(Function(s) "'" & s.Replace("'", "''") & "'"))
        If listaStatus = "" Then listaStatus = "NULL"

        Dim pars As New List(Of OracleParameter) From {
            New OracleParameter("P_INI", OracleDbType.Date, ParameterDirection.Input) With {.Value = pOpts.DataInicio.Date},
            New OracleParameter("P_FIM", OracleDbType.Date, ParameterDirection.Input) With {.Value = pOpts.DataFim.Date.AddDays(1)}
        }

        Dim sWhere As String =
            " WHERE R.DATA_CADASTRO >= :P_INI AND R.DATA_CADASTRO < :P_FIM" &
            "   AND R.STATUS_REMESSA IN (" & listaStatus & ")"

        ' :P_DEST is only added when used — ODP.NET (BindByName) rejects unused parameters
        If pAplicarRecipient AndAlso Not String.IsNullOrWhiteSpace(pOpts.RecipientFiltro) Then
            sWhere &= "   AND " & sChave & " = " & sChaveDoFiltro
            pars.Add(New OracleParameter("P_DEST", OracleDbType.Varchar2, ParameterDirection.Input) With {
                         .Value = pOpts.RecipientFiltro.Trim().ToUpperInvariant()})
        End If

        If pAplicarIds AndAlso pOpts.IdsSelecionados IsNot Nothing Then
            If pOpts.IdsSelecionados.Count = 0 Then
                sWhere &= "   AND 1 = 0"
            Else
                ' Oracle caps IN lists at 1000 items
                Dim blocos As New List(Of String)()
                For i As Integer = 0 To pOpts.IdsSelecionados.Count - 1 Step 1000
                    blocos.Add("R.ID_REMESSA IN (" & String.Join(",", pOpts.IdsSelecionados.Skip(i).Take(1000)) & ")")
                Next
                sWhere &= "   AND (" & String.Join(" OR ", blocos) & ")"
            End If
        End If

        pPar = pars.ToArray()

        Return "SELECT R.ID_REMESSA, R.CODIGO_REMESSA AS REMESSA_REF, R.CARRIER, R.TRACKING_NUMBER," &
               "       " & sNome & " AS RECIPIENT, " & sChave & " AS CLIENT_KEY, R.STATUS_REMESSA," &
               "       R.DATA_ENVIO, R.DATA_ENTREGA, R.DATA_CADASTRO," &
               "       (SELECT COUNT(*) FROM TBL_REMESSA_ITEM I WHERE I.ID_REMESSA = R.ID_REMESSA) AS TOTAL_ITEMS" &
               "  FROM TBL_REMESSA R" & sJoin & sWhere

    End Function

    Private Shared Function SomaStatus(pChave As String) As String
        Dim grupo = Models.ShipmentReportOptions.GruposStatus.First(Function(g) g.Chave = pChave)
        Return "SUM(CASE WHEN STATUS_REMESSA IN (" & String.Join(",", grupo.Status.Select(Function(s) "'" & s & "'")) & ") THEN 1 ELSE 0 END)"
    End Function

    ' One row per client group (CLIENT_KEY)
    Public Function selecionarResumoPorDestinatario(pOpts As Models.ShipmentReportOptions) As DataSet

        Dim oPar As OracleParameter() = Nothing
        Dim sBase As String = MontarBaseRelatorioDestinatario(pOpts, True, True, oPar)

        Dim sSql As String =
            "SELECT CLIENT_KEY," &
            "       COUNT(*)                     AS TOTAL_SHIPMENTS," &
            "       " & SomaStatus("DELIVERED")  & " AS DELIVERED_CNT," &
            "       " & SomaStatus("IN_TRANSIT") & " AS IN_TRANSIT_CNT," &
            "       " & SomaStatus("PENDING")    & " AS PENDING_CNT," &
            "       " & SomaStatus("ISSUES")     & " AS ISSUES_CNT," &
            "       SUM(TOTAL_ITEMS)             AS TOTAL_ITEMS," &
            "       COUNT(DISTINCT CARRIER)      AS CARRIERS_USED," &
            "       MIN(DATA_CADASTRO)           AS FIRST_SHIPMENT," &
            "       MAX(DATA_CADASTRO)           AS LAST_SHIPMENT" &
            "  FROM (" & sBase & ")" &
            " GROUP BY CLIENT_KEY" &
            " ORDER BY TOTAL_SHIPMENTS DESC, CLIENT_KEY"

        Try
            Return OracleHelper.ExecuteDataset(Me.ConnectionString, CommandType.Text, sSql, oPar)
        Catch ex As Exception
            Throw New Exception(ex.ToString)
        End Try

    End Function

    ' One row per shipment in the report (Detailed breakdown, Variations column)
    Public Function selecionarDetalhePorDestinatario(pOpts As Models.ShipmentReportOptions) As DataSet

        Dim oPar As OracleParameter() = Nothing
        Dim sSql As String = MontarBaseRelatorioDestinatario(pOpts, True, True, oPar) & " ORDER BY CLIENT_KEY, R.DATA_CADASTRO"

        Try
            Return OracleHelper.ExecuteDataset(Me.ConnectionString, CommandType.Text, sSql, oPar)
        Catch ex As Exception
            Throw New Exception(ex.ToString)
        End Try

    End Function

    ' Shipments matching the dialog filters, ignoring the manual selection (feeds the checklist)
    Public Function selecionarRemessasParaSelecao(pOpts As Models.ShipmentReportOptions) As DataSet

        Dim oPar As OracleParameter() = Nothing
        Dim sSql As String = MontarBaseRelatorioDestinatario(pOpts, True, False, oPar) & " ORDER BY R.DATA_CADASTRO DESC"

        Try
            Return OracleHelper.ExecuteDataset(Me.ConnectionString, CommandType.Text, sSql, oPar)
        Catch ex As Exception
            Throw New Exception(ex.ToString)
        End Try

    End Function

    ' Merge preview: distinct spellings vs distinct groups in the date/status scope
    ' (recipient filter and manual selection ignored). Columns: NOMES, GRUPOS.
    Public Function selecionarPreviewAgrupamento(pOpts As Models.ShipmentReportOptions) As DataSet

        Dim oPar As OracleParameter() = Nothing
        Dim sSql As String = "SELECT COUNT(DISTINCT RECIPIENT) AS NOMES, COUNT(DISTINCT CLIENT_KEY) AS GRUPOS" &
                             "  FROM (" & MontarBaseRelatorioDestinatario(pOpts, False, False, oPar) & ")"

        Try
            Return OracleHelper.ExecuteDataset(Me.ConnectionString, CommandType.Text, sSql, oPar)
        Catch ex As Exception
            Throw New Exception(ex.ToString)
        End Try

    End Function

    Public Function selecionarDestinatarios() As DataSet
        Try
            Return OracleHelper.ExecuteDataset(Me.ConnectionString, CommandType.Text,
                "SELECT DISTINCT UPPER(TRIM(DESTINATARIO)) AS RECIPIENT FROM TBL_REMESSA" &
                " WHERE DESTINATARIO IS NOT NULL ORDER BY 1")
        Catch ex As Exception
            Throw New Exception(ex.ToString)
        End Try
    End Function

    Public Function selecionarAliases() As DataSet
        Try
            Return OracleHelper.ExecuteDataset(Me.ConnectionString, CommandType.Text,
                "SELECT CANONICAL_NAME, ALIAS_NAME FROM TBL_CLIENT_ALIAS ORDER BY CANONICAL_NAME, ALIAS_NAME")
        Catch ex As Exception
            Throw New Exception(ex.ToString)
        End Try
    End Function

#End Region

    Public Function selecionarRemessasAtivas() As DataSet

        Dim oPar(0) As OracleParameter

        oPar(0) = New OracleParameter("P_CURSOR", OracleDbType.RefCursor, ParameterDirection.Output)

        Try

            Return OracleHelper.ExecuteDataset(Me.ConnectionString, CommandType.StoredProcedure, "PACK_REMESSA.PROC_SELECT_REMESSAS", oPar)

        Catch ex As Exception

            Throw New Exception(ex.ToString)

        End Try

    End Function

    Public Function selecionarRemessa(pIdRemessa As Integer) As DataSet

        Dim oPar(0) As OracleParameter

        oPar(0) = New OracleParameter("P_ID", OracleDbType.Int32, ParameterDirection.Input)

        Try

            oPar(0).Value = pIdRemessa

            Return OracleHelper.ExecuteDataset(Me.ConnectionString, CommandType.Text,
                "SELECT * FROM TBL_REMESSA WHERE ID_REMESSA = :P_ID", oPar)

        Catch ex As Exception

            Throw New Exception(ex.ToString)

        End Try

    End Function

    Public Function selecionarItensRemessa(pIdRemessa As Integer) As DataSet

        Dim oPar(0) As OracleParameter

        oPar(0) = New OracleParameter("P_ID_REMESSA", OracleDbType.Int32, ParameterDirection.Input)

        Try

            oPar(0).Value = pIdRemessa

            Return OracleHelper.ExecuteDataset(Me.ConnectionString, CommandType.Text,
                "SELECT RI.ID_REMESSA_ITEM, RI.ID_REMESSA, RI.ID_EQUIPAMENTO, RI.STATUS_ITEM, " &
                "       RI.SALE_PRICE_USD, RI.NOTES, RI.DATA_CADASTRO, " &
                "       E.INTERNAL_UID, E.MARCA, E.MODEL, E.SERIAL_NUMBER, " &
                "       E.PROCESSADOR, E.RAM_GB, E.STORAGE_GB, E.CONDITION_STATUS " &
                "  FROM TBL_REMESSA_ITEM RI " &
                "  JOIN TBL_EQUIPAMENTO  E ON E.ID_EQUIPAMENTO = RI.ID_EQUIPAMENTO " &
                " WHERE RI.ID_REMESSA = :P_ID_REMESSA " &
                " ORDER BY E.INTERNAL_UID", oPar)

        Catch ex As Exception

            Throw New Exception(ex.ToString)

        End Try

    End Function

    Public Function selecionarEquipamentosDisponiveis(pPesquisa As String) As DataSet

        Dim oPar(2) As OracleParameter
        oPar(0) = New OracleParameter("P_PESQ_UID",   OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(1) = New OracleParameter("P_PESQ_MODEL", OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(2) = New OracleParameter("P_PESQ_MARCA", OracleDbType.Varchar2, ParameterDirection.Input)

        Dim sFiltro As String = "%" & If(String.IsNullOrWhiteSpace(pPesquisa), "", pPesquisa.Trim().ToUpperInvariant()) & "%"

        oPar(0).Value = sFiltro
        oPar(1).Value = sFiltro
        oPar(2).Value = sFiltro

        Try

            Return OracleHelper.ExecuteDataset(Me.ConnectionString, CommandType.Text,
                "SELECT E.ID_EQUIPAMENTO, E.INTERNAL_UID, E.SERIAL_NUMBER, " &
                "       E.MARCA, E.MODEL, E.PROCESSADOR, E.RAM_GB, E.STORAGE_GB, E.STATUS " &
                "  FROM TBL_EQUIPAMENTO E " &
                " WHERE E.STATUS IN ('IN_STOCK','AVAILABLE','IN_REPAIR') " &
                "   AND (UPPER(NVL(E.INTERNAL_UID,''))  LIKE :P_PESQ_UID " &
                "    OR  UPPER(NVL(E.MODEL,''))          LIKE :P_PESQ_MODEL " &
                "    OR  UPPER(NVL(E.MARCA,''))          LIKE :P_PESQ_MARCA) " &
                " ORDER BY E.INTERNAL_UID", oPar)

        Catch ex As Exception

            Throw New Exception(ex.ToString)

        End Try

    End Function

#End Region

End Class
