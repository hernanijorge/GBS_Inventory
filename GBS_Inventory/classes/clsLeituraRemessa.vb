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
                " WHERE E.STATUS = 'IN_STOCK' " &
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
