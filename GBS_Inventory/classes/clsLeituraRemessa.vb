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

        oPar(0) = New OracleParameter("V_CURSOR", OracleDbType.RefCursor, ParameterDirection.Output)

        Try

            Return OracleHelper.ExecuteDataset(Me.ConnectionString, CommandType.StoredProcedure, "PACK_REMESSA.PROC_SELECT_REMESSAS_ATIVAS", oPar)

        Catch ex As Exception

            Throw New Exception(ex.ToString)

        End Try

    End Function

    Public Function selecionarRemessa(pIdRemessa As Integer) As DataSet

        Dim oPar(1) As OracleParameter

        oPar(0) = New OracleParameter("V_ID",     OracleDbType.Int32,     ParameterDirection.Input)
        oPar(1) = New OracleParameter("V_CURSOR", OracleDbType.RefCursor, ParameterDirection.Output)

        Try

            oPar(0).Value = pIdRemessa

            Return OracleHelper.ExecuteDataset(Me.ConnectionString, CommandType.StoredProcedure, "PACK_REMESSA.PROC_SELECT_REMESSA", oPar)

        Catch ex As Exception

            Throw New Exception(ex.ToString)

        End Try

    End Function

    Public Function selecionarItensRemessa(pIdRemessa As Integer) As DataSet

        Dim oPar(1) As OracleParameter

        oPar(0) = New OracleParameter("V_ID_REMESSA", OracleDbType.Int32,     ParameterDirection.Input)
        oPar(1) = New OracleParameter("V_CURSOR",     OracleDbType.RefCursor, ParameterDirection.Output)

        Try

            oPar(0).Value = pIdRemessa

            Return OracleHelper.ExecuteDataset(Me.ConnectionString, CommandType.StoredProcedure, "PACK_REMESSA.PROC_SELECT_ITENS_REMESSA", oPar)

        Catch ex As Exception

            Throw New Exception(ex.ToString)

        End Try

    End Function

#End Region

End Class
