Imports Oracle.ManagedDataAccess.Client
Imports System.Configuration
Imports System.Data
Imports GBS_Inventory.OracleHelper

''' <summary>
''' Classe de leitura de Upgrades
''' </summary>
Public Class clsLeituraUpgrade

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

    Public Function selecionarUpgradesEquipamento(pIdEquipamento As Integer) As DataSet

        Dim oPar(1) As OracleParameter

        oPar(0) = New OracleParameter("V_ID_EQUIPAMENTO", OracleDbType.Int32,     ParameterDirection.Input)
        oPar(1) = New OracleParameter("V_CURSOR",         OracleDbType.RefCursor, ParameterDirection.Output)

        Try

            oPar(0).Value = pIdEquipamento

            Return OracleHelper.ExecuteDataset(Me.ConnectionString, CommandType.StoredProcedure, "PACK_UPGRADE.PROC_SELECT_UPGRADES_EQUIP", oPar)

        Catch ex As Exception

            Throw New Exception(ex.ToString)

        End Try

    End Function

    Public Function selecionarUpgradesRecentes(pDias As Integer) As DataSet

        Dim oPar(1) As OracleParameter

        oPar(0) = New OracleParameter("V_DIAS",   OracleDbType.Int32,     ParameterDirection.Input)
        oPar(1) = New OracleParameter("V_CURSOR", OracleDbType.RefCursor, ParameterDirection.Output)

        Try

            oPar(0).Value = pDias

            Return OracleHelper.ExecuteDataset(Me.ConnectionString, CommandType.StoredProcedure, "PACK_UPGRADE.PROC_SELECT_UPGRADES_RECENTES", oPar)

        Catch ex As Exception

            Throw New Exception(ex.ToString)

        End Try

    End Function

#End Region

End Class
