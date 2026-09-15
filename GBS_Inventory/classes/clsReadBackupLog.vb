Imports System.Configuration
Imports System.Data
Imports GBS_Inventory.OracleHelper

Public Class clsReadBackupLog

#Region "Attributes"

    Private ConnectionString As String

#End Region

#Region "Constructor"

    Public Sub New()
        Me.ConnectionString = ConfigurationManager.ConnectionStrings("OracleDB").ConnectionString
    End Sub

#End Region

#Region "Queries"

    Public Function selectAll() As DataSet
        Try
            Return OracleHelper.ExecuteDataset(Me.ConnectionString, CommandType.Text,
                "SELECT ID_BACKUP_LOG, DATA_EXECUCAO, ARQUIVO_DMP, STATUS, MENSAGEM" &
                "  FROM TBL_BACKUP_LOG" &
                " ORDER BY DATA_EXECUCAO DESC")
        Catch ex As Exception
            Throw New Exception(ex.ToString)
        End Try
    End Function

#End Region

End Class
