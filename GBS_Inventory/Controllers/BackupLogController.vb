Imports System.Data

Public Class BackupLogController

#Region "Attributes"

    Private oReader As clsReadBackupLog

#End Region

#Region "Constructor"

    Public Sub New()
        oReader = New clsReadBackupLog()
    End Sub

#End Region

#Region "Queries"

    Public Function fetchAll() As DataSet
        Return oReader.selectAll()
    End Function

#End Region

End Class
