Imports System.Data
Imports GBS_Inventory.Models

Public Class ComponentController

#Region "Attributes"

    Private oReader As clsReadComponent
    Private oWriter As clsWriteComponent

#End Region

#Region "Constructor"

    Public Sub New()
        oReader = New clsReadComponent()
        oWriter = New clsWriteComponent()
    End Sub

#End Region

#Region "Queries"

    Public Function fetchAll() As DataSet
        Return oReader.selectByFilter(Nothing, Nothing, Nothing)
    End Function

    Public Function fetchByFilter(pFilter As String, pType As String, pStatus As String) As DataSet
        Return oReader.selectByFilter(pFilter, pType, pStatus)
    End Function

    Public Function fetchBrands() As DataSet
        Return oReader.selectDistinctBrands()
    End Function

    Public Function fetchModels() As DataSet
        Return oReader.selectDistinctModels()
    End Function

    Public Function fetchCpus() As DataSet
        Return oReader.selectDistinctCpus()
    End Function

    Public Function fetchSummary(pOpts As ComponentSummaryOptions) As DataSet
        Return oReader.selectSummary(pOpts)
    End Function

#End Region

#Region "Operations"

    Public Function add(pComp As Component) As String
        Try
            Return oWriter.insertComponent(pComp)
        Catch ex As Exception
            Throw New Exception("Error adding component: " & ex.Message)
        End Try
    End Function

    Public Sub updateStatus(pId As Integer, pStatus As String)
        Try
            oWriter.updateStatus(pId, pStatus)
        Catch ex As Exception
            Throw New Exception("Error updating status: " & ex.Message)
        End Try
    End Sub

    Public Sub mudarStatusEmLote(pIds As List(Of Integer), pNovoStatus As String, pNota As String)
        If pIds Is Nothing OrElse pIds.Count = 0 Then Return
        Try
            oWriter.updateStatusBatch(pIds, pNovoStatus, pNota)
        Catch ex As Exception
            Throw New Exception("Error updating status (no item was changed): " & ex.Message)
        End Try
    End Sub

#End Region

End Class
