Imports Oracle.ManagedDataAccess.Client
Imports System.Configuration
Imports System.Data
Imports GBS_Inventory.OracleHelper

Public Class clsReadComponent

#Region "Attributes"

    Private ConnectionString As String

#End Region

#Region "Constructor"

    Public Sub New()
        Me.ConnectionString = ConfigurationManager.ConnectionStrings("OracleDB").ConnectionString
    End Sub

#End Region

#Region "Queries"

    Public Function selectByFilter(pFilter As String, pType As String, pStatus As String) As DataSet

        Dim oPar(3) As OracleParameter
        oPar(0) = New OracleParameter("P_FILTER", OracleDbType.Varchar2,  ParameterDirection.Input)
        oPar(1) = New OracleParameter("P_TYPE",   OracleDbType.Varchar2,  ParameterDirection.Input)
        oPar(2) = New OracleParameter("P_STATUS", OracleDbType.Varchar2,  ParameterDirection.Input)
        oPar(3) = New OracleParameter("P_CURSOR", OracleDbType.RefCursor, ParameterDirection.Output)

        oPar(0).Value = If(String.IsNullOrEmpty(pFilter), DBNull.Value, CObj(pFilter))
        oPar(1).Value = If(String.IsNullOrEmpty(pType),   DBNull.Value, CObj(pType))
        oPar(2).Value = If(String.IsNullOrEmpty(pStatus), DBNull.Value, CObj(pStatus))

        Try
            Return OracleHelper.ExecuteDataset(Me.ConnectionString, CommandType.StoredProcedure,
                                               "PACK_COMPONENT.PROC_SELECT_FILTER", oPar)
        Catch ex As Exception
            Throw New Exception(ex.ToString)
        End Try

    End Function

    Public Function selectById(pId As Integer) As DataSet

        Dim oPar(1) As OracleParameter
        oPar(0) = New OracleParameter("P_ID",     OracleDbType.Int32,     ParameterDirection.Input)
        oPar(1) = New OracleParameter("P_CURSOR", OracleDbType.RefCursor, ParameterDirection.Output)
        oPar(0).Value = pId

        Try
            Return OracleHelper.ExecuteDataset(Me.ConnectionString, CommandType.StoredProcedure,
                                               "PACK_COMPONENT.PROC_SELECT_ID", oPar)
        Catch ex As Exception
            Throw New Exception(ex.ToString)
        End Try

    End Function

    ' Types and statuses only ever come from the fixed whitelists in ComponentSummaryOptions,
    ' never from free text, so they are safe to inline.
    Public Function selectSummary(pOpts As Models.ComponentSummaryOptions) As DataSet

        Dim colunasGroup As String = String.Join(", ", pOpts.ColunasDescritivas)

        Dim colunasSelect As New List(Of String) From {colunasGroup, "COUNT(*) AS TOTAL"}
        For Each st As String In Models.ComponentSummaryOptions.StatusDisponiveis.Where(Function(s) pOpts.ColunasStatus.Contains(s))
            colunasSelect.Add("SUM(CASE WHEN STATUS = '" & st & "' THEN 1 ELSE 0 END) AS " & st)
        Next

        Dim whereTipo As String = ""
        If Not pOpts.TodosOsTipos Then
            Dim conds As New List(Of String)()
            Dim tipos As String() = Models.ComponentSummaryOptions.TiposConhecidos.Where(Function(t) pOpts.TiposIncluidos.Contains(t)).ToArray()
            If tipos.Length > 0 Then
                conds.Add("COMPONENT_TYPE IN (" & String.Join(",", tipos.Select(Function(t) "'" & t & "'")) & ")")
            End If
            If pOpts.IncluirOutrosTipos Then
                conds.Add("COMPONENT_TYPE NOT IN (" & String.Join(",", Models.ComponentSummaryOptions.TiposConhecidos.Select(Function(t) "'" & t & "'")) & ")")
            End If
            whereTipo = " WHERE " & If(conds.Count = 0, "1 = 0", String.Join(" OR ", conds))
        End If

        Try
            Return OracleHelper.ExecuteDataset(Me.ConnectionString, CommandType.Text,
                "SELECT " & String.Join(", ", colunasSelect) &
                "  FROM TBL_COMPONENT" & whereTipo &
                " GROUP BY " & colunasGroup &
                " ORDER BY " & colunasGroup)
        Catch ex As Exception
            Throw New Exception(ex.ToString)
        End Try

    End Function

    Public Function selectDistinctBrands() As DataSet
        Try
            Return OracleHelper.ExecuteDataset(Me.ConnectionString, CommandType.Text,
                "SELECT DISTINCT BRAND FROM TBL_COMPONENT WHERE BRAND IS NOT NULL ORDER BY BRAND")
        Catch ex As Exception
            Throw New Exception(ex.ToString)
        End Try
    End Function

    Public Function selectDistinctModels() As DataSet
        Try
            Return OracleHelper.ExecuteDataset(Me.ConnectionString, CommandType.Text,
                "SELECT DISTINCT MODEL FROM TBL_COMPONENT WHERE MODEL IS NOT NULL ORDER BY MODEL")
        Catch ex As Exception
            Throw New Exception(ex.ToString)
        End Try
    End Function

    Public Function selectDistinctCpus() As DataSet
        Try
            Return OracleHelper.ExecuteDataset(Me.ConnectionString, CommandType.Text,
                "SELECT DISTINCT CPU FROM TBL_COMPONENT WHERE CPU IS NOT NULL ORDER BY CPU")
        Catch ex As Exception
            Throw New Exception(ex.ToString)
        End Try
    End Function

#End Region

End Class
