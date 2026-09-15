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

    Public Function selectSummary() As DataSet
        Try
            Return OracleHelper.ExecuteDataset(Me.ConnectionString, CommandType.Text,
                "SELECT COMPONENT_TYPE, CAPACITY_GB, GENERATION, SPEED_MHZ, CPU, STORAGE_GB," &
                "       COUNT(*)                                                  AS TOTAL,"     &
                "       SUM(CASE WHEN STATUS='IN_STOCK'   THEN 1 ELSE 0 END)      AS IN_STOCK,"  &
                "       SUM(CASE WHEN STATUS='INSTALLED'  THEN 1 ELSE 0 END)      AS INSTALLED," &
                "       SUM(CASE WHEN STATUS='SOLD'       THEN 1 ELSE 0 END)      AS SOLD,"      &
                "       SUM(CASE WHEN STATUS='SCRAPPED'   THEN 1 ELSE 0 END)      AS SCRAPPED"   &
                "  FROM TBL_COMPONENT" &
                " GROUP BY COMPONENT_TYPE, CAPACITY_GB, GENERATION, SPEED_MHZ, CPU, STORAGE_GB" &
                " ORDER BY COMPONENT_TYPE, CAPACITY_GB, GENERATION, SPEED_MHZ, CPU, STORAGE_GB")
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

#End Region

End Class
