Imports Oracle.ManagedDataAccess.Client
Imports System.Configuration
Imports System.Data
Imports GBS_Inventory.OracleHelper
Imports GBS_Inventory.Models

Public Class clsWriteComponent

#Region "Attributes"

    Private ConnectionString As String

#End Region

#Region "Constructor"

    Public Sub New()
        Me.ConnectionString = ConfigurationManager.ConnectionStrings("OracleDB").ConnectionString
    End Sub

#End Region

#Region "Operations"

    Public Function insertComponent(pComp As Component) As String

        Dim oPar(10) As OracleParameter
        oPar(0)  = New OracleParameter("P_COMPONENT_TYPE", OracleDbType.Varchar2,  ParameterDirection.Input)
        oPar(1)  = New OracleParameter("P_CAPACITY_GB",    OracleDbType.Int32,     ParameterDirection.Input)
        oPar(2)  = New OracleParameter("P_SPEED_MHZ",      OracleDbType.Int32,     ParameterDirection.Input)
        oPar(3)  = New OracleParameter("P_GENERATION",     OracleDbType.Varchar2,  ParameterDirection.Input)
        oPar(4)  = New OracleParameter("P_BRAND",          OracleDbType.Varchar2,  ParameterDirection.Input)
        oPar(5)  = New OracleParameter("P_PART_NUMBER",    OracleDbType.Varchar2,  ParameterDirection.Input)
        oPar(6)  = New OracleParameter("P_CONDITION",      OracleDbType.Varchar2,  ParameterDirection.Input)
        oPar(7)  = New OracleParameter("P_STATUS",         OracleDbType.Varchar2,  ParameterDirection.Input)
        oPar(8)  = New OracleParameter("P_SOURCE_BATCH",   OracleDbType.Varchar2,  ParameterDirection.Input)
        oPar(9)  = New OracleParameter("P_NOTES",          OracleDbType.Varchar2,  ParameterDirection.Input)
        oPar(10)           = New OracleParameter("P_UID_OUT", OracleDbType.Varchar2)
        oPar(10).Direction = ParameterDirection.Output
        oPar(10).Size      = 20

        oPar(0).Value = pComp.ComponentType
        oPar(1).Value = pComp.CapacityGB
        oPar(2).Value = If(pComp.SpeedMhz.HasValue, CObj(pComp.SpeedMhz.Value), DBNull.Value)
        oPar(3).Value = If(String.IsNullOrEmpty(pComp.Generation),  DBNull.Value, CObj(pComp.Generation))
        oPar(4).Value = If(String.IsNullOrEmpty(pComp.Brand),       DBNull.Value, CObj(pComp.Brand))
        oPar(5).Value = If(String.IsNullOrEmpty(pComp.PartNumber),  DBNull.Value, CObj(pComp.PartNumber))
        oPar(6).Value = If(String.IsNullOrEmpty(pComp.ConditionStatus), DBNull.Value, CObj(pComp.ConditionStatus))
        oPar(7).Value = If(String.IsNullOrEmpty(pComp.Status),      DBNull.Value, CObj(pComp.Status))
        oPar(8).Value = If(String.IsNullOrEmpty(pComp.SourceBatch), DBNull.Value, CObj(pComp.SourceBatch))
        oPar(9).Value = If(String.IsNullOrEmpty(pComp.Notes),       DBNull.Value, CObj(pComp.Notes))

        Try
            OracleHelper.ExecuteNonQuery(Me.ConnectionString, CommandType.StoredProcedure,
                                         "PACK_COMPONENT.PROC_INSERT", oPar)
            Return oPar(10).Value.ToString()
        Catch ex As Exception
            Throw New Exception(ex.ToString)
        End Try

    End Function

    Public Sub updateStatus(pId As Integer, pStatus As String)

        Dim oPar(1) As OracleParameter
        oPar(0) = New OracleParameter("P_ID",     OracleDbType.Int32,    ParameterDirection.Input)
        oPar(1) = New OracleParameter("P_STATUS", OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(0).Value = pId
        oPar(1).Value = pStatus

        Try
            OracleHelper.ExecuteNonQuery(Me.ConnectionString, CommandType.StoredProcedure,
                                         "PACK_COMPONENT.PROC_UPDATE_STATUS", oPar)
        Catch ex As Exception
            Throw New Exception(ex.ToString)
        End Try

    End Sub

#End Region

End Class
