Imports Oracle.ManagedDataAccess.Client
Imports System.Configuration
Imports System.Data
Imports System.Text
Imports System.Text.RegularExpressions
Imports GBS_Inventory.OracleHelper
Imports GBS_Inventory.Models

''' <summary>
''' Classe de gravacao de upgrades.
''' Ajustada para a assinatura atual do PACK_UPGRADE no Oracle.
''' </summary>
Public Class clsGravacaoUpgrade

#Region "Atributos"

    Private ConnectionString As String

    Private oTransacao As OracleTransaction
    Private oConexao As OracleConnection

#End Region

#Region "Construtor"

    Public Sub New()

        Try

            Me.ConnectionString = ConfigurationManager.ConnectionStrings("OracleDB").ConnectionString

        Catch ex As Exception

            Throw New Exception(ex.ToString)

        End Try

        oConexao = Nothing

    End Sub

#End Region

#Region "Transacao"

    Public Function beginTransacao() As Boolean

        Dim oCon As New OracleConnection(Me.ConnectionString)

        Try

            oCon.Open()

            Me.oConexao = oCon

            Me.oTransacao = Me.oConexao.BeginTransaction

            Return True

        Catch ex As Exception

            Throw New Exception(ex.ToString)

        End Try

    End Function

    Public Function commitTransacao() As Boolean

        Try

            Me.oTransacao.Commit()

            Me.oConexao.Close()

            Return True

        Catch ex As Exception

            Throw New Exception(ex.ToString)

        End Try

    End Function

    Public Function rollbackTransacao() As Boolean

        Try

            Me.oTransacao.Rollback()

            Me.oConexao.Close()

            Return True

        Catch ex As Exception

            Throw New Exception(ex.ToString)

        End Try

    End Function

#End Region

#Region "Metodos Transacionais"

    Private Function ExtrairNumeroInteiro(pValor As String) As Integer

        If String.IsNullOrWhiteSpace(pValor) Then
            Return 0
        End If

        Dim vMatch As Match = Regex.Match(pValor, "\d+")

        If Not vMatch.Success Then
            Return 0
        End If

        Dim vNumero As Integer
        Integer.TryParse(vMatch.Value, vNumero)
        Return vNumero

    End Function

    Private Function MontarObservacao(pUpgrade As Upgrade) As String

        Dim vSb As New StringBuilder()

        If Not String.IsNullOrWhiteSpace(pUpgrade.Notes) Then
            vSb.Append(pUpgrade.Notes.Trim())
        End If

        If Not String.IsNullOrWhiteSpace(pUpgrade.PartSerial) Then
            If vSb.Length > 0 Then vSb.Append(" | ")
            vSb.Append("PartSerial: ").Append(pUpgrade.PartSerial.Trim())
        End If

        If Not String.IsNullOrWhiteSpace(pUpgrade.SourceOrigem) Then
            If vSb.Length > 0 Then vSb.Append(" | ")
            vSb.Append("Source: ").Append(pUpgrade.SourceOrigem.Trim())
        End If

        If pUpgrade.CostUsd.HasValue Then
            If vSb.Length > 0 Then vSb.Append(" | ")
            vSb.Append("CostUSD: ").Append(pUpgrade.CostUsd.Value.ToString("0.00"))
        End If

        Return If(vSb.Length = 0, Nothing, vSb.ToString())

    End Function

    Public Function incluirUpgrade(pUpgrade As Upgrade) As Integer

        Dim oPar(10) As OracleParameter
        Dim vRamAnterior As Integer = 0
        Dim vRamNova As Integer = 0
        Dim vStorageAnterior As Integer = 0
        Dim vStorageNova As Integer = 0
        Dim vTipo As String = If(pUpgrade.ComponentType, String.Empty).Trim().ToUpperInvariant()
        Dim vObservacao As String = MontarObservacao(pUpgrade)

        Dim isRam     As Boolean = (vTipo = "RAM")
        Dim isStorage As Boolean = (vTipo = "SSD" OrElse vTipo = "HDD" OrElse vTipo = "STORAGE")

        Select Case vTipo
            Case "RAM"
                vRamAnterior = ExtrairNumeroInteiro(pUpgrade.ValueBefore)
                vRamNova = ExtrairNumeroInteiro(pUpgrade.ValueAfter)

            Case "SSD", "HDD", "STORAGE"
                vStorageAnterior = ExtrairNumeroInteiro(pUpgrade.ValueBefore)
                vStorageNova = ExtrairNumeroInteiro(pUpgrade.ValueAfter)

            Case Else
                ' Componentes nao mapeados ainda nao alteram RAM/STORAGE no banco.
        End Select

        oPar(0)  = New OracleParameter("P_ID_EQUIPAMENTO",      OracleDbType.Int32,    ParameterDirection.Input)
        oPar(1)  = New OracleParameter("P_TIPO_UPGRADE",         OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(2)  = New OracleParameter("P_RAM_ANTERIOR_GB",      OracleDbType.Int32,    ParameterDirection.Input)
        oPar(3)  = New OracleParameter("P_RAM_NOVA_GB",          OracleDbType.Int32,    ParameterDirection.Input)
        oPar(4)  = New OracleParameter("P_STORAGE_ANTERIOR_GB",  OracleDbType.Int32,    ParameterDirection.Input)
        oPar(5)  = New OracleParameter("P_STORAGE_NOVO_GB",      OracleDbType.Int32,    ParameterDirection.Input)
        oPar(6)  = New OracleParameter("P_TECNICO",              OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(7)  = New OracleParameter("P_OBSERVACAO",           OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(8)  = New OracleParameter("P_ID_COMPONENT",         OracleDbType.Int32,    ParameterDirection.Input)
        oPar(9)  = New OracleParameter("P_ACTION_TYPE",          OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(10) = New OracleParameter("P_COMP_NEW_STATUS",      OracleDbType.Varchar2, ParameterDirection.Input)

        Try

            ' REVERTED 2026-09-16: sending 0 instead of DBNull for the inapplicable pair
            ' avoids the ORA-01400 crash, but PACK_UPGRADE's cascade UPDATE uses
            ' "IS NOT NULL" to decide whether to touch RAM_GB/STORAGE_GB on TBL_EQUIPAMENTO
            ' — 0 satisfies that check, so it was silently zeroing out RAM_GB/STORAGE_GB on
            ' every non-RAM/Storage upgrade. Back to DBNull (crashes loudly, doesn't corrupt
            ' data) until the real fix — relaxing the NOT NULL constraint on the 4 log
            ' columns so genuine NULL can be sent again — is applied to the schema.
            oPar(0).Value  = pUpgrade.IdEquipamento
            oPar(1).Value  = vTipo
            oPar(2).Value  = If(isRam,     CObj(vRamAnterior),    DBNull.Value)
            oPar(3).Value  = If(isRam,     CObj(vRamNova),         DBNull.Value)
            oPar(4).Value  = If(isStorage, CObj(vStorageAnterior), DBNull.Value)
            oPar(5).Value  = If(isStorage, CObj(vStorageNova),     DBNull.Value)
            oPar(6).Value  = If(String.IsNullOrWhiteSpace(pUpgrade.Technician),    DBNull.Value, CObj(pUpgrade.Technician.Trim()))
            oPar(7).Value  = If(String.IsNullOrWhiteSpace(vObservacao),            DBNull.Value, CObj(vObservacao))
            oPar(8).Value  = If(Not pUpgrade.IdComponent.HasValue,                 DBNull.Value, CObj(pUpgrade.IdComponent.Value))
            oPar(9).Value  = If(String.IsNullOrWhiteSpace(pUpgrade.ActionType),    CObj("MANUAL"), CObj(pUpgrade.ActionType.Trim().ToUpperInvariant()))
            oPar(10).Value = If(String.IsNullOrWhiteSpace(pUpgrade.CompNewStatus), DBNull.Value, CObj(pUpgrade.CompNewStatus.Trim().ToUpperInvariant()))

            OracleHelper.ExecuteNonQuery(Me.oTransacao, CommandType.StoredProcedure, "PACK_UPGRADE.PROC_INSERT_UPGRADE", oPar)

        Catch ex As Exception

            Throw New Exception(ex.ToString)

        End Try

        Return 0

    End Function

    Public Function excluirUpgrade(pIdUpgrade As Integer) As Boolean

        Dim oPar(0) As OracleParameter

        oPar(0) = New OracleParameter("V_ID", OracleDbType.Int32, ParameterDirection.Input)

        Try

            oPar(0).Value = pIdUpgrade

            OracleHelper.ExecuteNonQuery(Me.oTransacao, CommandType.StoredProcedure, "PACK_UPGRADE.PROC_DELETE_UPGRADE", oPar)

        Catch ex As Exception

            Throw New Exception(ex.ToString)

        End Try

        Return True

    End Function

#End Region

End Class
