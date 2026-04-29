Imports Oracle.ManagedDataAccess.Client
Imports Oracle.ManagedDataAccess.Types
Imports System.Configuration
Imports GBS_Inventory.OracleHelper
Imports GBS_Inventory.Models

''' <summary>
''' Classe de gravacao de Invoice
''' </summary>
Public Class clsGravacaoInvoice

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
            Throw New Exception(ex.Message)
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
            Me.oTransacao = Me.oConexao.BeginTransaction()
            Return True
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Function

    Public Function commitTransacao() As Boolean

        Try
            Me.oTransacao.Commit()
            Me.oConexao.Close()
            Return True
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Function

    Public Function rollbackTransacao() As Boolean

        Try
            Me.oTransacao.Rollback()
            Me.oConexao.Close()
            Return True
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Function

#End Region

#Region "Metodos Transacionais"

    Public Function incluirInvoice(pInvoice As Invoice, ByRef pInvoiceNumber As String) As Integer

        Dim oPar(12) As OracleParameter

        oPar(0) = New OracleParameter("V_INVOICE_NUMBER", OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(1) = New OracleParameter("V_ID_CLIENTE", OracleDbType.Int32, ParameterDirection.Input)
        oPar(2) = New OracleParameter("V_ID_REMESSA", OracleDbType.Int32, ParameterDirection.Input)
        oPar(3) = New OracleParameter("V_ISSUE_DATE", OracleDbType.Date, ParameterDirection.Input)
        oPar(4) = New OracleParameter("V_DUE_DATE", OracleDbType.Date, ParameterDirection.Input)
        oPar(5) = New OracleParameter("V_CURRENCY", OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(6) = New OracleParameter("V_DISCOUNT_USD", OracleDbType.Decimal, ParameterDirection.Input)
        oPar(7) = New OracleParameter("V_SHIPPING_USD", OracleDbType.Decimal, ParameterDirection.Input)
        oPar(8) = New OracleParameter("V_TAX_USD", OracleDbType.Decimal, ParameterDirection.Input)
        oPar(9) = New OracleParameter("V_CLIENT_EMAIL", OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(10) = New OracleParameter("V_NOTES", OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(11) = New OracleParameter("V_ID", OracleDbType.Int32, ParameterDirection.Output)
        oPar(12) = New OracleParameter("V_NUMBER_OUT", OracleDbType.Varchar2, 40, Nothing, ParameterDirection.Output)

        Dim vIdGerado As OracleDecimal

        Try

            oPar(0).Value = If(String.IsNullOrWhiteSpace(pInvoice.InvoiceNumber), DBNull.Value, CObj(pInvoice.InvoiceNumber))
            oPar(1).Value = pInvoice.IdCliente
            oPar(2).Value = If(pInvoice.IdRemessa.HasValue, CObj(pInvoice.IdRemessa.Value), DBNull.Value)
            oPar(3).Value = pInvoice.IssueDate
            oPar(4).Value = If(pInvoice.DueDate.HasValue, CObj(pInvoice.DueDate.Value), DBNull.Value)
            oPar(5).Value = If(String.IsNullOrWhiteSpace(pInvoice.Currency), "USD", pInvoice.Currency)
            oPar(6).Value = pInvoice.DiscountUsd
            oPar(7).Value = pInvoice.ShippingUsd
            oPar(8).Value = pInvoice.TaxUsd
            oPar(9).Value = If(String.IsNullOrWhiteSpace(pInvoice.ClientEmail), DBNull.Value, CObj(pInvoice.ClientEmail))
            oPar(10).Value = If(String.IsNullOrWhiteSpace(pInvoice.Notes), DBNull.Value, CObj(pInvoice.Notes))

            OracleHelper.ExecuteNonQuery(Me.oTransacao, CommandType.StoredProcedure, "PACK_INVOICE.PROC_INSERT_INVOICE", oPar)

            vIdGerado = CType(oPar(11).Value, OracleDecimal)
            pInvoiceNumber = oPar(12).Value.ToString()

            Return CInt(vIdGerado.Value)

        Catch ex As Exception
            Throw New Exception(mensagemErroEstrutural(ex))
        End Try

    End Function

    Public Function incluirInvoiceItem(pItem As InvoiceItem) As Boolean

        Dim oPar(5) As OracleParameter

        oPar(0) = New OracleParameter("V_ID_INVOICE", OracleDbType.Int32, ParameterDirection.Input)
        oPar(1) = New OracleParameter("V_ID_EQUIPAMENTO", OracleDbType.Int32, ParameterDirection.Input)
        oPar(2) = New OracleParameter("V_DESCRIPTION", OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(3) = New OracleParameter("V_QTY", OracleDbType.Decimal, ParameterDirection.Input)
        oPar(4) = New OracleParameter("V_UNIT_PRICE_USD", OracleDbType.Decimal, ParameterDirection.Input)
        oPar(5) = New OracleParameter("V_NOTES", OracleDbType.Varchar2, ParameterDirection.Input)

        Try

            oPar(0).Value = pItem.IdInvoice
            oPar(1).Value = If(pItem.IdEquipamento.HasValue, CObj(pItem.IdEquipamento.Value), DBNull.Value)
            oPar(2).Value = pItem.Description
            oPar(3).Value = pItem.Qty
            oPar(4).Value = pItem.UnitPriceUsd
            oPar(5).Value = If(String.IsNullOrWhiteSpace(pItem.Notes), DBNull.Value, CObj(pItem.Notes))

            OracleHelper.ExecuteNonQuery(Me.oTransacao, CommandType.StoredProcedure, "PACK_INVOICE.PROC_INSERT_ITEM", oPar)
            Return True

        Catch ex As Exception
            Throw New Exception(mensagemErroEstrutural(ex))
        End Try

    End Function

    Public Function recalcularTotal(pIdInvoice As Integer) As Boolean

        Dim oPar(0) As OracleParameter
        oPar(0) = New OracleParameter("V_ID_INVOICE", OracleDbType.Int32, ParameterDirection.Input)

        Try
            oPar(0).Value = pIdInvoice
            OracleHelper.ExecuteNonQuery(Me.oTransacao, CommandType.StoredProcedure, "PACK_INVOICE.PROC_RECALC_TOTAL", oPar)
            Return True
        Catch ex As Exception
            Throw New Exception(mensagemErroEstrutural(ex))
        End Try

    End Function

    Public Function atualizarPaths(pIdInvoice As Integer, pLogoPath As String, pWordPath As String, pPdfPath As String) As Boolean

        Dim oPar(3) As OracleParameter

        oPar(0) = New OracleParameter("V_ID", OracleDbType.Int32, ParameterDirection.Input)
        oPar(1) = New OracleParameter("V_LOGO_PATH", OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(2) = New OracleParameter("V_WORD_FILE", OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(3) = New OracleParameter("V_PDF_FILE", OracleDbType.Varchar2, ParameterDirection.Input)

        Try
            oPar(0).Value = pIdInvoice
            oPar(1).Value = If(String.IsNullOrWhiteSpace(pLogoPath), DBNull.Value, CObj(pLogoPath))
            oPar(2).Value = If(String.IsNullOrWhiteSpace(pWordPath), DBNull.Value, CObj(pWordPath))
            oPar(3).Value = If(String.IsNullOrWhiteSpace(pPdfPath), DBNull.Value, CObj(pPdfPath))
            OracleHelper.ExecuteNonQuery(Me.oTransacao, CommandType.StoredProcedure, "PACK_INVOICE.PROC_UPDATE_PATHS", oPar)
            Return True
        Catch ex As Exception
            Throw New Exception(mensagemErroEstrutural(ex))
        End Try

    End Function

    Public Function atualizarStatus(pIdInvoice As Integer, pStatus As String) As Boolean

        Dim oPar(1) As OracleParameter

        oPar(0) = New OracleParameter("V_ID", OracleDbType.Int32, ParameterDirection.Input)
        oPar(1) = New OracleParameter("V_STATUS", OracleDbType.Varchar2, ParameterDirection.Input)

        Try
            oPar(0).Value = pIdInvoice
            oPar(1).Value = pStatus
            OracleHelper.ExecuteNonQuery(Me.oTransacao, CommandType.StoredProcedure, "PACK_INVOICE.PROC_UPDATE_STATUS", oPar)
            Return True
        Catch ex As Exception
            Throw New Exception(mensagemErroEstrutural(ex))
        End Try

    End Function

#End Region

#Region "Auxiliares"

    Private Function mensagemErroEstrutural(ex As Exception) As String

        Dim msg As String = ex.ToString().ToUpperInvariant()

        If msg.Contains("ORA-00942") OrElse
           msg.Contains("PLS-00201") OrElse
           msg.Contains("ORA-06550") Then
            Return "Estrutura de clientes/invoice ainda nao instalada no Oracle. Execute os scripts 07_DDL_CADASTROS_INVOICE.sql, 08_PACK_CLIENTE.sql e 09_PACK_INVOICE.sql."
        End If

        Return ex.Message

    End Function

#End Region

End Class
