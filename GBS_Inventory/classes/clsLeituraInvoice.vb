Imports Oracle.ManagedDataAccess.Client
Imports System.Configuration
Imports System.Data
Imports GBS_Inventory.OracleHelper

''' <summary>
''' Classe de leitura de Invoice
''' </summary>
Public Class clsLeituraInvoice

#Region "Atributos"

    Private ConnectionString As String

#End Region

#Region "Construtor"

    Public Sub New()

        Try
            Me.ConnectionString = ConfigurationManager.ConnectionStrings("OracleDB").ConnectionString
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Sub

#End Region

#Region "Consultas"

    Public Function selecionarInvoice(pIdInvoice As Integer) As DataSet

        Dim oPar(1) As OracleParameter

        oPar(0) = New OracleParameter("V_ID", OracleDbType.Int32, ParameterDirection.Input)
        oPar(1) = New OracleParameter("V_CURSOR", OracleDbType.RefCursor, ParameterDirection.Output)

        Try
            oPar(0).Value = pIdInvoice
            Return OracleHelper.ExecuteDataset(Me.ConnectionString, CommandType.StoredProcedure, "PACK_INVOICE.PROC_SELECT_INVOICE", oPar)
        Catch ex As Exception
            If IsErroObjetoInvoice(ex) Then
                Return selecionarInvoiceSemPackage(pIdInvoice)
            End If
            Throw New Exception(ex.Message)
        End Try

    End Function

    Public Function selecionarInvoices(pIdCliente As Integer?, pStatus As String) As DataSet

        Dim oPar(2) As OracleParameter

        oPar(0) = New OracleParameter("V_ID_CLIENTE", OracleDbType.Int32, ParameterDirection.Input)
        oPar(1) = New OracleParameter("V_STATUS", OracleDbType.Varchar2, ParameterDirection.Input)
        oPar(2) = New OracleParameter("V_CURSOR", OracleDbType.RefCursor, ParameterDirection.Output)

        Try
            oPar(0).Value = If(pIdCliente.HasValue, CObj(pIdCliente.Value), DBNull.Value)
            oPar(1).Value = If(String.IsNullOrWhiteSpace(pStatus), DBNull.Value, CObj(pStatus))
            Return OracleHelper.ExecuteDataset(Me.ConnectionString, CommandType.StoredProcedure, "PACK_INVOICE.PROC_SELECT_INVOICES", oPar)
        Catch ex As Exception
            If IsErroObjetoInvoice(ex) Then
                Return selecionarInvoicesSemPackage(pIdCliente, pStatus)
            End If
            Throw New Exception(ex.Message)
        End Try

    End Function

    Public Function selecionarItens(pIdInvoice As Integer) As DataSet

        Dim oPar(1) As OracleParameter

        oPar(0) = New OracleParameter("V_ID_INVOICE", OracleDbType.Int32, ParameterDirection.Input)
        oPar(1) = New OracleParameter("V_CURSOR", OracleDbType.RefCursor, ParameterDirection.Output)

        Try
            oPar(0).Value = pIdInvoice
            Return OracleHelper.ExecuteDataset(Me.ConnectionString, CommandType.StoredProcedure, "PACK_INVOICE.PROC_SELECT_ITENS", oPar)
        Catch ex As Exception
            If IsErroObjetoInvoice(ex) Then
                Return selecionarItensSemPackage(pIdInvoice)
            End If
            Throw New Exception(ex.Message)
        End Try

    End Function

    Public Function selecionarDadosDocumento(pIdInvoice As Integer) As DataSet

        Dim oPar(1) As OracleParameter

        oPar(0) = New OracleParameter("V_ID_INVOICE", OracleDbType.Int32, ParameterDirection.Input)
        oPar(1) = New OracleParameter("V_CURSOR", OracleDbType.RefCursor, ParameterDirection.Output)

        Try
            oPar(0).Value = pIdInvoice
            Return OracleHelper.ExecuteDataset(Me.ConnectionString, CommandType.StoredProcedure, "PACK_INVOICE.PROC_SELECT_DADOS_DOCUMENTO", oPar)
        Catch ex As Exception
            If IsErroObjetoInvoice(ex) Then
                Return selecionarDadosDocumentoSemPackage(pIdInvoice)
            End If
            Throw New Exception(ex.Message)
        End Try

    End Function

#End Region

#Region "Fallback sem package"

    Private Function selecionarInvoiceSemPackage(pIdInvoice As Integer) As DataSet

        Dim oPar(0) As OracleParameter
        oPar(0) = New OracleParameter("P_ID", OracleDbType.Int32, ParameterDirection.Input)
        oPar(0).Value = pIdInvoice

        Try
            Return OracleHelper.ExecuteDataset(Me.ConnectionString, CommandType.Text,
                                               "SELECT * FROM TBL_INVOICE WHERE ID_INVOICE = :P_ID", oPar)
        Catch ex As Exception
            Throw New Exception(mensagemErroEstrutural(ex))
        End Try

    End Function

    Private Function selecionarInvoicesSemPackage(pIdCliente As Integer?, pStatus As String) As DataSet

        Dim oPar(1) As OracleParameter
        oPar(0) = New OracleParameter("P_ID_CLIENTE", OracleDbType.Int32, ParameterDirection.Input)
        oPar(1) = New OracleParameter("P_STATUS", OracleDbType.Varchar2, ParameterDirection.Input)

        oPar(0).Value = If(pIdCliente.HasValue, CObj(pIdCliente.Value), DBNull.Value)
        oPar(1).Value = If(String.IsNullOrWhiteSpace(pStatus), DBNull.Value, CObj(pStatus))

        Try
            Return OracleHelper.ExecuteDataset(Me.ConnectionString, CommandType.Text,
                                               "SELECT I.ID_INVOICE, I.INVOICE_NUMBER, I.ID_CLIENTE, C.NOME_RAZAO AS CLIENTE_NOME, " &
                                               "       I.ID_REMESSA, I.ISSUE_DATE, I.DUE_DATE, I.SUBTOTAL_USD, I.DISCOUNT_USD, I.SHIPPING_USD, " &
                                               "       I.TAX_USD, I.TOTAL_USD, I.STATUS_INVOICE, I.CLIENT_EMAIL, I.DATA_CADASTRO " &
                                               "  FROM TBL_INVOICE I JOIN TBL_CLIENTE C ON I.ID_CLIENTE = C.ID_CLIENTE " &
                                               " WHERE (:P_ID_CLIENTE IS NULL OR I.ID_CLIENTE = :P_ID_CLIENTE) " &
                                               "   AND (:P_STATUS IS NULL OR I.STATUS_INVOICE = UPPER(:P_STATUS)) " &
                                               " ORDER BY I.DATA_CADASTRO DESC", oPar)
        Catch ex As Exception
            Throw New Exception(mensagemErroEstrutural(ex))
        End Try

    End Function

    Private Function selecionarItensSemPackage(pIdInvoice As Integer) As DataSet

        Dim oPar(0) As OracleParameter
        oPar(0) = New OracleParameter("P_ID_INVOICE", OracleDbType.Int32, ParameterDirection.Input)
        oPar(0).Value = pIdInvoice

        Try
            Return OracleHelper.ExecuteDataset(Me.ConnectionString, CommandType.Text,
                                               "SELECT II.ID_INVOICE_ITEM, II.ID_INVOICE, II.ID_EQUIPAMENTO, II.DESCRIPTION, " &
                                               "       II.QTY, II.UNIT_PRICE_USD, II.LINE_TOTAL_USD, II.NOTES, " &
                                               "       E.INTERNAL_UID, E.MARCA, E.MODEL " &
                                               "  FROM TBL_INVOICE_ITEM II " &
                                               "  LEFT JOIN TBL_EQUIPAMENTO E ON E.ID_EQUIPAMENTO = II.ID_EQUIPAMENTO " &
                                               " WHERE II.ID_INVOICE = :P_ID_INVOICE " &
                                               " ORDER BY II.ID_INVOICE_ITEM", oPar)
        Catch ex As Exception
            Throw New Exception(mensagemErroEstrutural(ex))
        End Try

    End Function

    Private Function selecionarDadosDocumentoSemPackage(pIdInvoice As Integer) As DataSet

        Dim oPar(0) As OracleParameter
        oPar(0) = New OracleParameter("P_ID_INVOICE", OracleDbType.Int32, ParameterDirection.Input)
        oPar(0).Value = pIdInvoice

        Try
            Return OracleHelper.ExecuteDataset(Me.ConnectionString, CommandType.Text,
                                               "SELECT I.ID_INVOICE, I.INVOICE_NUMBER, I.ISSUE_DATE, I.DUE_DATE, I.CURRENCY, " &
                                               "       I.SUBTOTAL_USD, I.DISCOUNT_USD, I.SHIPPING_USD, I.TAX_USD, I.TOTAL_USD, I.STATUS_INVOICE, " &
                                               "       I.CLIENT_EMAIL, I.NOTES, I.LOGO_PATH, I.WORD_FILE_PATH, I.PDF_FILE_PATH, I.ID_REMESSA, " &
                                               "       C.ID_CLIENTE, C.NOME_RAZAO, C.NOME_FANTASIA, C.DOCUMENTO, C.EMAIL AS EMAIL_CLIENTE, " &
                                               "       C.TELEFONE, C.ENDERECO1, C.ENDERECO2, C.CIDADE, C.ESTADO, C.ZIP_CODE, C.PAIS " &
                                               "  FROM TBL_INVOICE I JOIN TBL_CLIENTE C ON I.ID_CLIENTE = C.ID_CLIENTE " &
                                               " WHERE I.ID_INVOICE = :P_ID_INVOICE", oPar)
        Catch ex As Exception
            Throw New Exception(mensagemErroEstrutural(ex))
        End Try

    End Function

#End Region

#Region "Auxiliares"

    Private Function IsErroObjetoInvoice(ex As Exception) As Boolean

        Dim msg As String = ex.ToString().ToUpperInvariant()

        Return msg.Contains("PLS-00201") OrElse
               msg.Contains("ORA-06550") OrElse
               msg.Contains("PACK_INVOICE")

    End Function

    Private Function mensagemErroEstrutural(ex As Exception) As String

        Dim msg As String = ex.ToString().ToUpperInvariant()

        If msg.Contains("ORA-00942") Then
            Return "Estrutura de clientes/invoice ainda nao instalada no Oracle. Execute os scripts 07_DDL_CADASTROS_INVOICE.sql, 08_PACK_CLIENTE.sql e 09_PACK_INVOICE.sql."
        End If

        Return ex.Message

    End Function

#End Region

End Class
