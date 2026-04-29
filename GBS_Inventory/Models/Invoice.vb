Namespace Models

    ''' <summary>
    ''' Modelo - Invoice (cabecalho)
    ''' </summary>
    Public Class Invoice

        Public Property IdInvoice As Integer
        Public Property InvoiceNumber As String
        Public Property IdCliente As Integer
        Public Property IdRemessa As Integer?
        Public Property IssueDate As Date = Date.Today
        Public Property DueDate As Date?
        Public Property Currency As String = "USD"
        Public Property SubtotalUsd As Decimal = 0D
        Public Property DiscountUsd As Decimal = 0D
        Public Property ShippingUsd As Decimal = 0D
        Public Property TaxUsd As Decimal = 0D
        Public Property TotalUsd As Decimal = 0D
        Public Property StatusInvoice As String = "DRAFT"
        Public Property ClientEmail As String
        Public Property Notes As String
        Public Property LogoPath As String
        Public Property WordFilePath As String
        Public Property PdfFilePath As String
        Public Property DataCadastro As Date?
        Public Property DataAlteracao As Date?

    End Class

    ''' <summary>
    ''' Modelo - Item da Invoice
    ''' </summary>
    Public Class InvoiceItem

        Public Property IdInvoiceItem As Integer
        Public Property IdInvoice As Integer
        Public Property IdEquipamento As Integer?
        Public Property Description As String
        Public Property Qty As Decimal = 1D
        Public Property UnitPriceUsd As Decimal = 0D
        Public Property LineTotalUsd As Decimal = 0D
        Public Property Notes As String

        Public Sub RecalcularTotal()
            LineTotalUsd = Qty * UnitPriceUsd
        End Sub

    End Class

End Namespace
