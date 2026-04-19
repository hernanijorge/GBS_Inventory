Namespace Models

    ''' <summary>
    ''' Modelo — Remessa (FedEx, UPS, USPS)
    ''' Representa a entidade TBL_REMESSA do Oracle
    ''' </summary>
    Public Class Remessa

#Region "Identificação"

        Public Property IdRemessa           As Integer
        Public Property RemessaRef          As String
        Public Property Direction           As String = "OUTBOUND"
        Public Property Carrier             As String
        Public Property TrackingNumber      As String
        Public Property StatusRemessa       As String = "LABEL_CREATED"
        Public Property StatusDescricao     As String

#End Region

#Region "Remetente"

        Public Property SenderName          As String
        Public Property SenderAddress       As String

#End Region

#Region "Destinatário"

        Public Property RecipientName       As String
        Public Property RecipientAddress    As String

#End Region

#Region "Valores e Datas"

        Public Property WeightLbs           As Decimal?
        Public Property ShippingCostUsd     As Decimal?
        Public Property InsuranceUsd        As Decimal?
        Public Property ServiceLevel        As String
        Public Property EstimatedDelivery   As Date?
        Public Property ActualDelivery      As Date?

#End Region

#Region "Auditoria"

        Public Property Notes               As String
        Public Property DataCadastro        As Date?
        Public Property DataAlteracao       As Date?
        Public Property TotalItens          As Integer

#End Region

#Region "Propriedades Computadas"

        Public ReadOnly Property UrlRastreio As String
            Get
                If String.IsNullOrEmpty(TrackingNumber) Then Return ""
                Select Case Carrier.ToUpper()
                    Case "FEDEX" : Return "https://www.fedex.com/fedextrack/?trknbr=" & TrackingNumber
                    Case "UPS"   : Return "https://www.ups.com/track?tracknum=" & TrackingNumber
                    Case "USPS"  : Return "https://tools.usps.com/go/TrackConfirmAction?tLabels=" & TrackingNumber
                    Case "DHL"   : Return "https://www.dhl.com/en/express/tracking.html?AWB=" & TrackingNumber
                    Case Else    : Return ""
                End Select
            End Get
        End Property

#End Region

    End Class

    ''' <summary>
    ''' Modelo — Item de Remessa
    ''' </summary>
    Public Class ItemRemessa

        Public Property IdRemessaItem       As Integer
        Public Property IdRemessa           As Integer
        Public Property IdEquipamento       As Integer
        Public Property InternalUID         As String
        Public Property Manufacturer        As String
        Public Property Model               As String
        Public Property ConditionAtShip     As String = "GOOD"
        Public Property SalePriceUsd        As Decimal?
        Public Property Notes               As String

    End Class

End Namespace
