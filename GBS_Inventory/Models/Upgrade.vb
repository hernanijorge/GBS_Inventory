Namespace Models

    ''' <summary>
    ''' Modelo — Upgrade de componente (RAM, SSD, etc)
    ''' Representa a entidade TBL_EQUIPAMENTO_UPGRADE do Oracle
    ''' </summary>
    Public Class Upgrade

        Public Property IdUpgrade       As Integer
        Public Property IdEquipamento   As Integer
        Public Property InternalUID     As String

        ''' <summary>RAM, SSD, HDD, BATTERY, SCREEN, KEYBOARD, COVER, GPU, OTHER</summary>
        Public Property ComponentType   As String

        Public Property ValueBefore     As String
        Public Property ValueAfter      As String
        Public Property PartSerial      As String

        ''' <summary>REMESSA_EXTRA, NEW_PURCHASE, TRANSFERRED, WARRANTY, OTHER</summary>
        Public Property SourceOrigem    As String = "NEW_PURCHASE"

        Public Property CostUsd         As Decimal?
        Public Property Technician      As String
        Public Property Notes           As String
        Public Property DataUpgrade     As Date?

        ''' <summary>FK to TBL_COMPONENT — set when ACTION_TYPE is INSTALL or REMOVE</summary>
        Public Property IdComponent     As Integer?

        ''' <summary>INSTALL, REMOVE, or MANUAL (no stock tracking)</summary>
        Public Property ActionType      As String = "MANUAL"

        ''' <summary>For REMOVE action: IN_STOCK or SCRAPPED</summary>
        Public Property CompNewStatus   As String

        ''' <summary>Descrição amigável do componente</summary>
        Public ReadOnly Property DescricaoComponente As String
            Get
                Return If(String.IsNullOrEmpty(ComponentType), "",
                    Char.ToUpper(ComponentType(0)) & ComponentType.Substring(1).ToLower())
            End Get
        End Property

        ''' <summary>Descrição amigável da origem</summary>
        Public ReadOnly Property DescricaoOrigem As String
            Get
                Select Case SourceOrigem
                    Case "REMESSA_EXTRA" : Return "Extra de remessa"
                    Case "NEW_PURCHASE"  : Return "Compra nova"
                    Case "TRANSFERRED"   : Return "Transferido"
                    Case "WARRANTY"      : Return "Garantia"
                    Case "OTHER"         : Return "Outro"
                    Case Else            : Return If(String.IsNullOrEmpty(SourceOrigem), "—", SourceOrigem)
                End Select
            End Get
        End Property

    End Class

End Namespace
