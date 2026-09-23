Namespace Models

    ' Options for the Shipments "Report by Recipient" (frmShipmentReportOptions).
    ' TBL_REMESSA has no PENDING/CANCELLED statuses (Cancel Shipment deletes the row), so the
    ' dialog filters by status *groups* mapped onto the real STATUS_REMESSA values below.
    Public Class ShipmentReportOptions

        Public Shared ReadOnly GruposStatus As New List(Of (Chave As String, Rotulo As String, Status As String())) From {
            ("DELIVERED",  "Delivered",  {"DELIVERED"}),
            ("IN_TRANSIT", "In Transit", {"PICKED_UP", "IN_TRANSIT", "OUT_FOR_DELIVERY"}),
            ("PENDING",    "Pending",    {"LABEL_CREATED", "OPEN"}),
            ("ISSUES",     "Issues",     {"DELAYED", "EXCEPTION", "RETURNED"})
        }

        ' Client grouping modes
        Public Const AGRUP_NENHUM  As String = "NONE"     ' each recipient spelling separate
        Public Const AGRUP_PREFIXO As String = "PREFIX"   ' same first N characters
        Public Const AGRUP_ALIAS   As String = "ALIAS"    ' TBL_CLIENT_ALIAS mapping

        Public Const PREFIXO_MIN     As Integer = 3
        Public Const PREFIXO_MAX     As Integer = 15
        Public Const PREFIXO_PADRAO  As Integer = 6

        Public Property DataInicio      As Date = Date.Today.AddDays(-90)
        Public Property DataFim         As Date = Date.Today
        Public Property GruposIncluidos As New List(Of String)(GruposStatus.Select(Function(g) g.Chave))
        Public Property RecipientFiltro As String = Nothing     ' Nothing = all recipients
        Public Property Detalhado       As Boolean = False
        Public Property Agrupamento     As String = AGRUP_PREFIXO
        Public Property PrefixLen       As Integer = PREFIXO_PADRAO
        Public Property IdsSelecionados As List(Of Integer) = Nothing   ' Nothing = every matching shipment
        Public Property GeradoEm        As DateTime = DateTime.Now

        ' Real STATUS_REMESSA values for the selected groups
        Public ReadOnly Property StatusIncluidos As List(Of String)
            Get
                Return GruposStatus.Where(Function(g) GruposIncluidos.Contains(g.Chave)).SelectMany(Function(g) g.Status).ToList()
            End Get
        End Property

        Public ReadOnly Property RotulosIncluidos As String
            Get
                Return String.Join(", ", GruposStatus.Where(Function(g) GruposIncluidos.Contains(g.Chave)).Select(Function(g) g.Rotulo))
            End Get
        End Property

        Public ReadOnly Property Agrupado As Boolean
            Get
                Return Agrupamento <> AGRUP_NENHUM
            End Get
        End Property

        Public ReadOnly Property DescricaoAgrupamento As String
            Get
                Select Case Agrupamento
                    Case AGRUP_PREFIXO : Return "prefix (" & PrefixLen.ToString() & " chars)"
                    Case AGRUP_ALIAS   : Return "custom mapping"
                    Case Else          : Return "none"
                End Select
            End Get
        End Property

        ' Name shown in the Client column. Prefix groups have no canonical name, so the
        ' shortest spelling in the group is used (COMPUGANGAS over COMPUGANGAS HP).
        Public Function NomeCliente(pChave As String, pVariacoes As IEnumerable(Of String)) As String
            If Agrupamento <> AGRUP_PREFIXO OrElse Not pVariacoes.Any() Then Return pChave
            Return pVariacoes.OrderBy(Function(v) v.Length).ThenBy(Function(v) v, StringComparer.Ordinal).First()
        End Function

        Public Function NomeArquivoBase() As String
            Return "ShipmentReport_ByRecipient_" & GeradoEm.ToString("yyyyMMdd_HHmmss") & "_" &
                   If(Detalhado, "DETAILED", "SUMMARY")
        End Function

        Public Function Clonar() As ShipmentReportOptions
            Dim c As ShipmentReportOptions = DirectCast(Me.MemberwiseClone(), ShipmentReportOptions)
            c.GruposIncluidos = New List(Of String)(GruposIncluidos)
            c.IdsSelecionados = If(IdsSelecionados Is Nothing, Nothing, New List(Of Integer)(IdsSelecionados))
            Return c
        End Function

    End Class

End Namespace
