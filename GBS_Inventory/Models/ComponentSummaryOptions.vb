Namespace Models

    ' Options for the Components Consolidated/Detailed summary report (frmReportOptions).
    ' Defaults reproduce the original consolidated report: no Model, every type, every status column.
    Public Class ComponentSummaryOptions

        Public Shared ReadOnly TiposConhecidos    As String() = {"DESKTOP", "MINI_DESKTOP", "RAM", "SSD", "HDD"}
        Public Shared ReadOnly TiposMaquinas      As String() = {"DESKTOP", "MINI_DESKTOP"}
        Public Shared ReadOnly TiposPecas         As String() = {"RAM", "SSD", "HDD"}
        Public Shared ReadOnly StatusDisponiveis  As String() = {"IN_STOCK", "INSTALLED", "SOLD", "SCRAPPED"}

        Public Property AgruparPorModel    As Boolean = False
        Public Property TiposIncluidos     As New List(Of String)(TiposConhecidos)
        Public Property IncluirOutrosTipos As Boolean = True    ' types outside TiposConhecidos
        Public Property ColunasStatus      As New List(Of String)(StatusDisponiveis)
        Public Property GeradoEm           As DateTime = DateTime.Now

        Public ReadOnly Property TodosOsTipos As Boolean
            Get
                Return IncluirOutrosTipos AndAlso TiposConhecidos.All(Function(t) TiposIncluidos.Contains(t))
            End Get
        End Property

        ' Grouping / display columns, in report order (Model sits between CPU and Storage)
        Public ReadOnly Property ColunasDescritivas As List(Of String)
            Get
                Dim cols As New List(Of String) From {"COMPONENT_TYPE", "CAPACITY_GB", "GENERATION", "SPEED_MHZ", "CPU"}
                If AgruparPorModel Then cols.Add("MODEL")
                cols.Add("STORAGE_GB")
                Return cols
            End Get
        End Property

        ' TOTAL + chosen status columns, always in StatusDisponiveis order
        Public ReadOnly Property ColunasNumericas As List(Of String)
            Get
                Dim cols As New List(Of String) From {"TOTAL"}
                cols.AddRange(StatusDisponiveis.Where(Function(s) ColunasStatus.Contains(s)))
                Return cols
            End Get
        End Property

        Public ReadOnly Property Titulo As String
            Get
                Return "GBS Components — " & If(AgruparPorModel, "Detailed", "Consolidated") & " Summary"
            End Get
        End Property

        Public Function Subtitulo(pQtdGrupos As Integer) As String
            Dim tipos As String
            If TodosOsTipos Then
                tipos = "ALL"
            Else
                tipos = String.Join(", ", TiposIncluidos)
                If IncluirOutrosTipos Then tipos &= If(tipos = "", "", ", ") & "other types"
            End If
            Return "Types: " & tipos &
                   " | Grouping: " & If(AgruparPorModel, "with Model", "without Model") &
                   " | Generated: " & GeradoEm.ToString("yyyy-MM-dd HH:mm:ss") &
                   " | Groups: " & pQtdGrupos.ToString()
        End Function

        ' e.g. ComponentSummary_20260923_152500_DETAILED_ALL
        Public Function NomeArquivoBase() As String
            Dim escopo As String
            If TodosOsTipos Then
                escopo = "ALL"
            ElseIf Not IncluirOutrosTipos AndAlso MesmosTipos(TiposMaquinas) Then
                escopo = "DESKTOPS_ONLY"
            ElseIf Not IncluirOutrosTipos AndAlso MesmosTipos(TiposPecas) Then
                escopo = "PARTS_ONLY"
            Else
                escopo = "CUSTOM"
            End If
            Return "ComponentSummary_" & GeradoEm.ToString("yyyyMMdd_HHmmss") & "_" &
                   If(AgruparPorModel, "DETAILED", "CONSOLIDATED") & "_" & escopo
        End Function

        Private Function MesmosTipos(pTipos As String()) As Boolean
            Return TiposIncluidos.Count = pTipos.Length AndAlso pTipos.All(Function(t) TiposIncluidos.Contains(t))
        End Function

    End Class

End Namespace
