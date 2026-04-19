Namespace Models

    ''' <summary>
    ''' Modelo — Equipamento (laptop/desktop)
    ''' Representa a entidade TBL_EQUIPAMENTO do Oracle
    ''' </summary>
    Public Class Equipamento

#Region "Identificação"

        Public Property IdEquipamento       As Integer
        Public Property InternalUID         As String
        Public Property Manufacturer        As String
        Public Property Model               As String
        Public Property SerialNumber        As String

#End Region

#Region "Especificações Técnicas"

        Public Property CpuFamily           As String
        Public Property CpuModel            As String
        Public Property CpuSpeedGhz         As Decimal?
        Public Property RamGb               As String
        Public Property StorageGb           As String
        Public Property HardDriveType       As String
        Public Property Resolution          As String
        Public Property Graphics            As String

#End Region

#Region "Classificação"

        Public Property DeviceType          As String = "LAPTOP"
        Public Property ConditionStatus     As String = "GOOD"
        Public Property Status              As String = "IN_STOCK"
        Public Property StatusDescricao     As String
        Public Property Notes               As String
        Public Property SourceBatch         As String

#End Region

#Region "Auditoria"

        Public Property IdEmpresa           As Integer = 1
        Public Property DataCadastro        As Date?
        Public Property DataAlteracao       As Date?
        Public Property IdUsuarioCadastro   As Integer?

#End Region

#Region "Propriedades Computadas"

        Public ReadOnly Property DescricaoCompleta As String
            Get
                Dim sDesc As String = Manufacturer & " " & Model
                If Not String.IsNullOrEmpty(CpuModel) Then sDesc &= " · " & CpuModel
                If Not String.IsNullOrEmpty(RamGb) Then sDesc &= " · " & RamGb
                If Not String.IsNullOrEmpty(StorageGb) Then sDesc &= " · " & StorageGb
                Return sDesc
            End Get
        End Property

        Public ReadOnly Property CondicaoBoa As Boolean
            Get
                Return ConditionStatus = "EXCELLENT" Or ConditionStatus = "GOOD"
            End Get
        End Property

#End Region

    End Class

End Namespace
