Namespace Models

    Public Class Component

        Public Property IdComponent     As Integer
        Public Property InternalUID     As String
        Public Property ComponentType   As String   ' RAM, SSD, HDD, MINI_DESKTOP
        Public Property CapacityGB      As Integer
        Public Property SpeedMhz        As Integer?  ' RAM only
        Public Property Generation      As String    ' DDR4, DDR5, NVMe, SATA
        Public Property Brand           As String
        Public Property PartNumber      As String
        Public Property Cpu             As String    ' MINI_DESKTOP only
        Public Property StorageGb       As Integer?  ' MINI_DESKTOP only
        Public Property ConditionStatus As String
        Public Property Status          As String
        Public Property SourceBatch     As String
        Public Property Notes           As String
        Public Property DateCreated     As DateTime?
        Public Property DateUpdated     As DateTime?

        Public ReadOnly Property DisplayLabel As String
            Get
                Dim cap As String = If(CapacityGB > 0, CapacityGB.ToString() & " GB", "")
                Dim spd As String = If(SpeedMhz.HasValue AndAlso SpeedMhz.Value > 0, SpeedMhz.Value.ToString() & " MHz", "")
                Dim gen As String = If(Not String.IsNullOrEmpty(Generation), Generation, "")
                Dim cpuVal As String = If(Not String.IsNullOrEmpty(Cpu), Cpu, "")
                Dim stg As String = If(StorageGb.HasValue AndAlso StorageGb.Value > 0, StorageGb.Value.ToString() & " GB Storage", "")
                Return String.Join(" ", {ComponentType, cap, gen, spd, cpuVal, stg}.Where(Function(s) Not String.IsNullOrEmpty(s)))
            End Get
        End Property

    End Class

End Namespace
