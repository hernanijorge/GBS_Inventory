Imports System.Data
Imports GBS_Inventory.Models

''' <summary>
''' Controller MVC — Importação de Planilha Excel
''' Coordena leitura da planilha e gravação em massa no Oracle.
''' Uma única transação para toda a importação — rollback em caso de erro.
''' </summary>
Public Class ImportacaoController

#Region "Atributos"

    Private oGravacao As clsGravacaoEquipamento
    Private oExcel    As clsImportacaoExcel

    Public Property TotalInseridos    As Integer = 0
    Public Property TotalAtualizados  As Integer = 0
    Public Property TotalErros        As Integer = 0
    Public Property TotalSkipped      As Integer = 0
    Public Property Erros             As New List(Of String)
    Public Property Skipped           As New List(Of String)
    Public Property UIDsImportados    As New List(Of String)
    Public Property BatteryCheckByUid As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

#End Region

#Region "Construtor"

    Public Sub New()

        oGravacao = New clsGravacaoEquipamento()
        oExcel    = New clsImportacaoExcel()

    End Sub

#End Region

#Region "Importação"

    ''' <summary>
    ''' Importa todos os equipamentos da planilha para o Oracle.
    ''' Chama PROC_UPSERT_EQUIPAMENTO para cada linha (INSERT se novo, UPDATE se UID já existir).
    ''' </summary>
    Public Function importarPlanilha(pCaminhoArquivo As String,
                                     Optional pCallbackProgresso As Action(Of Integer, Integer, String) = Nothing) _
                                     As Boolean

        Me.TotalInseridos    = 0
        Me.TotalAtualizados  = 0
        Me.TotalErros        = 0
        Me.TotalSkipped      = 0
        Me.Erros.Clear()
        Me.Skipped.Clear()
        Me.UIDsImportados.Clear()
        Me.BatteryCheckByUid.Clear()

        Dim listaEquipamentos As List(Of Equipamento)

        ' 1. Lê a planilha (pode falhar se o arquivo não existir ou estiver corrompido)
        Try

            listaEquipamentos = oExcel.lerPlanilha(pCaminhoArquivo)

        Catch ex As Exception

            Throw New Exception("Erro ao ler planilha: " & ex.Message)

        End Try

        If listaEquipamentos Is Nothing OrElse listaEquipamentos.Count = 0 Then
            Throw New Exception("Nenhum equipamento encontrado na planilha.")
        End If

        ' 2. Grava tudo em uma única transação
        Try

            oGravacao.beginTransacao()

            Dim vContador As Integer = 0
            Dim vTotal    As Integer = listaEquipamentos.Count

            For Each equip As Equipamento In listaEquipamentos

                vContador += 1

                ' Correction 1: skip rows with no INTERNAL_UID before calling DB
                If String.IsNullOrWhiteSpace(equip.InternalUID) Then
                    Me.TotalSkipped += 1
                    Me.Skipped.Add($"Row {vContador}: skipped — INTERNAL_UID is empty")
                    Continue For
                End If

                Try

                    Dim sResultado As String = oGravacao.upsertEquipamento(equip)

                    Select Case sResultado
                        Case "INSERTED"
                            Me.TotalInseridos += 1
                            Me.UIDsImportados.Add(equip.InternalUID)
                            Me.BatteryCheckByUid(equip.InternalUID) = equip.BatteryCheck
                        Case "UPDATED"
                            Me.TotalAtualizados += 1
                            Me.UIDsImportados.Add(equip.InternalUID)
                            Me.BatteryCheckByUid(equip.InternalUID) = equip.BatteryCheck
                        Case Else
                            Me.TotalErros += 1
                            Me.Erros.Add($"UID {equip.InternalUID}: {sResultado}")
                    End Select

                Catch ex As Exception

                    Me.TotalErros += 1
                    Me.Erros.Add($"UID {equip.InternalUID}: {ex.Message}")

                End Try

                ' Reporta progresso se callback fornecido
                If pCallbackProgresso IsNot Nothing AndAlso (vContador Mod 25 = 0 OrElse vContador = vTotal) Then
                    pCallbackProgresso(vContador, vTotal, equip.InternalUID)
                End If

            Next

            oGravacao.commitTransacao()

            Return True

        Catch ex As Exception

            oGravacao.rollbackTransacao()
            Throw New Exception("Erro na importação (rollback executado): " & ex.Message)

        End Try

    End Function

    Public ReadOnly Property TotalLinhasPlanilha As Integer
        Get
            Return oExcel.TotalLinhasLidas
        End Get
    End Property

    Public ReadOnly Property TotalAbas As Integer
        Get
            Return oExcel.TotalAbasProcessadas
        End Get
    End Property

#End Region

End Class
