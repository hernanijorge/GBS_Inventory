Imports System.Data
Imports GBS_Inventory.Models

''' <summary>
''' Controller MVC — Upgrades
''' </summary>
Public Class UpgradeController

#Region "Atributos"

    Private oGravacao As clsGravacaoUpgrade
    Private oLeitura  As clsLeituraUpgrade

#End Region

#Region "Construtor"

    Public Sub New()

        oGravacao = New clsGravacaoUpgrade()
        oLeitura  = New clsLeituraUpgrade()

    End Sub

#End Region

#Region "Operações Transacionais"

    Public Function incluir(pUpgrade As Upgrade) As Integer

        Dim vIdGerado As Integer

        Try

            oGravacao.beginTransacao()

            vIdGerado = oGravacao.incluirUpgrade(pUpgrade)

            oGravacao.commitTransacao()

            Return vIdGerado

        Catch ex As Exception

            oGravacao.rollbackTransacao()
            Throw New Exception("Erro ao incluir upgrade: " & ex.Message)

        End Try

    End Function

    Public Function excluir(pIdUpgrade As Integer) As Boolean

        Try

            oGravacao.beginTransacao()
            oGravacao.excluirUpgrade(pIdUpgrade)
            oGravacao.commitTransacao()

            Return True

        Catch ex As Exception

            oGravacao.rollbackTransacao()
            Throw New Exception("Erro ao excluir upgrade: " & ex.Message)

        End Try

    End Function

#End Region

#Region "Consultas"

    Public Function buscarPorEquipamento(pIdEquipamento As Integer) As DataSet
        Return oLeitura.selecionarUpgradesEquipamento(pIdEquipamento)
    End Function

    Public Function buscarRecentes(pDias As Integer) As DataSet
        Return oLeitura.selecionarUpgradesRecentes(pDias)
    End Function

    Public Function buscarComCliente(pDias As Integer) As DataSet
        Return oLeitura.selecionarUpgradesComCliente(pDias)
    End Function

    Public Function buscarOrigensDistintas() As DataSet
        Return oLeitura.selecionarOrigensDistintas()
    End Function

    Public Function buscarComponentesParaInstalar(pType As String) As DataSet
        Return oLeitura.selectComponentesParaInstalar(pType)
    End Function

    Public Function buscarComponentesInstalados(pType As String) As DataSet
        Return oLeitura.selectComponentesInstalados(pType)
    End Function

#End Region

End Class
