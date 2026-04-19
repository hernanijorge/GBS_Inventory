Imports System.Data
Imports GBS_Inventory.Models

''' <summary>
''' Controller MVC — Equipamento
''' Coordena Gravação + Leitura, centraliza transações
''' </summary>
Public Class EquipamentoController

#Region "Atributos"

    Private oGravacao As clsGravacaoEquipamento
    Private oLeitura  As clsLeituraEquipamento

#End Region

#Region "Construtor"

    Public Sub New()

        oGravacao = New clsGravacaoEquipamento()
        oLeitura  = New clsLeituraEquipamento()

    End Sub

#End Region

#Region "Operações Transacionais"

    Public Function incluir(pEquipamento As Equipamento) As Integer

        Dim vIdGerado As Integer

        Try

            oGravacao.beginTransacao()

            vIdGerado = oGravacao.incluirEquipamento(pEquipamento)

            oGravacao.commitTransacao()

            Return vIdGerado

        Catch ex As Exception

            oGravacao.rollbackTransacao()
            Throw New Exception("Erro ao incluir equipamento: " & ex.Message)

        End Try

    End Function

    Public Function alterar(pEquipamento As Equipamento) As Boolean

        Try

            oGravacao.beginTransacao()

            oGravacao.alterarEquipamento(pEquipamento)

            oGravacao.commitTransacao()

            Return True

        Catch ex As Exception

            oGravacao.rollbackTransacao()
            Throw New Exception("Erro ao alterar equipamento: " & ex.Message)

        End Try

    End Function

    Public Function atualizarStatus(pIdEquipamento As Integer, pStatus As String) As Boolean

        Try

            oGravacao.beginTransacao()

            oGravacao.atualizarStatus(pIdEquipamento, pStatus)

            oGravacao.commitTransacao()

            Return True

        Catch ex As Exception

            oGravacao.rollbackTransacao()
            Throw New Exception("Erro ao atualizar status: " & ex.Message)

        End Try

    End Function

    Public Function excluir(pIdEquipamento As Integer) As Boolean

        Try

            oGravacao.beginTransacao()

            oGravacao.excluirEquipamento(pIdEquipamento)

            oGravacao.commitTransacao()

            Return True

        Catch ex As Exception

            oGravacao.rollbackTransacao()
            Throw New Exception("Erro ao excluir equipamento: " & ex.Message)

        End Try

    End Function

#End Region

#Region "Consultas"

    Public Function buscarPorId(pIdEquipamento As Integer) As DataSet
        Return oLeitura.selecionarPorId(pIdEquipamento)
    End Function

    Public Function buscarPorUID(pInternalUID As String) As DataSet
        Return oLeitura.selecionarPorUID(pInternalUID)
    End Function

    Public Function buscarPorSerial(pSerial As String) As DataSet
        Return oLeitura.selecionarPorSerial(pSerial)
    End Function

    Public Function buscarTodos(pIdEmpresa As Integer) As DataSet
        Return oLeitura.selecionarTodos(pIdEmpresa)
    End Function

    Public Function buscarPorFiltro(pPesquisa As String, pManufacturer As String, pCondition As String, pStatus As String) As DataSet
        Return oLeitura.selecionarPorFiltro(pPesquisa, pManufacturer, pCondition, pStatus)
    End Function

    Public Function obterDashboardTotais(pIdEmpresa As Integer) As DataSet
        Return oLeitura.obterDashboardTotais(pIdEmpresa)
    End Function

    Public Function obterResumoWhatsApp(pIdEmpresa As Integer) As DataSet
        Return oLeitura.obterResumoWhatsApp(pIdEmpresa)
    End Function

#End Region

End Class
