Imports System.Data
Imports GBS_Inventory.Models

''' <summary>
''' Controller MVC — Remessas
''' </summary>
Public Class RemessaController

#Region "Atributos"

    Private oGravacao As clsGravacaoRemessa
    Private oLeitura  As clsLeituraRemessa

#End Region

#Region "Construtor"

    Public Sub New()

        oGravacao = New clsGravacaoRemessa()
        oLeitura  = New clsLeituraRemessa()

    End Sub

#End Region

#Region "Operações Transacionais"

    ''' <summary>
    ''' Cria uma nova remessa com todos os itens em uma única transação.
    ''' Atualiza status dos equipamentos para SHIPPED automaticamente.
    ''' </summary>
    Public Function incluirRemessa(pRemessa As Remessa, pItens As List(Of ItemRemessa), ByRef pRemessaRef As String) As Integer

        Dim vIdGerado As Integer

        Try

            oGravacao.beginTransacao()

            vIdGerado = oGravacao.incluirRemessa(pRemessa, pRemessaRef)

            For Each item As ItemRemessa In pItens
                item.IdRemessa = vIdGerado
                oGravacao.incluirItemRemessa(item)
            Next

            oGravacao.commitTransacao()

            Return vIdGerado

        Catch ex As Exception

            oGravacao.rollbackTransacao()
            Throw New Exception("Erro ao incluir remessa: " & ex.Message)

        End Try

    End Function

    Public Function atualizarStatus(pIdRemessa As Integer, pStatus As String, pDataEntrega As Date?) As Boolean

        Try

            oGravacao.beginTransacao()
            oGravacao.atualizarStatusRemessa(pIdRemessa, pStatus, pDataEntrega)
            oGravacao.commitTransacao()

            Return True

        Catch ex As Exception

            oGravacao.rollbackTransacao()
            Throw New Exception("Erro ao atualizar status da remessa: " & ex.Message)

        End Try

    End Function

    Public Function excluir(pIdRemessa As Integer) As Boolean

        Try

            oGravacao.beginTransacao()
            oGravacao.excluirRemessa(pIdRemessa)
            oGravacao.commitTransacao()

            Return True

        Catch ex As Exception

            oGravacao.rollbackTransacao()
            Throw New Exception("Erro ao excluir remessa: " & ex.Message)

        End Try

    End Function

    Public Function cancelar(pIdRemessa As Integer) As Boolean

        Try

            oGravacao.beginTransacao()
            oGravacao.cancelarRemessa(pIdRemessa)
            oGravacao.commitTransacao()

            Return True

        Catch ex As Exception

            oGravacao.rollbackTransacao()
            Throw New Exception("Erro ao cancelar remessa: " & ex.Message)

        End Try

    End Function

#End Region

#Region "Consultas"

    Public Function buscarRemessasAtivas() As DataSet
        Return oLeitura.selecionarRemessasAtivas()
    End Function

    Public Function buscarRemessa(pIdRemessa As Integer) As DataSet
        Return oLeitura.selecionarRemessa(pIdRemessa)
    End Function

    Public Function buscarItensRemessa(pIdRemessa As Integer) As DataSet
        Return oLeitura.selecionarItensRemessa(pIdRemessa)
    End Function

    Public Function buscarEquipamentosDisponiveis(pPesquisa As String) As DataSet
        Return oLeitura.selecionarEquipamentosDisponiveis(pPesquisa)
    End Function

#End Region

End Class
