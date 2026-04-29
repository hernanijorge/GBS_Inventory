Imports System.Data
Imports GBS_Inventory.Models

''' <summary>
''' Controller MVC - Cliente
''' </summary>
Public Class ClienteController

#Region "Atributos"

    Private oGravacao As clsGravacaoCliente
    Private oLeitura As clsLeituraCliente

#End Region

#Region "Construtor"

    Public Sub New()
        oGravacao = New clsGravacaoCliente()
        oLeitura = New clsLeituraCliente()
    End Sub

#End Region

#Region "Operacoes Transacionais"

    Public Function incluir(pCliente As Cliente) As Integer

        Try
            oGravacao.beginTransacao()
            Dim vId As Integer = oGravacao.incluirCliente(pCliente)
            oGravacao.commitTransacao()
            Return vId
        Catch ex As Exception
            oGravacao.rollbackTransacao()
            Throw New Exception("Erro ao incluir cliente: " & ex.Message)
        End Try

    End Function

    Public Function alterar(pCliente As Cliente) As Boolean

        Try
            oGravacao.beginTransacao()
            oGravacao.alterarCliente(pCliente)
            oGravacao.commitTransacao()
            Return True
        Catch ex As Exception
            oGravacao.rollbackTransacao()
            Throw New Exception("Erro ao alterar cliente: " & ex.Message)
        End Try

    End Function

    Public Function excluir(pIdCliente As Integer) As Boolean

        Try
            oGravacao.beginTransacao()
            oGravacao.excluirCliente(pIdCliente)
            oGravacao.commitTransacao()
            Return True
        Catch ex As Exception
            oGravacao.rollbackTransacao()
            Throw New Exception("Erro ao excluir cliente: " & ex.Message)
        End Try

    End Function

#End Region

#Region "Consultas"

    Public Function buscarPorId(pIdCliente As Integer) As DataSet
        Return oLeitura.selecionarCliente(pIdCliente)
    End Function

    Public Function buscarTodos(pAtivo As String) As DataSet
        Return oLeitura.selecionarClientes(pAtivo)
    End Function

    Public Function buscarPorFiltro(pPesquisa As String) As DataSet
        Return oLeitura.selecionarClientesFiltro(pPesquisa)
    End Function

#End Region

End Class
