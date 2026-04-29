Imports System.Data
Imports GBS_Inventory.Models

''' <summary>
''' Controller MVC - Invoice
''' </summary>
Public Class InvoiceController

#Region "Atributos"

    Private oGravacao As clsGravacaoInvoice
    Private oLeitura As clsLeituraInvoice

#End Region

#Region "Construtor"

    Public Sub New()
        oGravacao = New clsGravacaoInvoice()
        oLeitura = New clsLeituraInvoice()
    End Sub

#End Region

#Region "Operacoes Transacionais"

    Public Function incluirComItens(pInvoice As Invoice,
                                    pItens As List(Of InvoiceItem),
                                    ByRef pInvoiceNumber As String) As Integer

        Dim vId As Integer = 0

        Try

            oGravacao.beginTransacao()

            vId = oGravacao.incluirInvoice(pInvoice, pInvoiceNumber)

            For Each item As InvoiceItem In pItens
                item.IdInvoice = vId
                item.RecalcularTotal()
                oGravacao.incluirInvoiceItem(item)
            Next

            oGravacao.recalcularTotal(vId)
            oGravacao.commitTransacao()

            Return vId

        Catch ex As Exception

            oGravacao.rollbackTransacao()
            Throw New Exception("Erro ao incluir invoice: " & ex.Message)

        End Try

    End Function

    Public Function atualizarPaths(pIdInvoice As Integer, pLogoPath As String, pWordPath As String, pPdfPath As String) As Boolean

        Try
            oGravacao.beginTransacao()
            oGravacao.atualizarPaths(pIdInvoice, pLogoPath, pWordPath, pPdfPath)
            oGravacao.commitTransacao()
            Return True
        Catch ex As Exception
            oGravacao.rollbackTransacao()
            Throw New Exception("Erro ao atualizar caminhos da invoice: " & ex.Message)
        End Try

    End Function

    Public Function atualizarStatus(pIdInvoice As Integer, pStatus As String) As Boolean

        Try
            oGravacao.beginTransacao()
            oGravacao.atualizarStatus(pIdInvoice, pStatus)
            oGravacao.commitTransacao()
            Return True
        Catch ex As Exception
            oGravacao.rollbackTransacao()
            Throw New Exception("Erro ao atualizar status da invoice: " & ex.Message)
        End Try

    End Function

#End Region

#Region "Consultas"

    Public Function buscarPorId(pIdInvoice As Integer) As DataSet
        Return oLeitura.selecionarInvoice(pIdInvoice)
    End Function

    Public Function buscarInvoices(pIdCliente As Integer?, pStatus As String) As DataSet
        Return oLeitura.selecionarInvoices(pIdCliente, pStatus)
    End Function

    Public Function buscarItens(pIdInvoice As Integer) As DataSet
        Return oLeitura.selecionarItens(pIdInvoice)
    End Function

    Public Function buscarDadosDocumento(pIdInvoice As Integer) As DataSet
        Return oLeitura.selecionarDadosDocumento(pIdInvoice)
    End Function

#End Region

End Class
