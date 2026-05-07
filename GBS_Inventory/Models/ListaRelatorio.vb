Imports System.Data

''' <summary>
''' Lista de relatório em memória: permite montar um conjunto personalizado de
''' equipamentos independente dos filtros ativos no grid.
''' A lista sobrevive a trocas de aba mas é descartada ao fechar o form.
''' </summary>
Public Class ListaRelatorio

    Private ReadOnly _itens As New List(Of DataRow)
    Private ReadOnly _chave As String

    ''' <param name="pChave">Coluna usada como chave de deduplicação (padrão: INTERNAL_UID).</param>
    Public Sub New(Optional pChave As String = "INTERNAL_UID")
        _chave = pChave
    End Sub

    Public ReadOnly Property Itens As List(Of DataRow)
        Get
            Return _itens
        End Get
    End Property

    Public ReadOnly Property Count As Integer
        Get
            Return _itens.Count
        End Get
    End Property

    Public ReadOnly Property EstaAtiva As Boolean
        Get
            Return _itens.Count > 0
        End Get
    End Property

    ''' <summary>
    ''' Adiciona a linha se ainda não estiver na lista (deduplicação pela chave configurada).
    ''' Retorna True se adicionado, False se duplicata ignorada.
    ''' </summary>
    Public Function Adicionar(row As DataRow) As Boolean
        Dim uid As String = ObterChave(row)
        If Not String.IsNullOrEmpty(uid) Then
            If _itens.Any(Function(r) ObterChave(r) = uid) Then Return False
        End If
        _itens.Add(row)
        Return True
    End Function

    Public Function ContemChave(pValor As String) As Boolean
        Return _itens.Any(Function(r) ObterChave(r) = pValor)
    End Function

    Public Sub Limpar()
        _itens.Clear()
    End Sub

    Private Function ObterChave(row As DataRow) As String
        If row.Table.Columns.Contains(_chave) AndAlso Not IsDBNull(row(_chave)) Then
            Return row(_chave).ToString().Trim()
        End If
        Return ""
    End Function

End Class
