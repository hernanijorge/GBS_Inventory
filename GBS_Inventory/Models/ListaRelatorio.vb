Imports System.Data

''' <summary>
''' Lista de relatório em memória: permite montar um conjunto personalizado de
''' equipamentos independente dos filtros ativos no grid.
''' A lista sobrevive a trocas de aba mas é descartada ao fechar o form.
''' </summary>
Public Class ListaRelatorio

    Private ReadOnly _itens As New List(Of DataRow)

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
    ''' Adiciona a linha se ainda não estiver na lista (chave: INTERNAL_UID).
    ''' Retorna True se adicionado, False se duplicata ignorada.
    ''' </summary>
    Public Function Adicionar(row As DataRow) As Boolean
        Dim uid As String = ObterUID(row)
        If Not String.IsNullOrEmpty(uid) Then
            If _itens.Any(Function(r) ObterUID(r) = uid) Then Return False
        End If
        _itens.Add(row)
        Return True
    End Function

    Public Sub Limpar()
        _itens.Clear()
    End Sub

    Private Shared Function ObterUID(row As DataRow) As String
        For Each col As String In {"INTERNAL_UID", "ID"}
            If row.Table.Columns.Contains(col) AndAlso Not IsDBNull(row(col)) Then
                Return row(col).ToString().Trim()
            End If
        Next
        Return ""
    End Function

End Class
