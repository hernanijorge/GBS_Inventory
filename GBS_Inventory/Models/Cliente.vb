Namespace Models

    ''' <summary>
    ''' Modelo - Cliente
    ''' </summary>
    Public Class Cliente

        Public Property IdCliente As Integer
        Public Property NomeRazao As String
        Public Property NomeFantasia As String
        Public Property Documento As String
        Public Property Email As String
        Public Property Telefone As String
        Public Property Endereco1 As String
        Public Property Endereco2 As String
        Public Property Cidade As String
        Public Property Estado As String
        Public Property ZipCode As String
        Public Property Pais As String = "USA"
        Public Property Ativo As String = "Y"
        Public Property Observacoes As String
        Public Property DataCadastro As Date?
        Public Property DataAlteracao As Date?

        Public ReadOnly Property NomeExibicao As String
            Get
                If String.IsNullOrWhiteSpace(NomeFantasia) Then
                    Return NomeRazao
                End If

                Return NomeRazao & " (" & NomeFantasia & ")"
            End Get
        End Property

    End Class

End Namespace
