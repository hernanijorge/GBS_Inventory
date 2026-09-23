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

    ' Report by Recipient data. Resumo gets two computed columns taken from the per-shipment
    ' rows: CLIENT (display name) and VARIATIONS (distinct spellings merged into the group).
    ' Resumo is Nothing when no shipment matches.
    Public Function buscarRelatorioPorDestinatario(pOpts As Models.ShipmentReportOptions) As (Resumo As DataTable, Detalhe As DataTable)

        Dim dsResumo As DataSet = oLeitura.selecionarResumoPorDestinatario(pOpts)
        If dsResumo Is Nothing OrElse dsResumo.Tables.Count = 0 OrElse dsResumo.Tables(0).Rows.Count = 0 Then Return (Nothing, Nothing)
        Dim resumo  As DataTable = dsResumo.Tables(0)
        Dim detalhe As DataTable = oLeitura.selecionarDetalhePorDestinatario(pOpts).Tables(0)

        Dim variacoesPorChave As Dictionary(Of String, List(Of String)) =
            detalhe.AsEnumerable().GroupBy(Function(r) r("CLIENT_KEY").ToString()) _
                   .ToDictionary(Function(g) g.Key,
                                 Function(g) g.Select(Function(r) r("RECIPIENT").ToString()).Distinct().OrderBy(Function(s) s, StringComparer.Ordinal).ToList())

        resumo.Columns.Add("CLIENT",     GetType(String))
        resumo.Columns.Add("VARIATIONS", GetType(String))
        For Each r As DataRow In resumo.Rows
            Dim chave As String = r("CLIENT_KEY").ToString()
            Dim variacoes As List(Of String) = If(variacoesPorChave.ContainsKey(chave), variacoesPorChave(chave), New List(Of String)())
            r("CLIENT")     = pOpts.NomeCliente(chave, variacoes)
            r("VARIATIONS") = String.Join(" | ", variacoes)
        Next

        ' Display names can differ from the SQL key (prefix mode), so re-sort on the name shown
        Dim dv As New DataView(resumo) With {.Sort = "TOTAL_SHIPMENTS DESC, CLIENT ASC"}
        Return (dv.ToTable(), detalhe)

    End Function

    Public Function buscarRemessasParaSelecao(pOpts As Models.ShipmentReportOptions) As DataSet
        Return oLeitura.selecionarRemessasParaSelecao(pOpts)
    End Function

    Public Function buscarPreviewAgrupamento(pOpts As Models.ShipmentReportOptions) As (Nomes As Integer, Grupos As Integer)
        Dim dt As DataTable = oLeitura.selecionarPreviewAgrupamento(pOpts).Tables(0)
        If dt.Rows.Count = 0 Then Return (0, 0)
        Return (Convert.ToInt32(dt.Rows(0)("NOMES")), Convert.ToInt32(dt.Rows(0)("GRUPOS")))
    End Function

    Public Function buscarDestinatarios() As DataSet
        Return oLeitura.selecionarDestinatarios()
    End Function

    ' Canonical name -> aliases
    Public Function buscarAliases() As Dictionary(Of String, List(Of String))
        Dim mapa As New Dictionary(Of String, List(Of String))(StringComparer.Ordinal)
        For Each r As DataRow In oLeitura.selecionarAliases().Tables(0).Rows
            Dim canonico As String = r("CANONICAL_NAME").ToString()
            If Not mapa.ContainsKey(canonico) Then mapa(canonico) = New List(Of String)()
            mapa(canonico).Add(r("ALIAS_NAME").ToString())
        Next
        Return mapa
    End Function

    Public Sub salvarAliases(pMapa As Dictionary(Of String, List(Of String)))
        Try
            oGravacao.beginTransacao()
            oGravacao.substituirAliases(pMapa, Environment.UserName)
            oGravacao.commitTransacao()
        Catch ex As Exception
            oGravacao.rollbackTransacao()
            Throw New Exception("Error saving client mapping: " & ex.Message)
        End Try
    End Sub

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
