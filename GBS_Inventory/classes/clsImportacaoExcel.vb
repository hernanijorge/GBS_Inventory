Imports System.Data
Imports System.Data.OleDb
Imports System.IO
Imports GBS_Inventory.Models

''' <summary>
''' Classe utilitária para leitura de planilhas Excel (.xlsx)
''' Percorre TODAS as abas da planilha e retorna uma lista unificada de Equipamentos.
''' Coluna "Internal UID" é a chave natural — equipamentos duplicados entre abas
''' serão tratados via PROC_UPSERT_EQUIPAMENTO.
''' </summary>
Public Class clsImportacaoExcel

#Region "Atributos"

    Public Property TotalLinhasLidas   As Integer = 0
    Public Property TotalLinhasVazias  As Integer = 0
    Public Property TotalAbasProcessadas As Integer = 0
    Public Property Erros              As New List(Of String)

#End Region

#Region "Leitura da Planilha"

    ''' <summary>
    ''' Lê a planilha completa e retorna uma lista de objetos Equipamento.
    ''' </summary>
    ''' <param name="pCaminhoArquivo">Caminho completo do arquivo .xlsx</param>
    Public Function lerPlanilha(pCaminhoArquivo As String) As List(Of Equipamento)

        If Not File.Exists(pCaminhoArquivo) Then
            Throw New FileNotFoundException("Planilha não encontrada: " & pCaminhoArquivo)
        End If

        Dim lista As New List(Of Equipamento)()

        ' Connection string para .xlsx via ACE.OLEDB 12.0
        Dim connStr As String = String.Format(
            "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1""",
            pCaminhoArquivo)

        Using oCon As New OleDbConnection(connStr)

            oCon.Open()

            ' Pega todas as abas (tabelas)
            Dim dtSchema As DataTable = oCon.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, New Object() {Nothing, Nothing, Nothing, "TABLE"})

            If dtSchema Is Nothing Then Return lista

            For Each rowAba As DataRow In dtSchema.Rows

                Dim sAba As String = rowAba("TABLE_NAME").ToString()

                ' Ignora áreas de impressão ou intervalos nomeados
                If sAba.EndsWith("_xlnm#_FilterDatabase") Then Continue For
                If sAba.StartsWith("'_") Then Continue For

                Try

                    Dim ds As DataSet = lerAba(oCon, sAba)
                    If ds IsNot Nothing AndAlso ds.Tables.Count > 0 Then

                        For Each row As DataRow In ds.Tables(0).Rows

                            Dim equipamento As Equipamento = mapearLinha(row, sAba)

                            If equipamento IsNot Nothing Then
                                lista.Add(equipamento)
                                Me.TotalLinhasLidas += 1
                            Else
                                Me.TotalLinhasVazias += 1
                            End If

                        Next

                        Me.TotalAbasProcessadas += 1

                    End If

                Catch ex As Exception

                    Me.Erros.Add("Erro na aba '" & sAba & "': " & ex.Message)

                End Try

            Next

        End Using

        Return lista

    End Function

#End Region

#Region "Leitura de uma aba"

    Private Function lerAba(oCon As OleDbConnection, pAba As String) As DataSet

        Dim sSQL As String = "SELECT * FROM [" & pAba.Replace("'", "") & "]"

        ' Garante bracket correto
        If Not pAba.EndsWith("$") And Not pAba.EndsWith("$'") Then
            sSQL = "SELECT * FROM [" & pAba.Trim("'"c) & "]"
        End If

        Dim oDa As New OleDbDataAdapter(sSQL, oCon)
        Dim ds  As New DataSet()
        oDa.Fill(ds)

        Return ds

    End Function

#End Region

#Region "Mapeamento da linha → Equipamento"

    ''' <summary>
    ''' Converte uma DataRow da planilha em objeto Equipamento.
    ''' Ignora linhas vazias ou cabeçalhos.
    ''' </summary>
    Private Function mapearLinha(pRow As DataRow, pAba As String) As Equipamento

        Dim vUID      As String = lerCelula(pRow, "Internal UID")
        Dim vSerial   As String = lerCelula(pRow, "Serial")
        Dim vMfr      As String = lerCelula(pRow, "Manufacturer")
        Dim vModel    As String = lerCelula(pRow, "Model")

        ' Linha vazia / cabeçalho
        If String.IsNullOrWhiteSpace(vUID) AndAlso String.IsNullOrWhiteSpace(vSerial) Then
            Return Nothing
        End If

        ' Linhas que são cabeçalho repetido dentro da aba
        If vUID.Trim().ToUpper() = "INTERNAL UID" Then Return Nothing

        Dim equipamento As New Equipamento()

        equipamento.InternalUID     = vUID.Trim()
        equipamento.SerialNumber    = vSerial.Trim().ToUpper()
        equipamento.Manufacturer    = vMfr.Trim().ToUpper()
        equipamento.Model           = vModel.Trim().ToUpper()
        equipamento.CpuFamily       = lerCelula(pRow, "CPU Family").Trim()
        equipamento.CpuModel        = lerCelula(pRow, "CPU Model").Trim()
        equipamento.CpuSpeedGhz     = parseDecimal(lerCelula(pRow, "CPU Speed"))
        equipamento.StorageGb       = lerCelula(pRow, "HDD Size").Trim()
        equipamento.RamGb           = lerCelula(pRow, "Memory(last#total)").Trim()
        equipamento.HardDriveType   = lerCelula(pRow, "Hard Drive Type").Trim()
        equipamento.Resolution      = lerCelula(pRow, "Resolution").Trim()
        equipamento.Graphics        = lerCelula(pRow, "Graphics").Trim()
        equipamento.ConditionStatus = normalizarCondicao(lerCelula(pRow, "Battery Condition"))
        equipamento.Notes           = lerCelula(pRow, "Notes").Trim()
        equipamento.SourceBatch     = pAba    ' o nome da aba vira o "lote"
        equipamento.DeviceType      = "LAPTOP"
        equipamento.Status          = "IN_STOCK"
        equipamento.IdEmpresa       = 1

        Return equipamento

    End Function

#End Region

#Region "Utilitários de parse"

    Private Function lerCelula(pRow As DataRow, pColuna As String) As String

        Try
            If pRow.Table.Columns.Contains(pColuna) Then
                Dim val As Object = pRow(pColuna)
                If val Is Nothing OrElse IsDBNull(val) Then Return ""
                Return val.ToString()
            End If
        Catch ex As Exception
            ' ignora
        End Try

        ' Tenta variações comuns (case-insensitive)
        For Each col As DataColumn In pRow.Table.Columns
            If String.Equals(col.ColumnName.Trim(), pColuna.Trim(), StringComparison.OrdinalIgnoreCase) Then
                Dim val As Object = pRow(col)
                If val Is Nothing OrElse IsDBNull(val) Then Return ""
                Return val.ToString()
            End If
        Next

        Return ""

    End Function

    Private Function parseDecimal(pValor As String) As Decimal?

        If String.IsNullOrWhiteSpace(pValor) Then Return Nothing

        ' Remove "GHz", "GB", espaços
        Dim sLimpo As String = pValor.ToUpper() _
                                     .Replace("GHZ", "") _
                                     .Replace("GB", "") _
                                     .Replace("MHZ", "") _
                                     .Trim()

        Dim vResult As Decimal
        If Decimal.TryParse(sLimpo.Replace(",", "."), Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, vResult) Then
            Return vResult
        End If

        Return Nothing

    End Function

    ''' <summary>Normaliza valores diversos para os enumerados aceitos pelo banco.</summary>
    Private Function normalizarCondicao(pValor As String) As String

        If String.IsNullOrWhiteSpace(pValor) Then Return "GOOD"

        Select Case pValor.Trim().ToUpper()
            Case "EXCELLENT", "E"              : Return "EXCELLENT"
            Case "GOOD", "G", "OK"             : Return "GOOD"
            Case "FAIR", "F", "AVG", "AVERAGE" : Return "FAIR"
            Case "POOR", "P", "BAD"            : Return "POOR"
            Case "NO BATTERY", "NB", "NONE"    : Return "NO_BATTERY"
            Case "Y", "YES"                    : Return "GOOD"
            Case "N", "NO", "N/A", ""          : Return "N/A"
            Case Else                          : Return "GOOD"
        End Select

    End Function

#End Region

End Class
