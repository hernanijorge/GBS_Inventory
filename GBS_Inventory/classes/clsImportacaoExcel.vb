Imports System.Data
Imports System.IO
Imports GBS_Inventory.Models
Imports OfficeOpenXml

''' <summary>
''' Classe utilitaria para leitura de planilhas Excel (.xlsx).
''' Percorre todas as abas da planilha e retorna uma lista unificada de Equipamentos.
''' </summary>
Public Class clsImportacaoExcel

#Region "Atributos"

    Public Property TotalLinhasLidas As Integer = 0
    Public Property TotalLinhasVazias As Integer = 0
    Public Property TotalAbasProcessadas As Integer = 0
    Public Property Erros As New List(Of String)

#End Region

#Region "Leitura da Planilha"

    ''' <summary>
    ''' Le a planilha completa e retorna uma lista de objetos Equipamento.
    ''' </summary>
    ''' <param name="pCaminhoArquivo">Caminho completo do arquivo .xlsx</param>
    Public Function lerPlanilha(pCaminhoArquivo As String) As List(Of Equipamento)

        If Not File.Exists(pCaminhoArquivo) Then
            Throw New FileNotFoundException("Planilha nao encontrada: " & pCaminhoArquivo)
        End If

        Dim lista As New List(Of Equipamento)()

        Using pkg As New ExcelPackage(New FileInfo(pCaminhoArquivo))

            For Each ws As ExcelWorksheet In pkg.Workbook.Worksheets

                Dim sAba As String = ws.Name

                Try

                    Dim ds As DataSet = lerAba(ws)
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

    Private Function lerAba(pWs As ExcelWorksheet) As DataSet

        Dim ds As New DataSet()
        Dim dt As New DataTable(pWs.Name)

        If pWs.Dimension Is Nothing Then
            ds.Tables.Add(dt)
            Return ds
        End If

        Dim headerRow As Integer = pWs.Dimension.Start.Row
        Dim firstCol As Integer = pWs.Dimension.Start.Column
        Dim lastCol As Integer = pWs.Dimension.End.Column
        Dim lastRow As Integer = pWs.Dimension.End.Row

        For col As Integer = firstCol To lastCol
            Dim nomeColuna As String = Convert.ToString(pWs.Cells(headerRow, col).Value).Trim()
            If String.IsNullOrWhiteSpace(nomeColuna) Then nomeColuna = "Column" & col.ToString()

            Dim nomeOriginal As String = nomeColuna
            Dim contador As Integer = 1

            While dt.Columns.Contains(nomeColuna)
                contador += 1
                nomeColuna = nomeOriginal & "_" & contador.ToString()
            End While

            dt.Columns.Add(nomeColuna)
        Next

        For row As Integer = headerRow + 1 To lastRow
            Dim dr As DataRow = dt.NewRow()
            Dim temValor As Boolean = False

            For col As Integer = firstCol To lastCol
                Dim valor As Object = pWs.Cells(row, col).Value
                If valor IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(valor.ToString()) Then
                    temValor = True
                End If

                dr(col - firstCol) = If(valor Is Nothing, "", valor.ToString())
            Next

            If temValor Then dt.Rows.Add(dr)
        Next

        ds.Tables.Add(dt)

        Return ds

    End Function

#End Region

#Region "Mapeamento da linha para Equipamento"

    ''' <summary>
    ''' Converte uma DataRow da planilha em objeto Equipamento.
    ''' Ignora linhas vazias ou cabecalhos.
    ''' </summary>
    Private Function mapearLinha(pRow As DataRow, pAba As String) As Equipamento

        Dim vUID As String = lerCelula(pRow, "Internal UID")
        Dim vSerial As String = lerCelula(pRow, "Serial")
        Dim vMfr As String = lerCelula(pRow, "Manufacturer")
        Dim vModel As String = lerCelula(pRow, "Model")
        Dim vBattery As String = lerPrimeiraCelula(pRow,
                                                   "Battery Condition",
                                                   "Battery",
                                                   "Battery Status",
                                                   "Battery Health",
                                                   "Bateria",
                                                   "Condicao Bateria")

        If String.IsNullOrWhiteSpace(vUID) AndAlso String.IsNullOrWhiteSpace(vSerial) Then
            Return Nothing
        End If

        If vUID.Trim().ToUpper() = "INTERNAL UID" Then Return Nothing

        Dim equipamento As New Equipamento()

        equipamento.InternalUID = vUID.Trim()
        equipamento.SerialNumber = vSerial.Trim().ToUpper()
        equipamento.Manufacturer = vMfr.Trim().ToUpper()
        equipamento.Model = vModel.Trim().ToUpper()
        equipamento.CpuFamily = lerCelula(pRow, "CPU Family").Trim()
        equipamento.CpuModel = lerCelula(pRow, "CPU Model").Trim()
        equipamento.CpuSpeedGhz = parseDecimal(lerCelula(pRow, "CPU Speed"))
        equipamento.StorageGb = lerCelula(pRow, "HDD Size").Trim()
        equipamento.RamGb = lerCelula(pRow, "Memory(last#total)").Trim()
        equipamento.HardDriveType = lerCelula(pRow, "Hard Drive Type").Trim()
        equipamento.Resolution = lerCelula(pRow, "Resolution").Trim()
        equipamento.Graphics = lerCelula(pRow, "Graphics").Trim()
        equipamento.ConditionStatus = normalizarCondicao(vBattery)
        equipamento.BatteryCheck = vBattery.Trim()
        equipamento.Notes = lerCelula(pRow, "Notes").Trim()
        equipamento.SourceBatch = pAba
        equipamento.DeviceType = "LAPTOP"
        equipamento.Status = "IN_STOCK"
        equipamento.IdEmpresa = 1

        Return equipamento

    End Function

#End Region

#Region "Utilitarios de parse"

    Private Function lerCelula(pRow As DataRow, pColuna As String) As String

        Try
            If pRow.Table.Columns.Contains(pColuna) Then
                Dim val As Object = pRow(pColuna)
                If val Is Nothing OrElse IsDBNull(val) Then Return ""
                Return val.ToString()
            End If
        Catch ex As Exception
            ' Ignora e tenta a busca case-insensitive abaixo.
        End Try

        For Each col As DataColumn In pRow.Table.Columns
            If String.Equals(col.ColumnName.Trim(), pColuna.Trim(), StringComparison.OrdinalIgnoreCase) Then
                Dim val As Object = pRow(col)
                If val Is Nothing OrElse IsDBNull(val) Then Return ""
                Return val.ToString()
            End If
        Next

        Return ""

    End Function

    Private Function lerPrimeiraCelula(pRow As DataRow, ParamArray pColunas() As String) As String

        For Each coluna As String In pColunas
            Dim valor As String = lerCelula(pRow, coluna).Trim()
            If Not String.IsNullOrWhiteSpace(valor) Then Return valor
        Next

        Return ""

    End Function

    Private Function parseDecimal(pValor As String) As Decimal?

        If String.IsNullOrWhiteSpace(pValor) Then Return Nothing

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
            Case "EXCELLENT", "E"
                Return "EXCELLENT"
            Case "GOOD", "G", "OK", "NORMAL"
                Return "GOOD"
            Case "FAIR", "F", "AVG", "AVERAGE"
                Return "FAIR"
            Case "POOR", "P", "BAD", "REPLACE", "REPLACED", "FAIL", "FAILED"
                Return "POOR"
            Case "NO BATTERY", "NO_BATTERY", "NB", "NONE"
                Return "GOOD"
            Case "Y", "YES"
                Return "GOOD"
            Case "N", "NO", "N/A", "NA", ""
                Return "GOOD"
            Case Else
                Return "GOOD"
        End Select

    End Function

#End Region

End Class
