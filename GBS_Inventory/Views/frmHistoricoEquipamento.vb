Imports System.Windows.Forms
Imports System.Drawing
Imports System.Data
Imports System.IO
Imports System.Text

''' <summary>
''' Histórico completo de um equipamento: chegada, upgrades e remessas.
''' </summary>
Public Class frmHistoricoEquipamento

#Region "Atributos"

    Private _idEquipamento  As Integer
    Private _internalUID    As String
    Private _oHistorico     As clsLeituraHistorico

    Private _dtEquipamento  As DataTable
    Private _dtUpgrades     As DataTable
    Private _dtShipment     As DataTable

#End Region

#Region "Construtor"

    Public Sub New(pIdEquipamento As Integer, pInternalUID As String)
        InitializeComponent()
        _idEquipamento = pIdEquipamento
        _internalUID   = pInternalUID
        _oHistorico    = New clsLeituraHistorico()
        TemaEscuro.aplicarHelius(Me)
    End Sub

#End Region

#Region "Carregamento"

    Private Sub frmHistoricoEquipamento_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            CarregarTudo()
        Catch ex As Exception
            MessageBox.Show("Error loading history: " & ex.Message, "Error",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub CarregarTudo()
        CarregarEquipamento()
        CarregarUpgrades()
        CarregarShipment()
    End Sub

    Private Sub CarregarEquipamento()

        Dim ds As DataSet = _oHistorico.selecionarEquipamento(_idEquipamento)
        If ds Is Nothing OrElse ds.Tables.Count = 0 OrElse ds.Tables(0).Rows.Count = 0 Then Return

        _dtEquipamento = ds.Tables(0)
        Dim row As DataRow = _dtEquipamento.Rows(0)

        Dim uid   As String = ObterTexto(row, {"INTERNAL_UID"})
        Dim marca As String = ObterTexto(row, {"MANUFACTURER", "MARCA"})
        Dim model As String = ObterTexto(row, {"MODEL", "MODELO"})
        Dim sn    As String = ObterTexto(row, {"SERIAL_NUMBER"})
        Dim cpu   As String = ObterTexto(row, {"CPU_MODEL", "PROCESSADOR"})
        Dim ram   As String = ObterTexto(row, {"RAM_GB"})
        Dim sto   As String = ObterTexto(row, {"STORAGE_GB"})
        Dim sta   As String = ObterTexto(row, {"STATUS"})

        lblUID.Text    = uid
        lblHeaderInfo.Text = $"{marca}  {model}  |  S/N: {sn}  |  {cpu}  {ram}GB RAM  {sto}GB  |  Status: {sta}"

        ' ABA 1 — Arrival
        Dim dataCad  As String = ObterTexto(row, {"DATA_CADASTRO"})
        Dim batch    As String = ObterTexto(row, {"SOURCE_BATCH"})
        Dim condSt   As String = ObterTexto(row, {"CONDITION_STATUS"})
        Dim obs      As String = ObterTexto(row, {"OBSERVACAO", "NOTES"})

        SetField(lblArrDateVal,      dataCad)
        SetField(lblArrBatchVal,     batch)
        SetField(lblArrCondVal,      condSt)
        SetField(lblArrStatusVal,    sta)
        SetField(lblArrSerialVal,    sn)
        SetField(txtArrObs,         obs)

    End Sub

    Private Sub CarregarUpgrades()

        Dim ds As DataSet = _oHistorico.selecionarUpgradesEquipamento(_idEquipamento)
        If ds IsNot Nothing AndAlso ds.Tables.Count > 0 Then
            _dtUpgrades = ds.Tables(0)
        Else
            _dtUpgrades = New DataTable()
        End If

        dgvUpgrades.DataSource = _dtUpgrades

        ConfigurarColunasUpgrade()
        CalcularTotalUpgrade()

    End Sub

    Private Sub ConfigurarColunasUpgrade()

        Dim mapa As New Dictionary(Of String, String) From {
            {"ID_UPGRADE",     ""},
            {"DATA_UPGRADE",   "Date"},
            {"COMPONENT_TYPE", "Component"},
            {"VALUE_BEFORE",   "Before"},
            {"VALUE_AFTER",    "After"},
            {"SOURCE_ORIGEM",  "Source"},
            {"COST_USD",       "Cost (USD)"},
            {"PART_SERIAL",    "Part S/N"},
            {"TECHNICIAN",     "Technician"},
            {"NOTES",          "Notes"}
        }

        For Each col As DataGridViewColumn In dgvUpgrades.Columns
            If mapa.ContainsKey(col.Name) Then
                If String.IsNullOrEmpty(mapa(col.Name)) Then
                    col.Visible = False
                Else
                    col.HeaderText = mapa(col.Name)
                End If
            End If
        Next

        If dgvUpgrades.Columns.Contains("DATA_UPGRADE") Then
            dgvUpgrades.Columns("DATA_UPGRADE").DefaultCellStyle.Format = "yyyy-MM-dd"
            dgvUpgrades.Columns("DATA_UPGRADE").Width = 90
        End If
        If dgvUpgrades.Columns.Contains("COST_USD") Then
            dgvUpgrades.Columns("COST_USD").DefaultCellStyle.Format = "N2"
            dgvUpgrades.Columns("COST_USD").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            dgvUpgrades.Columns("COST_USD").Width = 90
        End If
        If dgvUpgrades.Columns.Contains("NOTES") Then
            dgvUpgrades.Columns("NOTES").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
        End If

    End Sub

    Private Sub CalcularTotalUpgrade()

        Dim total As Decimal = 0
        If _dtUpgrades IsNot Nothing AndAlso _dtUpgrades.Columns.Contains("COST_USD") Then
            For Each row As DataRow In _dtUpgrades.Rows
                If Not IsDBNull(row("COST_USD")) Then
                    total += CDec(row("COST_USD"))
                End If
            Next
        End If

        Dim cntUpg As Integer = If(_dtUpgrades IsNot Nothing, _dtUpgrades.Rows.Count, 0)
        lblUpgradeTotal.Text = "Total upgrade cost: $ " & total.ToString("N2") & "   |   " & cntUpg.ToString() & " record(s)"

    End Sub

    Private Sub CarregarShipment()

        Dim ds As DataSet = _oHistorico.selecionarShipmentEquipamento(_idEquipamento)
        If ds IsNot Nothing AndAlso ds.Tables.Count > 0 Then
            _dtShipment = ds.Tables(0)
        Else
            _dtShipment = New DataTable()
        End If

        dgvShipment.DataSource = _dtShipment
        ConfigurarColunasShipment()
        CalcularTotalShipment()

    End Sub

    Private Sub ConfigurarColunasShipment()

        Dim mapa As New Dictionary(Of String, String) From {
            {"ID_REMESSA",       ""},
            {"REMESSA_REF",      "Shipment Ref"},
            {"CARRIER",          "Carrier"},
            {"TRACKING_NUMBER",  "Tracking"},
            {"RECIPIENT_NAME",   "Recipient"},
            {"DATA_ENVIO",       "Est. Delivery"},
            {"DATA_ENTREGA",     "Actual Delivery"},
            {"STATUS_REMESSA",   "Status"},
            {"SHIPPING_COST_USD","Shipping Cost"},
            {"SALE_PRICE_USD",   "Sale Price (USD)"},
            {"CONDITION_AT_SHIP","Condition"},
            {"ITEM_NOTES",       "Notes"}
        }

        For Each col As DataGridViewColumn In dgvShipment.Columns
            If mapa.ContainsKey(col.Name) Then
                If String.IsNullOrEmpty(mapa(col.Name)) Then
                    col.Visible = False
                Else
                    col.HeaderText = mapa(col.Name)
                End If
            End If
        Next

        For Each colName As String In {"DATA_ENVIO", "DATA_ENTREGA"}
            If dgvShipment.Columns.Contains(colName) Then
                dgvShipment.Columns(colName).DefaultCellStyle.Format = "yyyy-MM-dd"
                dgvShipment.Columns(colName).Width = 100
            End If
        Next

        For Each colName As String In {"SHIPPING_COST_USD", "SALE_PRICE_USD"}
            If dgvShipment.Columns.Contains(colName) Then
                dgvShipment.Columns(colName).DefaultCellStyle.Format = "N2"
                dgvShipment.Columns(colName).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                dgvShipment.Columns(colName).Width = 110
            End If
        Next

        If dgvShipment.Columns.Contains("ITEM_NOTES") Then
            dgvShipment.Columns("ITEM_NOTES").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
        End If

    End Sub

    Private Sub CalcularTotalShipment()

        Dim totalSale     As Decimal = 0
        Dim totalShipping As Decimal = 0
        If _dtShipment IsNot Nothing Then
            For Each row As DataRow In _dtShipment.Rows
                If _dtShipment.Columns.Contains("SALE_PRICE_USD") AndAlso Not IsDBNull(row("SALE_PRICE_USD")) Then
                    totalSale += CDec(row("SALE_PRICE_USD"))
                End If
                If _dtShipment.Columns.Contains("SHIPPING_COST_USD") AndAlso Not IsDBNull(row("SHIPPING_COST_USD")) Then
                    totalShipping += CDec(row("SHIPPING_COST_USD"))
                End If
            Next
        End If

        Dim cntShip As Integer = If(_dtShipment IsNot Nothing, _dtShipment.Rows.Count, 0)
        lblShipmentTotal.Text = "Sale price: $ " & totalSale.ToString("N2") & "   |   Shipping cost: $ " & totalShipping.ToString("N2") & "   |   " & cntShip.ToString() & " shipment(s)"

    End Sub

#End Region

#Region "Relatório"

    Private Sub btnHistoryReport_Click(sender As Object, e As EventArgs) Handles btnHistoryReport.Click
        GerarRelatorioHistorico()
    End Sub

    Private Sub btnUpgradeReport_Click(sender As Object, e As EventArgs) Handles btnUpgradeReport.Click
        GerarRelatorioUpgrade()
    End Sub

    Private Sub GerarRelatorioUpgrade()

        Try
            Dim outputPath As String = System.Configuration.ConfigurationManager.AppSettings("ReportsOutputPath")
            If String.IsNullOrWhiteSpace(outputPath) Then outputPath = "C:\GBS\Reports"
            If Not Directory.Exists(outputPath) Then Directory.CreateDirectory(outputPath)

            Dim logoPath As String = System.Configuration.ConfigurationManager.AppSettings("InvoiceLogoPath")
            If String.IsNullOrWhiteSpace(logoPath) Then logoPath = ""

            Dim nomeArq As String = "UpgradeReport_" & _internalUID.Replace("/", "-") & "_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".doc"
            Dim caminho As String = Path.Combine(outputPath, nomeArq)

            File.WriteAllText(caminho, MontarHtmlUpgrade(logoPath), System.Text.Encoding.UTF8)
            System.Diagnostics.Process.Start(caminho)

            MessageBox.Show("Upgrade report generated!" & vbCrLf & caminho,
                            "Upgrade Report", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show("Error generating upgrade report: " & ex.Message, "Error",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Function MontarHtmlUpgrade(pLogoPath As String) As String

        Dim sb As New StringBuilder()

        sb.AppendLine("<html><head><meta charset='utf-8'/>")
        sb.AppendLine("<style>")
        sb.AppendLine("body{font-family:Segoe UI,Arial,sans-serif;font-size:10pt;color:#1f2937;margin:40px;}")
        sb.AppendLine("h2{color:#1f2937;margin-bottom:2px;font-size:16pt;}")
        sb.AppendLine(".sub{font-size:9pt;color:#6b7280;margin-bottom:18px;}")
        sb.AppendLine(".device-box{background:#f3f4f6;border:1px solid #d1d5db;border-radius:6px;padding:12px 16px;margin-bottom:22px;font-size:9.5pt;line-height:1.8;}")
        sb.AppendLine(".device-box b{color:#374151;}")
        sb.AppendLine("table{table-layout:fixed;width:100%;border-collapse:collapse;margin-top:6px;}")
        sb.AppendLine("th{background:#1f2937;color:#fff;padding:7px 8px;text-align:left;font-size:9pt;}")
        sb.AppendLine("td{border:1px solid #e5e7eb;padding:6px 8px;font-size:9pt;vertical-align:middle;word-wrap:break-word;}")
        sb.AppendLine("tr:nth-child(even) td{background:#f9fafb;}")
        sb.AppendLine(".footer-note{font-size:8.5pt;color:#6b7280;margin-top:24px;border-top:1px solid #e5e7eb;padding-top:10px;}")
        sb.AppendLine(".total{font-size:10pt;font-weight:bold;color:#1f2937;margin-top:10px;}")
        sb.AppendLine("</style></head><body>")

        ' Logo
        If Not String.IsNullOrWhiteSpace(pLogoPath) AndAlso File.Exists(pLogoPath) Then
            Dim bytes As Byte() = File.ReadAllBytes(pLogoPath)
            Dim b64   As String = Convert.ToBase64String(bytes)
            Dim mime  As String = If(pLogoPath.ToLower().EndsWith(".png"), "image/png", "image/jpeg")
            sb.AppendLine("<div style='margin-bottom:16px;'><img src='data:" & mime & ";base64," & b64 & "' style='max-height:60px;max-width:200px;'/></div>")
        End If

        sb.AppendLine("<h2>Hardware Upgrade Report</h2>")
        sb.AppendLine("<div class='sub'>Issued: " & DateTime.Now.ToString("MMMM dd, yyyy") & " &nbsp;|&nbsp; Ref: " & HtmlE(_internalUID) & "</div>")

        ' Bloco do equipamento
        If _dtEquipamento IsNot Nothing AndAlso _dtEquipamento.Rows.Count > 0 Then
            Dim r As DataRow = _dtEquipamento.Rows(0)
            sb.AppendLine("<div class='device-box'>")
            sb.AppendLine("<b>Device:</b> " & HtmlE(ObterTexto(r, {"MANUFACTURER", "MARCA"})) & " " & HtmlE(ObterTexto(r, {"MODEL"})) & "<br/>")
            sb.AppendLine("<b>Serial Number:</b> " & HtmlE(ObterTexto(r, {"SERIAL_NUMBER"})) & "<br/>")
            sb.AppendLine("<b>Processor:</b> " & HtmlE(ObterTexto(r, {"CPU_MODEL", "PROCESSADOR"})) & "<br/>")
            sb.AppendLine("<b>Current RAM:</b> " & HtmlE(ObterTexto(r, {"RAM_GB"})) & " GB &nbsp;&nbsp; <b>Current Storage:</b> " & HtmlE(ObterTexto(r, {"STORAGE_GB"})) & " GB")
            sb.AppendLine("</div>")
        End If

        ' Tabela de upgrades
        sb.AppendLine("<table>")
        sb.AppendLine("<tr>")
        sb.AppendLine("<th style='width:90px'>Date</th>")
        sb.AppendLine("<th style='width:110px'>Component</th>")
        sb.AppendLine("<th style='width:80px'>Before</th>")
        sb.AppendLine("<th style='width:80px'>After</th>")
        sb.AppendLine("<th style='width:90px'>Technician</th>")
        sb.AppendLine("<th>Notes</th>")
        sb.AppendLine("</tr>")

        Dim totalUpg As Decimal = 0
        If _dtUpgrades IsNot Nothing Then
            For Each row As DataRow In _dtUpgrades.Rows
                sb.Append("<tr>")
                sb.Append("<td>" & HtmlE(FormatarData(row, "DATA_UPGRADE")) & "</td>")
                sb.Append("<td>" & HtmlE(ObterTexto(row, {"COMPONENT_TYPE"})) & "</td>")
                sb.Append("<td>" & HtmlE(ObterTexto(row, {"VALUE_BEFORE"})) & "</td>")
                sb.Append("<td><b>" & HtmlE(ObterTexto(row, {"VALUE_AFTER"})) & "</b></td>")
                sb.Append("<td>" & HtmlE(ObterTexto(row, {"TECHNICIAN"})) & "</td>")
                sb.Append("<td>" & HtmlE(LimparObservacao(ObterTexto(row, {"NOTES"}))) & "</td>")
                totalUpg += ObterDecimal(row, "COST_USD")
                sb.AppendLine("</tr>")
            Next
        End If

        sb.AppendLine("</table>")

        If totalUpg > 0 Then
            sb.AppendLine("<div class='total'>Total upgrade cost: $ " & totalUpg.ToString("N2") & "</div>")
        End If

        sb.AppendLine("<div class='footer-note'>")
        sb.AppendLine("This document certifies the hardware upgrades performed on the device described above.<br/>")
        sb.AppendLine("GBS — Global Business Solution &nbsp;|&nbsp; " & DateTime.Now.ToString("yyyy"))
        sb.AppendLine("</div>")

        sb.AppendLine("</body></html>")
        Return sb.ToString()

    End Function

    ' Remove campos internos (Source: / CostUSD:) da observação antes de exibir ao cliente
    Private Function LimparObservacao(pNotes As String) As String
        If String.IsNullOrWhiteSpace(pNotes) Then Return ""
        Dim partes As String() = pNotes.Split({" | "}, StringSplitOptions.RemoveEmptyEntries)
        Dim resultado As New StringBuilder()
        For Each parte As String In partes
            Dim t As String = parte.Trim()
            If Not t.StartsWith("Source:") AndAlso Not t.StartsWith("CostUSD:") Then
                If resultado.Length > 0 Then resultado.Append(" | ")
                resultado.Append(t)
            End If
        Next
        Return resultado.ToString()
    End Function

    Private Sub GerarRelatorioHistorico()

        Try
            Dim outputPath As String = System.Configuration.ConfigurationManager.AppSettings("ReportsOutputPath")
            If String.IsNullOrWhiteSpace(outputPath) Then outputPath = "C:\GBS\Reports"
            If Not Directory.Exists(outputPath) Then Directory.CreateDirectory(outputPath)

            Dim logoPath As String = System.Configuration.ConfigurationManager.AppSettings("InvoiceLogoPath")
            If String.IsNullOrWhiteSpace(logoPath) Then logoPath = ""

            Dim nomeArq  As String = "History_" & _internalUID.Replace("/", "-") & "_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".doc"
            Dim caminho  As String = Path.Combine(outputPath, nomeArq)

            Dim html As String = MontarHtmlHistorico(logoPath)
            File.WriteAllText(caminho, html, System.Text.Encoding.UTF8)

            System.Diagnostics.Process.Start(caminho)

            MessageBox.Show("History report generated!" & vbCrLf & caminho,
                            "Generate History Report", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show("Error generating report: " & ex.Message, "Error",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Function MontarHtmlHistorico(pLogoPath As String) As String

        Dim sb As New StringBuilder()

        sb.AppendLine("<html><head><meta charset='utf-8'/>")
        sb.AppendLine("<style>")
        sb.AppendLine("body{font-family:Segoe UI,Arial,sans-serif;font-size:10pt;color:#1f2937;margin:36px;}")
        sb.AppendLine("h2{color:#1f2937;margin-bottom:4px;} h3{color:#374151;margin:18px 0 6px;}")
        sb.AppendLine(".sub{font-size:9pt;color:#6b7280;margin-bottom:14px;}")
        sb.AppendLine(".info-box{background:#f3f4f6;border:1px solid #d1d5db;border-radius:6px;padding:10px 14px;margin-bottom:18px;font-size:9pt;}")
        sb.AppendLine(".info-box b{color:#374151;}")
        sb.AppendLine("table{table-layout:fixed;width:100%;border-collapse:collapse;margin-top:6px;}")
        sb.AppendLine("th{background:#374151;color:#fff;padding:6px;text-align:left;font-size:8.5pt;}")
        sb.AppendLine("td{border:1px solid #e5e7eb;padding:5px 6px;font-size:8.5pt;vertical-align:top;word-wrap:break-word;}")
        sb.AppendLine("tr:nth-child(even) td{background:#f9fafb;}")
        sb.AppendLine(".footer{font-size:9pt;color:#374151;margin-top:6px;font-weight:bold;}")
        sb.AppendLine("</style></head><body>")

        ' Logo
        If Not String.IsNullOrWhiteSpace(pLogoPath) AndAlso File.Exists(pLogoPath) Then
            Dim bytes As Byte() = File.ReadAllBytes(pLogoPath)
            Dim b64   As String = Convert.ToBase64String(bytes)
            Dim mime  As String = If(pLogoPath.ToLower().EndsWith(".png"), "image/png", "image/jpeg")
            sb.AppendLine("<div style='margin-bottom:12px;'><img src='data:" & mime & ";base64," & b64 & "' style='max-height:60px;max-width:200px;'/></div>")
        End If

        sb.AppendLine("<h2>Equipment History Report</h2>")
        sb.AppendLine("<div class='sub'>UID: <b>" & HtmlE(_internalUID) & "</b>  &nbsp;|&nbsp; Generated: " & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & "</div>")

        ' Header info box
        If _dtEquipamento IsNot Nothing AndAlso _dtEquipamento.Rows.Count > 0 Then
            Dim r As DataRow = _dtEquipamento.Rows(0)
            sb.AppendLine("<div class='info-box'>")
            sb.AppendLine("<b>Manufacturer:</b> " & HtmlE(ObterTexto(r, {"MANUFACTURER", "MARCA"})) & " &nbsp; ")
            sb.AppendLine("<b>Model:</b> " & HtmlE(ObterTexto(r, {"MODEL"})) & " &nbsp; ")
            sb.AppendLine("<b>S/N:</b> " & HtmlE(ObterTexto(r, {"SERIAL_NUMBER"})) & " &nbsp; ")
            sb.AppendLine("<b>CPU:</b> " & HtmlE(ObterTexto(r, {"CPU_MODEL", "PROCESSADOR"})) & " &nbsp; ")
            sb.AppendLine("<b>RAM:</b> " & HtmlE(ObterTexto(r, {"RAM_GB"})) & " GB &nbsp; ")
            sb.AppendLine("<b>Storage:</b> " & HtmlE(ObterTexto(r, {"STORAGE_GB"})) & " GB &nbsp; ")
            sb.AppendLine("<b>Status:</b> " & HtmlE(ObterTexto(r, {"STATUS"})))
            sb.AppendLine("</div>")
        End If

        ' ABA 1 — Arrival
        sb.AppendLine("<h3>1. Arrival</h3>")
        If _dtEquipamento IsNot Nothing AndAlso _dtEquipamento.Rows.Count > 0 Then
            Dim r As DataRow = _dtEquipamento.Rows(0)
            sb.AppendLine("<table><tr>")
            sb.AppendLine("<th style='width:120px'>Date Registered</th><th style='width:120px'>Source Batch</th><th style='width:100px'>Condition</th><th style='width:100px'>Status</th><th style='width:120px'>Serial Number</th><th>Notes</th>")
            sb.AppendLine("</tr><tr>")
            sb.AppendLine("<td>" & HtmlE(ObterTexto(r, {"DATA_CADASTRO"})) & "</td>")
            sb.AppendLine("<td>" & HtmlE(ObterTexto(r, {"SOURCE_BATCH"})) & "</td>")
            sb.AppendLine("<td>" & HtmlE(ObterTexto(r, {"CONDITION_STATUS"})) & "</td>")
            sb.AppendLine("<td>" & HtmlE(ObterTexto(r, {"STATUS"})) & "</td>")
            sb.AppendLine("<td>" & HtmlE(ObterTexto(r, {"SERIAL_NUMBER"})) & "</td>")
            sb.AppendLine("<td>" & HtmlE(ObterTexto(r, {"OBSERVACAO", "NOTES"})) & "</td>")
            sb.AppendLine("</tr></table>")
        End If

        ' ABA 2 — Upgrades
        sb.AppendLine("<h3>2. Upgrades</h3>")
        sb.AppendLine("<table><tr>")
        sb.AppendLine("<th style='width:90px'>Date</th><th style='width:100px'>Component</th><th style='width:90px'>Before</th><th style='width:90px'>After</th><th style='width:80px'>Source</th><th style='width:80px'>Cost (USD)</th><th style='width:100px'>Part S/N</th><th style='width:90px'>Technician</th><th>Notes</th>")
        sb.AppendLine("</tr>")

        Dim totalUpg As Decimal = 0
        If _dtUpgrades IsNot Nothing Then
            For Each row As DataRow In _dtUpgrades.Rows
                sb.Append("<tr>")
                sb.Append("<td>" & HtmlE(FormatarData(row, "DATA_UPGRADE")) & "</td>")
                sb.Append("<td>" & HtmlE(ObterTexto(row, {"COMPONENT_TYPE"})) & "</td>")
                sb.Append("<td>" & HtmlE(ObterTexto(row, {"VALUE_BEFORE"})) & "</td>")
                sb.Append("<td>" & HtmlE(ObterTexto(row, {"VALUE_AFTER"})) & "</td>")
                sb.Append("<td>" & HtmlE(ObterTexto(row, {"SOURCE_ORIGEM"})) & "</td>")
                Dim custo As Decimal = ObterDecimal(row, "COST_USD")
                totalUpg += custo
                sb.Append("<td style='text-align:right'>" & custo.ToString("N2") & "</td>")
                sb.Append("<td>" & HtmlE(ObterTexto(row, {"PART_SERIAL"})) & "</td>")
                sb.Append("<td>" & HtmlE(ObterTexto(row, {"TECHNICIAN"})) & "</td>")
                sb.Append("<td>" & HtmlE(ObterTexto(row, {"NOTES"})) & "</td>")
                sb.AppendLine("</tr>")
            Next
        End If

        sb.AppendLine("</table>")
        sb.AppendLine("<div class='footer'>Total upgrade cost: $ " & totalUpg.ToString("N2") & "</div>")

        ' ABA 3 — Shipment
        sb.AppendLine("<h3>3. Shipments</h3>")
        sb.AppendLine("<table><tr>")
        sb.AppendLine("<th style='width:90px'>Ref</th><th style='width:80px'>Carrier</th><th style='width:110px'>Tracking</th><th style='width:120px'>Recipient</th><th style='width:90px'>Est. Delivery</th><th style='width:90px'>Actual Delivery</th><th style='width:80px'>Status</th><th style='width:90px'>Sale Price</th><th style='width:90px'>Ship. Cost</th><th>Notes</th>")
        sb.AppendLine("</tr>")

        Dim totalSale As Decimal = 0
        If _dtShipment IsNot Nothing Then
            For Each row As DataRow In _dtShipment.Rows
                sb.Append("<tr>")
                sb.Append("<td>" & HtmlE(ObterTexto(row, {"REMESSA_REF"})) & "</td>")
                sb.Append("<td>" & HtmlE(ObterTexto(row, {"CARRIER"})) & "</td>")
                sb.Append("<td>" & HtmlE(ObterTexto(row, {"TRACKING_NUMBER"})) & "</td>")
                sb.Append("<td>" & HtmlE(ObterTexto(row, {"RECIPIENT_NAME"})) & "</td>")
                sb.Append("<td>" & HtmlE(FormatarData(row, "DATA_ENVIO")) & "</td>")
                sb.Append("<td>" & HtmlE(FormatarData(row, "DATA_ENTREGA")) & "</td>")
                sb.Append("<td>" & HtmlE(ObterTexto(row, {"STATUS_REMESSA"})) & "</td>")
                Dim preco As Decimal = ObterDecimal(row, "SALE_PRICE_USD")
                totalSale += preco
                sb.Append("<td style='text-align:right'>" & preco.ToString("N2") & "</td>")
                sb.Append("<td style='text-align:right'>" & ObterDecimal(row, "SHIPPING_COST_USD").ToString("N2") & "</td>")
                sb.Append("<td>" & HtmlE(ObterTexto(row, {"ITEM_NOTES"})) & "</td>")
                sb.AppendLine("</tr>")
            Next
        End If

        sb.AppendLine("</table>")
        sb.AppendLine("<div class='footer'>Total sale price: $ " & totalSale.ToString("N2") & "</div>")

        sb.AppendLine("</body></html>")
        Return sb.ToString()

    End Function

#End Region

#Region "Botão Close"

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

#End Region

#Region "Helpers"

    Private Function ObterTexto(pRow As DataRow, pColunas As String()) As String
        If pRow Is Nothing Then Return ""
        For Each nome As String In pColunas
            If pRow.Table.Columns.Contains(nome) AndAlso Not IsDBNull(pRow(nome)) Then
                Return pRow(nome).ToString().Trim()
            End If
        Next
        Return ""
    End Function

    Private Function ObterDecimal(pRow As DataRow, pColuna As String) As Decimal
        If pRow Is Nothing OrElse Not pRow.Table.Columns.Contains(pColuna) OrElse IsDBNull(pRow(pColuna)) Then Return 0
        Dim v As Decimal
        If Decimal.TryParse(pRow(pColuna).ToString(), v) Then Return v
        Return 0
    End Function

    Private Function FormatarData(pRow As DataRow, pColuna As String) As String
        If pRow Is Nothing OrElse Not pRow.Table.Columns.Contains(pColuna) OrElse IsDBNull(pRow(pColuna)) Then Return ""
        Dim dt As DateTime
        If DateTime.TryParse(pRow(pColuna).ToString(), dt) Then Return dt.ToString("yyyy-MM-dd")
        Return pRow(pColuna).ToString()
    End Function

    Private Function HtmlE(s As String) As String
        If s Is Nothing Then Return ""
        Return s.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("""", "&quot;")
    End Function

    Private Sub SetField(pCtrl As Label, pValor As String)
        If pCtrl IsNot Nothing Then pCtrl.Text = If(String.IsNullOrWhiteSpace(pValor), "—", pValor)
    End Sub

    Private Sub SetField(pCtrl As TextBox, pValor As String)
        If pCtrl IsNot Nothing Then pCtrl.Text = pValor
    End Sub

#End Region

End Class
