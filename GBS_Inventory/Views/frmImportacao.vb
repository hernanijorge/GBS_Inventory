Imports System.Windows.Forms
Imports System.Drawing
Imports System.IO
Imports System.Data
Imports Oracle.ManagedDataAccess.Client
Imports System.Configuration
Imports GBS_Inventory.OracleHelper

''' <summary>
''' Tela de Importação de Planilha Excel da GBS.
''' </summary>
Public Class frmImportacao

    Private oController          As ImportacaoController
    Private sArquivoSelecionado  As String    = ""
    Private _dtResultado         As DataTable = Nothing
    Private _chkShowIssuesOnly   As CheckBox
    Private _dgvResultado        As DataGridView
    Private _lblResultadoFooter  As Label
    Private _lblAnaliseTitle     As Label

    Public Sub New()
        InitializeComponent()
        oController = New ImportacaoController()
        TemaEscuro.aplicarHelius(Me)
        InicializarGridAnalise()
    End Sub

    Private Sub btnSelecionar_Click(sender As Object, e As EventArgs) Handles btnSelecionar.Click

        Using ofd As New OpenFileDialog()
            ofd.Filter = "Planilhas Excel (*.xlsx)|*.xlsx|Todos os arquivos (*.*)|*.*"
            ofd.Title  = "Selecione a planilha da GBS"

            If ofd.ShowDialog() = DialogResult.OK Then
                sArquivoSelecionado = ofd.FileName
                lblArquivo.Text     = Path.GetFileName(ofd.FileName)
                lblArquivo.ForeColor = TemaEscuro.Accent
                btnImportar.Enabled = True
            End If
        End Using

    End Sub

    Private Sub btnImportar_Click(sender As Object, e As EventArgs) Handles btnImportar.Click

        If String.IsNullOrEmpty(sArquivoSelecionado) Then
            MessageBox.Show("Select a spreadsheet first.", "Warning",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        btnImportar.Enabled   = False
        btnSelecionar.Enabled = False
        btnFechar.Enabled     = False

        lblStatus.Text        = "Importing..."
        lblStatus.ForeColor   = TemaEscuro.Accent
        progress.Value        = 0
        progress.Visible      = True

        Application.DoEvents()

        Try

            Dim callback As Action(Of Integer, Integer, String) = _
                Sub(atual, total, uid)
                    Me.Invoke(Sub()
                        progress.Maximum = total
                        progress.Value   = Math.Min(atual, total)
                        lblStatus.Text   = $"Processando {atual}/{total} — UID {uid}"
                        Application.DoEvents()
                    End Sub)
                End Sub

            oController.importarPlanilha(sArquivoSelecionado, callback)

            progress.Value     = progress.Maximum
            lblStatus.Text     = "Done!"
            lblStatus.ForeColor= TemaEscuro.Accent

            lblInseridos.Text   = "Inserted: "     & oController.TotalInseridos.ToString()
            lblAtualizados.Text = "Updated: "      & oController.TotalAtualizados.ToString()
            lblAbas.Text        = "Sheets processed: " & oController.TotalAbas.ToString()
            lblLinhas.Text      = "Lines read: "   & oController.TotalLinhasPlanilha.ToString()
            lblErros.Text       = "Errors: "       & oController.TotalErros.ToString()

            lblSkipped.Text = "Skipped: " & oController.TotalSkipped.ToString()

            If oController.TotalErros > 0 Then
                txtErros.Lines     = oController.Erros.ToArray()
                txtErros.Visible   = True
                lblErros.ForeColor = TemaEscuro.Vermelho
            End If

            CarregarGridAnalise()

            MessageBox.Show($"Import completed!{vbCrLf & vbCrLf}" &
                            $"Inserted: {oController.TotalInseridos}{vbCrLf}" &
                            $"Updated: {oController.TotalAtualizados}{vbCrLf}" &
                            $"Errors: {oController.TotalErros}",
                            "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception

            MessageBox.Show("Import error: " & ex.Message, "Error",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
            lblStatus.Text     = "Import failed"
            lblStatus.ForeColor= TemaEscuro.Vermelho

        Finally

            btnImportar.Enabled   = True
            btnSelecionar.Enabled = True
            btnFechar.Enabled     = True

        End Try

    End Sub

    Private Sub btnFechar_Click(sender As Object, e As EventArgs) Handles btnFechar.Click
        Me.Close()
    End Sub

#Region "Grid de Análise"

    Private Sub InicializarGridAnalise()

        _lblAnaliseTitle           = New Label()
        _lblAnaliseTitle.Text      = "Import Analysis"
        _lblAnaliseTitle.Font      = New Font("Segoe UI", 10, FontStyle.Bold)
        _lblAnaliseTitle.ForeColor = TemaEscuro.Accent
        _lblAnaliseTitle.Location  = New Point(20, 574)
        _lblAnaliseTitle.AutoSize  = True
        _lblAnaliseTitle.Visible   = False

        _chkShowIssuesOnly           = New CheckBox()
        _chkShowIssuesOnly.Text      = "Show issues only"
        _chkShowIssuesOnly.Location  = New Point(430, 576)
        _chkShowIssuesOnly.AutoSize  = True
        _chkShowIssuesOnly.ForeColor = TemaEscuro.Texto
        _chkShowIssuesOnly.BackColor = Color.Transparent
        _chkShowIssuesOnly.Visible   = False
        AddHandler _chkShowIssuesOnly.CheckedChanged, AddressOf OnFiltroChanged

        _dgvResultado          = New DataGridView()
        _dgvResultado.Location = New Point(20, 600)
        _dgvResultado.Size     = New Size(660, 220)
        _dgvResultado.Anchor   = AnchorStyles.Top Or AnchorStyles.Left Or
                                  AnchorStyles.Right Or AnchorStyles.Bottom
        _dgvResultado.Visible  = False
        TemaEscuro.estilizarGridHelius(_dgvResultado)
        AddHandler _dgvResultado.CellFormatting, AddressOf OnGridCellFormatting

        _lblResultadoFooter           = New Label()
        _lblResultadoFooter.Text      = ""
        _lblResultadoFooter.Location  = New Point(20, 828)
        _lblResultadoFooter.Size      = New Size(660, 20)
        _lblResultadoFooter.Anchor    = AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
        _lblResultadoFooter.ForeColor = TemaEscuro.TextoMutado
        _lblResultadoFooter.Visible   = False

        Me.Controls.AddRange(New Control() {
            _lblAnaliseTitle, _chkShowIssuesOnly, _dgvResultado, _lblResultadoFooter
        })

    End Sub

    Private Sub CarregarGridAnalise()
        If oController.UIDsImportados.Count = 0 Then Return

        Try
            Dim cs As String = ConfigurationManager.ConnectionStrings("OracleDB").ConnectionString
            Dim w  As String = BuildInClause(oController.UIDsImportados)

            Dim sqlFull As String =
                "SELECT INTERNAL_UID, MARCA AS MANUFACTURER, MODEL," &
                "       CONDITION_STATUS, SOURCE_BATCH, OBSERVACAO" &
                "  FROM TBL_EQUIPAMENTO WHERE " & w & " ORDER BY SOURCE_BATCH, INTERNAL_UID"

            Dim sqlNoExt As String =
                "SELECT INTERNAL_UID, MARCA AS MANUFACTURER, MODEL," &
                "       CAST(NULL AS VARCHAR2(20)) AS CONDITION_STATUS," &
                "       CAST(NULL AS VARCHAR2(100)) AS SOURCE_BATCH, OBSERVACAO" &
                "  FROM TBL_EQUIPAMENTO WHERE " & w & " ORDER BY INTERNAL_UID"

            Dim sqlEn As String =
                "SELECT INTERNAL_UID, MANUFACTURER, MODEL," &
                "       CAST(NULL AS VARCHAR2(20)) AS CONDITION_STATUS," &
                "       CAST(NULL AS VARCHAR2(100)) AS SOURCE_BATCH," &
                "       CAST(NULL AS VARCHAR2(4000)) AS OBSERVACAO" &
                "  FROM TBL_EQUIPAMENTO WHERE " & w & " ORDER BY INTERNAL_UID"

            Dim dtSrc As DataTable = Nothing
            Dim lastEx As Exception = Nothing
            For Each sql As String In New String() {sqlFull, sqlNoExt, sqlEn}
                Try
                    Dim ds As DataSet = OracleHelper.ExecuteDataset(cs, CommandType.Text, sql)
                    If ds IsNot Nothing AndAlso ds.Tables.Count > 0 Then
                        dtSrc = ds.Tables(0)
                        Exit For
                    End If
                Catch ex As Exception
                    lastEx = ex
                    If Not ex.Message.Contains("ORA-00904") Then Throw
                End Try
            Next
            If dtSrc Is Nothing Then
                If lastEx IsNot Nothing Then Throw lastEx
                Return
            End If

            ' Build result DataTable with ISSUE column
            _dtResultado = New DataTable()
            _dtResultado.Columns.Add("INTERNAL_UID",    GetType(String))
            _dtResultado.Columns.Add("MANUFACTURER",    GetType(String))
            _dtResultado.Columns.Add("MODEL",           GetType(String))
            _dtResultado.Columns.Add("CONDITION_STATUS",GetType(String))
            _dtResultado.Columns.Add("SOURCE_BATCH",    GetType(String))
            _dtResultado.Columns.Add("OBSERVATION",     GetType(String))
            _dtResultado.Columns.Add("ISSUE",           GetType(String))
            _dtResultado.Columns.Add("_FLAG",           GetType(Boolean))

            For Each row As DataRow In dtSrc.Rows
                Dim cond As String = ColStr(row, "CONDITION_STATUS").ToUpperInvariant().Trim()
                Dim obs  As String = ColStr(row, "OBSERVACAO")
                Dim flag As Boolean = (cond = "FAIR" OrElse cond = "POOR") OrElse
                                       Not String.IsNullOrEmpty(obs)
                _dtResultado.Rows.Add(
                    ColStr(row, "INTERNAL_UID"),
                    ColStr(row, "MANUFACTURER"),
                    ColStr(row, "MODEL"),
                    cond,
                    ColStr(row, "SOURCE_BATCH"),
                    obs,
                    If(flag, "Yes", "No"),
                    flag
                )
            Next

            AplicarFiltro()

            _lblAnaliseTitle.Visible  = True
            _chkShowIssuesOnly.Visible = True
            _dgvResultado.Visible     = True
            _lblResultadoFooter.Visible = True

        Catch ex As Exception
            MessageBox.Show("Error loading analysis grid: " & ex.Message, "Warning",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub

    Private Sub AplicarFiltro()
        If _dtResultado Is Nothing Then Return

        Dim dtView As DataTable
        If _chkShowIssuesOnly IsNot Nothing AndAlso _chkShowIssuesOnly.Checked Then
            dtView = _dtResultado.Clone()
            For Each row As DataRow In _dtResultado.Rows
                If CBool(row("_FLAG")) Then dtView.ImportRow(row)
            Next
        Else
            dtView = _dtResultado
        End If

        _dgvResultado.DataSource = dtView
        If _dgvResultado.Columns.Contains("_FLAG") Then
            _dgvResultado.Columns("_FLAG").Visible = False
        End If

        Dim total  As Integer = _dtResultado.Rows.Count
        Dim issues As Integer = 0
        For Each r As DataRow In _dtResultado.Rows
            If CBool(r("_FLAG")) Then issues += 1
        Next
        Dim ok As Integer = total - issues
        _lblResultadoFooter.Text = $"Total: {total}  |  OK: {ok}  |  Issues: {issues}"
        _lblResultadoFooter.ForeColor = If(issues > 0, TemaEscuro.Vermelho, TemaEscuro.Verde)
    End Sub

    Private Sub OnFiltroChanged(sender As Object, e As EventArgs)
        AplicarFiltro()
    End Sub

    Private Sub OnGridCellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs)
        If e.RowIndex < 0 Then Return
        If _dgvResultado Is Nothing OrElse _dgvResultado.DataSource Is Nothing Then Return
        Dim dtCurrent As DataTable = TryCast(_dgvResultado.DataSource, DataTable)
        If dtCurrent Is Nothing OrElse e.RowIndex >= dtCurrent.Rows.Count Then Return
        If Not dtCurrent.Columns.Contains("_FLAG") Then Return
        Try
            If CBool(dtCurrent.Rows(e.RowIndex)("_FLAG")) Then
                e.CellStyle.BackColor          = Color.FromArgb(255, 204, 204)
                e.CellStyle.ForeColor          = Color.FromArgb(140, 20, 20)
                e.CellStyle.SelectionBackColor = Color.FromArgb(255, 160, 160)
                e.CellStyle.SelectionForeColor = Color.FromArgb(80, 0, 0)
            End If
        Catch
        End Try
    End Sub

    Private Function ColStr(row As DataRow, colName As String) As String
        If Not row.Table.Columns.Contains(colName) Then Return ""
        If IsDBNull(row(colName)) Then Return ""
        Return row(colName).ToString()
    End Function

    Private Function BuildInClause(uids As List(Of String)) As String
        If uids.Count = 0 Then Return "1=0"
        Dim sb As New System.Text.StringBuilder()
        Dim i As Integer = 0
        While i < uids.Count
            If i > 0 Then sb.Append(" OR ")
            sb.Append("INTERNAL_UID IN (")
            Dim first As Boolean = True
            Dim limit As Integer = Math.Min(i + 999, uids.Count)
            While i < limit
                If Not first Then sb.Append(",")
                sb.Append("'").Append(uids(i).Replace("'", "''")).Append("'")
                first = False
                i += 1
            End While
            sb.Append(")")
        End While
        Return sb.ToString()
    End Function

#End Region

End Class
