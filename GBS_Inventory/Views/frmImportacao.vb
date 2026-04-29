Imports System.Windows.Forms
Imports System.Drawing
Imports System.IO

''' <summary>
''' Tela de Importação de Planilha Excel da GBS.
''' </summary>
Public Class frmImportacao

    Private oController As ImportacaoController
    Private sArquivoSelecionado As String = ""

    Public Sub New()
        InitializeComponent()
        oController = New ImportacaoController()
        TemaEscuro.aplicarHelius(Me)
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

            If oController.TotalErros > 0 Then
                txtErros.Lines = oController.Erros.ToArray()
                txtErros.Visible = True
                lblErros.ForeColor = TemaEscuro.Vermelho
            End If

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

End Class
