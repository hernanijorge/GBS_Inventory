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
        TemaEscuro.aplicar(Me)
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
            MessageBox.Show("Selecione uma planilha primeiro.", "Atenção",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        btnImportar.Enabled   = False
        btnSelecionar.Enabled = False
        btnFechar.Enabled     = False

        lblStatus.Text        = "Importando..."
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
            lblStatus.Text     = "Concluído!"
            lblStatus.ForeColor= TemaEscuro.Accent

            lblInseridos.Text   = "Inseridos: "    & oController.TotalInseridos.ToString()
            lblAtualizados.Text = "Atualizados: "  & oController.TotalAtualizados.ToString()
            lblAbas.Text        = "Abas processadas: " & oController.TotalAbas.ToString()
            lblLinhas.Text      = "Linhas lidas: " & oController.TotalLinhasPlanilha.ToString()
            lblErros.Text       = "Erros: "        & oController.TotalErros.ToString()

            If oController.TotalErros > 0 Then
                txtErros.Lines = oController.Erros.ToArray()
                txtErros.Visible = True
                lblErros.ForeColor = TemaEscuro.Vermelho
            End If

            MessageBox.Show($"Importação concluída!{vbCrLf & vbCrLf}" &
                            $"Inseridos: {oController.TotalInseridos}{vbCrLf}" &
                            $"Atualizados: {oController.TotalAtualizados}{vbCrLf}" &
                            $"Erros: {oController.TotalErros}",
                            "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception

            MessageBox.Show("Erro na importação: " & ex.Message, "Erro",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
            lblStatus.Text     = "Falha na importação"
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
