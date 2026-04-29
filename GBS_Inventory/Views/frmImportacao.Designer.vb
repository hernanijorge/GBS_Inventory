Imports System.Drawing
Imports System.Windows.Forms

Partial Class frmImportacao
    Inherits Form

    Friend WithEvents lblTitulo As Label
    Friend WithEvents btnSelecionar As Button
    Friend WithEvents btnImportar As Button
    Friend WithEvents btnFechar As Button
    Friend WithEvents lblArquivo As Label
    Friend WithEvents lblStatus As Label
    Friend WithEvents progress As ProgressBar
    Friend WithEvents lblInseridos As Label
    Friend WithEvents lblAtualizados As Label
    Friend WithEvents lblAbas As Label
    Friend WithEvents lblLinhas As Label
    Friend WithEvents lblErros As Label
    Friend WithEvents txtErros As TextBox

    Private Sub InitializeComponent()

        Me.SuspendLayout()

        Me.Text = "Import Excel Spreadsheet — GBS"
        Me.Size = New Size(720, 620)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.BackColor = TemaEscuro.Fundo
        Me.Font = New Font("Segoe UI", 9)

        Me.lblTitulo = New Label()
        Me.lblTitulo.Text = "Spreadsheet Import"
        Me.lblTitulo.Font = New Font("Segoe UI", 14, FontStyle.Bold)
        Me.lblTitulo.ForeColor = TemaEscuro.Accent
        Me.lblTitulo.Location = New Point(20, 15)
        Me.lblTitulo.AutoSize = True

        Dim lblArq As New Label() With {
            .Text = "File:",
            .Location = New Point(20, 60),
            .AutoSize = True,
            .ForeColor = TemaEscuro.TextoMutado
        }

        Me.lblArquivo = New Label()
        Me.lblArquivo.Text = "(none)"
        Me.lblArquivo.Location = New Point(75, 60)
        Me.lblArquivo.Size = New Size(500, 20)
        Me.lblArquivo.ForeColor = TemaEscuro.TextoMutado
        Me.lblArquivo.Font = New Font("Segoe UI", 9, FontStyle.Italic)

        Me.btnSelecionar = New Button()
        Me.btnSelecionar.Text = "Select Spreadsheet..."
        Me.btnSelecionar.Location = New Point(20, 90)
        Me.btnSelecionar.Size = New Size(180, 32)

        Me.btnImportar = New Button()
        Me.btnImportar.Text = "Start Import"
        Me.btnImportar.Location = New Point(210, 90)
        Me.btnImportar.Size = New Size(180, 32)
        Me.btnImportar.Enabled = False

        Me.btnFechar = New Button()
        Me.btnFechar.Text = "Close"
        Me.btnFechar.Location = New Point(560, 90)
        Me.btnFechar.Size = New Size(120, 32)

        Me.progress = New ProgressBar()
        Me.progress.Location = New Point(20, 140)
        Me.progress.Size = New Size(660, 18)
        Me.progress.Visible = False
        Me.progress.Style = ProgressBarStyle.Continuous

        Me.lblStatus = New Label()
        Me.lblStatus.Text = "Waiting..."
        Me.lblStatus.Location = New Point(20, 165)
        Me.lblStatus.Size = New Size(660, 20)
        Me.lblStatus.ForeColor = TemaEscuro.TextoMutado

        ' Card de resultados
        Dim pnlResult As New Panel() With {
            .Location = New Point(20, 200),
            .Size = New Size(660, 160),
            .BackColor = TemaEscuro.Surface
        }
        AddHandler pnlResult.Paint, Sub(sender, e)
                                        Using p As New Pen(TemaEscuro.Accent, 2)
                                            e.Graphics.DrawRectangle(p, 0, 0, pnlResult.Width - 1, pnlResult.Height - 1)
                                        End Using
                                    End Sub

        Dim lblCab As New Label() With {
            .Text = "Results",
            .Location = New Point(15, 12),
            .AutoSize = True,
            .ForeColor = TemaEscuro.Accent,
            .Font = New Font("Segoe UI", 10, FontStyle.Bold)
        }

        Me.lblInseridos = criarInfo(15, 45, "Inserted: —", TemaEscuro.Accent)
        Me.lblAtualizados = criarInfo(15, 70, "Updated: —", TemaEscuro.Texto)
        Me.lblAbas = criarInfo(15, 95, "Sheets processed: —", TemaEscuro.Texto)
        Me.lblLinhas = criarInfo(15, 120, "Lines read: —", TemaEscuro.Texto)
        Me.lblErros = criarInfo(330, 45, "Errors: 0", TemaEscuro.TextoMutado)

        pnlResult.Controls.AddRange({lblCab, lblInseridos, lblAtualizados,
                                      lblAbas, lblLinhas, lblErros})

        ' TextBox de erros
        Dim lblDetErros As New Label() With {
            .Text = "Error details:",
            .Location = New Point(20, 375),
            .AutoSize = True,
            .ForeColor = TemaEscuro.TextoMutado
        }

        Me.txtErros = New TextBox()
        Me.txtErros.Location = New Point(20, 400)
        Me.txtErros.Size = New Size(660, 160)
        Me.txtErros.Multiline = True
        Me.txtErros.ScrollBars = ScrollBars.Vertical
        Me.txtErros.ReadOnly = True
        Me.txtErros.Font = New Font("Consolas", 8.5F)
        Me.txtErros.Visible = False

        Me.Controls.AddRange({lblTitulo, lblArq, lblArquivo, btnSelecionar, btnImportar,
                               btnFechar, progress, lblStatus, pnlResult, lblDetErros, txtErros})

        Me.ResumeLayout(False)

    End Sub

    Private Function criarInfo(pX As Integer, pY As Integer, pText As String, pCor As Color) As Label
        Return New Label() With {
            .Location = New Point(pX, pY),
            .Size = New Size(310, 22),
            .Text = pText,
            .ForeColor = pCor,
            .Font = New Font("Segoe UI", 10)
        }
    End Function

End Class
