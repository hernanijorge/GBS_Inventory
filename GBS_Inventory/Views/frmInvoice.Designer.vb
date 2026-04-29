Imports System.Drawing
Imports System.Windows.Forms

Partial Class frmInvoice
    Inherits Form

    Friend WithEvents lblTitulo As Label
    Friend WithEvents cboCliente As ComboBox
    Friend WithEvents cboRemessa As ComboBox
    Friend WithEvents txtInvoiceNumber As TextBox
    Friend WithEvents dtIssue As DateTimePicker
    Friend WithEvents dtDue As DateTimePicker
    Friend WithEvents txtEmail As TextBox
    Friend WithEvents txtDiscount As TextBox
    Friend WithEvents txtShipping As TextBox
    Friend WithEvents txtTax As TextBox
    Friend WithEvents txtLogo As TextBox
    Friend WithEvents txtOutput As TextBox
    Friend WithEvents txtNotes As TextBox
    Friend WithEvents btnCarregarItens As Button
    Friend WithEvents btnSelecionarLogo As Button
    Friend WithEvents btnSelecionarPasta As Button
    Friend WithEvents btnGerar As Button
    Friend WithEvents btnEnviar As Button
    Friend WithEvents btnFechar As Button
    Friend WithEvents chkEnviarSomentePdf As CheckBox
    Friend WithEvents dgvItens As DataGridView
    Friend WithEvents lblSubtotal As Label
    Friend WithEvents lblTotal As Label

    Private Sub InitializeComponent()

        Me.SuspendLayout()

        Me.Text = "Invoice"
        Me.Size = New Size(1320, 780)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.BackColor = TemaEscuro.Fundo
        Me.Font = New Font("Segoe UI", 9)

        Me.lblTitulo = New Label()
        Me.lblTitulo.Text = "Invoice (PDF + Word + Email)"
        Me.lblTitulo.Location = New Point(20, 12)
        Me.lblTitulo.AutoSize = True
        Me.lblTitulo.Font = New Font("Segoe UI", 14, FontStyle.Bold)
        Me.lblTitulo.ForeColor = TemaEscuro.Accent

        Dim pnlTopo As New Panel() With {
            .Location = New Point(20, 45),
            .Size = New Size(1260, 200),
            .BackColor = TemaEscuro.Surface
        }

        AddHandler pnlTopo.Paint, Sub(sender, e)
                                      Using p As New Pen(TemaEscuro.Borda, 1)
                                          e.Graphics.DrawRectangle(p, 0, 0, pnlTopo.Width - 1, pnlTopo.Height - 1)
                                      End Using
                                  End Sub

        pnlTopo.Controls.Add(CriarLabel("Customer *", 15, 14))
        Me.cboCliente = New ComboBox() With {.Location = New Point(15, 34), .Size = New Size(330, 24), .DropDownStyle = ComboBoxStyle.DropDownList}
        pnlTopo.Controls.Add(cboCliente)

        pnlTopo.Controls.Add(CriarLabel("Active Shipment *", 360, 14))
        Me.cboRemessa = New ComboBox() With {.Location = New Point(360, 34), .Size = New Size(240, 24), .DropDownStyle = ComboBoxStyle.DropDownList}
        pnlTopo.Controls.Add(cboRemessa)

        Me.btnCarregarItens = New Button() With {.Text = "Load Items", .Location = New Point(615, 31), .Size = New Size(130, 30)}
        pnlTopo.Controls.Add(btnCarregarItens)

        Dim lblInvoiceNumber As New Label() With {.Text = "Invoice Number", .Location = New Point(760, 14), .AutoSize = True, .ForeColor = TemaEscuro.TextoMutado}
        pnlTopo.Controls.Add(lblInvoiceNumber)
        Me.txtInvoiceNumber = New TextBox() With {.Location = New Point(760, 34), .Size = New Size(180, 24), .ReadOnly = True}
        pnlTopo.Controls.Add(txtInvoiceNumber)

        Dim lblIssue As New Label() With {.Text = "Issue Date", .Location = New Point(955, 14), .AutoSize = True, .ForeColor = TemaEscuro.Accent, .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold)}
        pnlTopo.Controls.Add(lblIssue)
        Me.dtIssue = New DateTimePicker() With {.Location = New Point(955, 34), .Size = New Size(140, 24), .Format = DateTimePickerFormat.Short, .Name = "dtIssue"}
        pnlTopo.Controls.Add(dtIssue)

        Dim lblDue As New Label() With {.Text = "Due Date", .Location = New Point(1110, 14), .AutoSize = True, .ForeColor = TemaEscuro.Accent, .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold)}
        pnlTopo.Controls.Add(lblDue)
        Me.dtDue = New DateTimePicker() With {.Location = New Point(1110, 34), .Size = New Size(130, 24), .Format = DateTimePickerFormat.Short, .Name = "dtDue"}
        pnlTopo.Controls.Add(dtDue)

        pnlTopo.Controls.Add(CriarLabel("Customer Email", 15, 70))
        Me.txtEmail = New TextBox() With {.Location = New Point(15, 90), .Size = New Size(330, 24)}
        pnlTopo.Controls.Add(txtEmail)

        Dim lblDiscount As New Label() With {.Text = "Discount USD", .Location = New Point(360, 70), .AutoSize = True, .ForeColor = TemaEscuro.Accent, .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold)}
        pnlTopo.Controls.Add(lblDiscount)
        Me.txtDiscount = New TextBox() With {.Location = New Point(360, 90), .Size = New Size(110, 24), .Text = "0.00", .TextAlign = HorizontalAlignment.Right}
        pnlTopo.Controls.Add(txtDiscount)

        Dim lblShipping As New Label() With {.Text = "Shipping USD", .Location = New Point(485, 70), .AutoSize = True, .ForeColor = TemaEscuro.Accent, .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold)}
        pnlTopo.Controls.Add(lblShipping)
        Me.txtShipping = New TextBox() With {.Location = New Point(485, 90), .Size = New Size(110, 24), .Text = "0.00", .TextAlign = HorizontalAlignment.Right}
        pnlTopo.Controls.Add(txtShipping)

        Dim lblTax As New Label() With {.Text = "Tax USD", .Location = New Point(610, 70), .AutoSize = True, .ForeColor = TemaEscuro.Accent, .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold)}
        pnlTopo.Controls.Add(lblTax)
        Me.txtTax = New TextBox() With {.Location = New Point(610, 90), .Size = New Size(110, 24), .Text = "0.00", .TextAlign = HorizontalAlignment.Right}
        pnlTopo.Controls.Add(txtTax)

        Dim lblLogo As New Label() With {.Text = "Logo", .Location = New Point(735, 70), .AutoSize = True, .ForeColor = TemaEscuro.Accent, .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold)}
        pnlTopo.Controls.Add(lblLogo)
        Me.txtLogo = New TextBox() With {.Location = New Point(735, 90), .Size = New Size(420, 24)}
        pnlTopo.Controls.Add(txtLogo)

        Me.btnSelecionarLogo = New Button() With {.Text = "...", .Location = New Point(1162, 89), .Size = New Size(38, 26)}
        pnlTopo.Controls.Add(btnSelecionarLogo)

        pnlTopo.Controls.Add(CriarLabel("Output Folder", 15, 126))
        Me.txtOutput = New TextBox() With {.Location = New Point(15, 146), .Size = New Size(700, 24)}
        pnlTopo.Controls.Add(txtOutput)

        Me.btnSelecionarPasta = New Button() With {.Text = "...", .Location = New Point(722, 145), .Size = New Size(38, 26)}
        pnlTopo.Controls.Add(btnSelecionarPasta)

        pnlTopo.Controls.Add(CriarLabel("Notes", 775, 126))
        Me.txtNotes = New TextBox() With {.Location = New Point(775, 146), .Size = New Size(465, 24)}
        pnlTopo.Controls.Add(txtNotes)

        Me.dgvItens = New DataGridView()
        Me.dgvItens.Location = New Point(20, 260)
        Me.dgvItens.Size = New Size(1260, 380)
        Me.dgvItens.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Bottom

        Dim pnlRodape As New Panel() With {
            .Location = New Point(20, 650),
            .Size = New Size(1260, 70),
            .BackColor = TemaEscuro.Surface,
            .Anchor = AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Bottom
        }

        AddHandler pnlRodape.Paint, Sub(sender, e)
                                        Using p As New Pen(TemaEscuro.Borda, 1)
                                            e.Graphics.DrawRectangle(p, 0, 0, pnlRodape.Width - 1, pnlRodape.Height - 1)
                                        End Using
                                    End Sub

        Me.lblSubtotal = New Label() With {
            .Text = "Subtotal: 0.00",
            .Location = New Point(16, 12),
            .AutoSize = True,
            .ForeColor = TemaEscuro.Texto
        }

        Me.lblTotal = New Label() With {
            .Text = "Total: 0.00",
            .Location = New Point(16, 36),
            .AutoSize = True,
            .ForeColor = TemaEscuro.Accent,
            .Font = New Font("Segoe UI", 10, FontStyle.Bold)
        }

        Me.chkEnviarSomentePdf = New CheckBox() With {
            .Text = "Send PDF only",
            .Location = New Point(640, 25),
            .AutoSize = True
        }

        Me.btnGerar = New Button() With {.Text = "Generate PDF + Word", .Location = New Point(830, 18), .Size = New Size(150, 34)}
        Me.btnEnviar = New Button() With {.Text = "Send Email", .Location = New Point(990, 18), .Size = New Size(120, 34)}
        Me.btnFechar = New Button() With {.Text = "Close", .Location = New Point(1120, 18), .Size = New Size(120, 34)}

        pnlRodape.Controls.AddRange({lblSubtotal, lblTotal, chkEnviarSomentePdf, btnGerar, btnEnviar, btnFechar})

        Me.Controls.AddRange({lblTitulo, pnlTopo, dgvItens, pnlRodape})

        Me.ResumeLayout(False)

    End Sub

    Private Function CriarLabel(pTexto As String, pX As Integer, pY As Integer) As Label

        Return New Label() With {
            .Text = pTexto,
            .Location = New Point(pX, pY),
            .Size = New Size(240, 18),
            .ForeColor = TemaEscuro.TextoMutado
        }

    End Function

End Class
