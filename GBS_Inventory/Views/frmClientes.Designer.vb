Imports System.Drawing
Imports System.Windows.Forms

Partial Class frmClientes
    Inherits Form

    Friend WithEvents lblTitulo As Label
    Friend WithEvents txtPesquisa As TextBox
    Friend WithEvents btnPesquisar As Button
    Friend WithEvents dgvClientes As DataGridView
    Friend WithEvents txtNomeRazao As TextBox
    Friend WithEvents txtNomeFantasia As TextBox
    Friend WithEvents txtDocumento As TextBox
    Friend WithEvents txtEmail As TextBox
    Friend WithEvents txtTelefone As TextBox
    Friend WithEvents txtEndereco1 As TextBox
    Friend WithEvents txtEndereco2 As TextBox
    Friend WithEvents txtCidade As TextBox
    Friend WithEvents txtEstado As TextBox
    Friend WithEvents txtZipCode As TextBox
    Friend WithEvents txtPais As TextBox
    Friend WithEvents txtObservacoes As TextBox
    Friend WithEvents chkAtivo As CheckBox
    Friend WithEvents btnNovo As Button
    Friend WithEvents btnSalvar As Button
    Friend WithEvents btnExcluir As Button
    Friend WithEvents btnAtualizar As Button
    Friend WithEvents btnFechar As Button
    Friend WithEvents lblSelecionado As Label

    Private Sub InitializeComponent()

        Me.SuspendLayout()

        Me.Text = "Customer Registration"
        Me.Size = New Size(1280, 760)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.BackColor = TemaEscuro.Fundo
        Me.Font = New Font("Segoe UI", 9)

        Me.lblTitulo = New Label()
        Me.lblTitulo.Text = "Customers"
        Me.lblTitulo.Font = New Font("Segoe UI", 14, FontStyle.Bold)
        Me.lblTitulo.ForeColor = TemaEscuro.Accent
        Me.lblTitulo.Location = New Point(20, 12)
        Me.lblTitulo.AutoSize = True

        Me.txtPesquisa = New TextBox()
        Me.txtPesquisa.Location = New Point(20, 50)
        Me.txtPesquisa.Size = New Size(300, 24)

        Me.btnPesquisar = New Button()
        Me.btnPesquisar.Text = "Search"
        Me.btnPesquisar.Location = New Point(330, 48)
        Me.btnPesquisar.Size = New Size(100, 28)

        Me.dgvClientes = New DataGridView()
        Me.dgvClientes.Location = New Point(20, 90)
        Me.dgvClientes.Size = New Size(620, 560)
        Me.dgvClientes.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Bottom

        Dim pnlCadastro As New Panel() With {
            .Location = New Point(660, 90),
            .Size = New Size(590, 560),
            .BackColor = TemaEscuro.Surface,
            .Anchor = AnchorStyles.Top Or AnchorStyles.Right Or AnchorStyles.Bottom
        }

        AddHandler pnlCadastro.Paint, Sub(sender, e)
                                          Using p As New Pen(TemaEscuro.Borda, 1)
                                              e.Graphics.DrawRectangle(p, 0, 0, pnlCadastro.Width - 1, pnlCadastro.Height - 1)
                                          End Using
                                      End Sub

        Me.lblSelecionado = New Label()
        Me.lblSelecionado.Text = "New customer"
        Me.lblSelecionado.Location = New Point(15, 10)
        Me.lblSelecionado.Size = New Size(560, 22)
        Me.lblSelecionado.ForeColor = TemaEscuro.Accent
        Me.lblSelecionado.Font = New Font("Segoe UI", 10, FontStyle.Bold)

        Dim y As Integer = 40
        Dim h As Integer = 22

        pnlCadastro.Controls.Add(CriarLabel("Name / Company *", 15, y))
        Me.txtNomeRazao = New TextBox() With {.Location = New Point(15, y + 18), .Size = New Size(560, h)}
        pnlCadastro.Controls.Add(txtNomeRazao)

        y += 50
        pnlCadastro.Controls.Add(CriarLabel("Trade Name", 15, y))
        Me.txtNomeFantasia = New TextBox() With {.Location = New Point(15, y + 18), .Size = New Size(560, h)}
        pnlCadastro.Controls.Add(txtNomeFantasia)

        y += 50
        pnlCadastro.Controls.Add(CriarLabel("Document", 15, y))
        Me.txtDocumento = New TextBox() With {.Location = New Point(15, y + 18), .Size = New Size(180, h)}
        pnlCadastro.Controls.Add(txtDocumento)

        pnlCadastro.Controls.Add(CriarLabel("Email", 210, y))
        Me.txtEmail = New TextBox() With {.Location = New Point(210, y + 18), .Size = New Size(220, h)}
        pnlCadastro.Controls.Add(txtEmail)

        pnlCadastro.Controls.Add(CriarLabel("Phone", 440, y))
        Me.txtTelefone = New TextBox() With {.Location = New Point(440, y + 18), .Size = New Size(135, h)}
        pnlCadastro.Controls.Add(txtTelefone)

        y += 50
        pnlCadastro.Controls.Add(CriarLabel("Address 1", 15, y))
        Me.txtEndereco1 = New TextBox() With {.Location = New Point(15, y + 18), .Size = New Size(560, h)}
        pnlCadastro.Controls.Add(txtEndereco1)

        y += 50
        pnlCadastro.Controls.Add(CriarLabel("Address 2", 15, y))
        Me.txtEndereco2 = New TextBox() With {.Location = New Point(15, y + 18), .Size = New Size(560, h)}
        pnlCadastro.Controls.Add(txtEndereco2)

        y += 50
        Dim lblCity As New Label() With {.Text = "City", .Location = New Point(15, y), .AutoSize = True, .ForeColor = TemaEscuro.TextoMutado}
        pnlCadastro.Controls.Add(lblCity)
        Me.txtCidade = New TextBox() With {.Location = New Point(15, y + 18), .Size = New Size(210, h)}
        pnlCadastro.Controls.Add(txtCidade)

        Dim lblState As New Label() With {
            .Text = "State",
            .Location = New Point(235, y),
            .AutoSize = True,
            .ForeColor = TemaEscuro.Accent,
            .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold)
        }
        pnlCadastro.Controls.Add(lblState)
        Me.txtEstado = New TextBox() With {.Location = New Point(235, y + 18), .Size = New Size(90, h)}
        pnlCadastro.Controls.Add(txtEstado)

        Dim lblZip As New Label() With {
            .Text = "ZIP Code",
            .Location = New Point(335, y),
            .AutoSize = True,
            .ForeColor = TemaEscuro.Accent,
            .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold)
        }
        pnlCadastro.Controls.Add(lblZip)
        Me.txtZipCode = New TextBox() With {.Location = New Point(335, y + 18), .Size = New Size(120, h)}
        pnlCadastro.Controls.Add(txtZipCode)

        Dim lblCountry As New Label() With {
            .Text = "Country",
            .Location = New Point(465, y),
            .AutoSize = True,
            .ForeColor = TemaEscuro.Accent,
            .Font = New Font("Segoe UI", 8.5F, FontStyle.Bold)
        }
        pnlCadastro.Controls.Add(lblCountry)
        Me.txtPais = New TextBox() With {.Location = New Point(465, y + 18), .Size = New Size(110, h), .Text = "USA"}
        pnlCadastro.Controls.Add(txtPais)

        y += 50
        pnlCadastro.Controls.Add(CriarLabel("Notes", 15, y))
        Me.txtObservacoes = New TextBox() With {
            .Location = New Point(15, y + 18),
            .Size = New Size(560, 120),
            .Multiline = True,
            .ScrollBars = ScrollBars.Vertical
        }
        pnlCadastro.Controls.Add(txtObservacoes)

        Me.chkAtivo = New CheckBox() With {
            .Text = "Active customer",
            .Location = New Point(15, 430),
            .Size = New Size(120, 24),
            .Checked = True
        }
        pnlCadastro.Controls.Add(chkAtivo)

        Me.btnNovo = New Button() With {.Text = "New", .Location = New Point(15, 475), .Size = New Size(90, 34)}
        Me.btnSalvar = New Button() With {.Text = "Save", .Location = New Point(115, 475), .Size = New Size(100, 34)}
        Me.btnExcluir = New Button() With {.Text = "Delete", .Location = New Point(225, 475), .Size = New Size(100, 34)}
        Me.btnAtualizar = New Button() With {.Text = "Refresh", .Location = New Point(335, 475), .Size = New Size(100, 34)}
        Me.btnFechar = New Button() With {.Text = "Close", .Location = New Point(475, 475), .Size = New Size(100, 34)}

        pnlCadastro.Controls.AddRange({btnNovo, btnSalvar, btnExcluir, btnAtualizar, btnFechar, lblSelecionado})

        Me.Controls.AddRange({lblTitulo, txtPesquisa, btnPesquisar, dgvClientes, pnlCadastro})

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
