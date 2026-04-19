Imports System.Drawing
Imports System.Windows.Forms

Partial Class frmRemessa
    Inherits Form

    Friend WithEvents lblTitulo As Label
    Friend WithEvents lblTotal As Label
    Friend WithEvents dgvRemessas As DataGridView
    Friend WithEvents lblSelecionada As Label

    ' Painel criar remessa
    Friend WithEvents cboDirection As ComboBox
    Friend WithEvents cboCarrier As ComboBox
    Friend WithEvents txtTracking As TextBox
    Friend WithEvents txtRecipientName As TextBox
    Friend WithEvents txtRecipientAddress As TextBox
    Friend WithEvents txtServiceLevel As TextBox
    Friend WithEvents txtPeso As TextBox
    Friend WithEvents txtCustoEnvio As TextBox
    Friend WithEvents txtNotes As TextBox
    Friend WithEvents btnNovaRemessa As Button

    ' Painel atualizar status
    Friend WithEvents cboNovoStatus As ComboBox
    Friend WithEvents btnAtualizarStatus As Button
    Friend WithEvents btnAbrirRastreio As Button
    Friend WithEvents btnFechar As Button

    Private Sub InitializeComponent()

        Me.SuspendLayout()

        Me.Text = "Remessas — FedEx · UPS · USPS · DHL"
        Me.Size = New Size(1280, 760)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.BackColor = TemaEscuro.Fundo
        Me.Font = New Font("Segoe UI", 9)

        Me.lblTitulo = New Label()
        Me.lblTitulo.Text = "Controle de Remessas"
        Me.lblTitulo.Font = New Font("Segoe UI", 14, FontStyle.Bold)
        Me.lblTitulo.ForeColor = TemaEscuro.Accent
        Me.lblTitulo.Location = New Point(20, 15)
        Me.lblTitulo.AutoSize = True

        Me.lblTotal = New Label()
        Me.lblTotal.Text = "Total ativas: 0"
        Me.lblTotal.Location = New Point(260, 22)
        Me.lblTotal.AutoSize = True
        Me.lblTotal.ForeColor = TemaEscuro.TextoMutado

        ' Grid de remessas (esquerda)
        Me.dgvRemessas = New DataGridView()
        Me.dgvRemessas.Location = New Point(15, 55)
        Me.dgvRemessas.Size = New Size(720, 480)
        Me.dgvRemessas.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Bottom

        Me.lblSelecionada = New Label()
        Me.lblSelecionada.Text = "Nenhuma remessa selecionada"
        Me.lblSelecionada.Location = New Point(15, 545)
        Me.lblSelecionada.Size = New Size(720, 22)
        Me.lblSelecionada.ForeColor = TemaEscuro.Accent
        Me.lblSelecionada.Anchor = AnchorStyles.Left Or AnchorStyles.Bottom

        ' Painel direita: Criar nova remessa
        Dim pnlCriar As New Panel() With {
            .Location = New Point(750, 55),
            .Size = New Size(500, 480),
            .BackColor = TemaEscuro.Surface,
            .Anchor = AnchorStyles.Top Or AnchorStyles.Right Or AnchorStyles.Bottom
        }
        AddHandler pnlCriar.Paint, Sub(sender, e)
                                       Using p As New Pen(TemaEscuro.Accent, 2)
                                           e.Graphics.DrawRectangle(p, 0, 0, pnlCriar.Width - 1, pnlCriar.Height - 1)
                                       End Using
                                   End Sub

        Dim lblCab As New Label() With {
            .Text = "Nova Remessa",
            .Location = New Point(15, 12),
            .AutoSize = True,
            .ForeColor = TemaEscuro.Accent,
            .Font = New Font("Segoe UI", 11, FontStyle.Bold)
        }

        Dim lblDir As New Label() With {.Text = "Direção", .Location = New Point(15, 45), .AutoSize = True, .ForeColor = TemaEscuro.TextoMutado}
        Me.cboDirection = New ComboBox()
        Me.cboDirection.Location = New Point(15, 65)
        Me.cboDirection.Size = New Size(140, 24)
        Me.cboDirection.DropDownStyle = ComboBoxStyle.DropDownList

        Dim lblCar As New Label() With {.Text = "Carrier", .Location = New Point(170, 45), .AutoSize = True, .ForeColor = TemaEscuro.TextoMutado}
        Me.cboCarrier = New ComboBox()
        Me.cboCarrier.Location = New Point(170, 65)
        Me.cboCarrier.Size = New Size(140, 24)
        Me.cboCarrier.DropDownStyle = ComboBoxStyle.DropDownList

        Dim lblTrk As New Label() With {.Text = "Tracking Number *", .Location = New Point(325, 45), .AutoSize = True, .ForeColor = TemaEscuro.Accent}
        Me.txtTracking = New TextBox()
        Me.txtTracking.Location = New Point(325, 65)
        Me.txtTracking.Size = New Size(160, 24)

        Dim lblRec As New Label() With {.Text = "Destinatário *", .Location = New Point(15, 100), .AutoSize = True, .ForeColor = TemaEscuro.Accent}
        Me.txtRecipientName = New TextBox()
        Me.txtRecipientName.Location = New Point(15, 120)
        Me.txtRecipientName.Size = New Size(470, 24)

        Dim lblAdr As New Label() With {.Text = "Endereço do Destinatário", .Location = New Point(15, 150), .AutoSize = True, .ForeColor = TemaEscuro.TextoMutado}
        Me.txtRecipientAddress = New TextBox()
        Me.txtRecipientAddress.Location = New Point(15, 170)
        Me.txtRecipientAddress.Size = New Size(470, 24)

        Dim lblSvc As New Label() With {.Text = "Serviço", .Location = New Point(15, 205), .AutoSize = True, .ForeColor = TemaEscuro.TextoMutado}
        Me.txtServiceLevel = New TextBox()
        Me.txtServiceLevel.Location = New Point(15, 225)
        Me.txtServiceLevel.Size = New Size(200, 24)

        Dim lblPeso As New Label() With {.Text = "Peso (lbs)", .Location = New Point(230, 205), .AutoSize = True, .ForeColor = TemaEscuro.TextoMutado}
        Me.txtPeso = New TextBox()
        Me.txtPeso.Location = New Point(230, 225)
        Me.txtPeso.Size = New Size(100, 24)
        Me.txtPeso.TextAlign = HorizontalAlignment.Right

        Dim lblCusto As New Label() With {.Text = "Custo (USD)", .Location = New Point(345, 205), .AutoSize = True, .ForeColor = TemaEscuro.TextoMutado}
        Me.txtCustoEnvio = New TextBox()
        Me.txtCustoEnvio.Location = New Point(345, 225)
        Me.txtCustoEnvio.Size = New Size(140, 24)
        Me.txtCustoEnvio.TextAlign = HorizontalAlignment.Right

        Dim lblNt As New Label() With {.Text = "Observações", .Location = New Point(15, 260), .AutoSize = True, .ForeColor = TemaEscuro.TextoMutado}
        Me.txtNotes = New TextBox()
        Me.txtNotes.Location = New Point(15, 280)
        Me.txtNotes.Size = New Size(470, 110)
        Me.txtNotes.Multiline = True
        Me.txtNotes.ScrollBars = ScrollBars.Vertical

        Me.btnNovaRemessa = New Button()
        Me.btnNovaRemessa.Text = "Criar Remessa"
        Me.btnNovaRemessa.Location = New Point(325, 410)
        Me.btnNovaRemessa.Size = New Size(160, 40)
        Me.btnNovaRemessa.Font = New Font("Segoe UI", 10, FontStyle.Bold)

        pnlCriar.Controls.AddRange({lblCab, lblDir, cboDirection, lblCar, cboCarrier,
                                     lblTrk, txtTracking, lblRec, txtRecipientName,
                                     lblAdr, txtRecipientAddress, lblSvc, txtServiceLevel,
                                     lblPeso, txtPeso, lblCusto, txtCustoEnvio,
                                     lblNt, txtNotes, btnNovaRemessa})

        ' Painel inferior: ações sobre remessa selecionada
        Dim pnlAcoes As New Panel() With {
            .Location = New Point(15, 580),
            .Size = New Size(1235, 100),
            .BackColor = TemaEscuro.Surface,
            .Anchor = AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Bottom
        }
        AddHandler pnlAcoes.Paint, Sub(sender, e)
                                       Using p As New Pen(TemaEscuro.Borda, 1)
                                           e.Graphics.DrawRectangle(p, 0, 0, pnlAcoes.Width - 1, pnlAcoes.Height - 1)
                                       End Using
                                   End Sub

        Dim lblAcoes As New Label() With {
            .Text = "Ações sobre remessa selecionada",
            .Location = New Point(15, 10),
            .AutoSize = True,
            .Font = New Font("Segoe UI", 10, FontStyle.Bold),
            .ForeColor = TemaEscuro.Accent
        }

        Dim lblNovoSt As New Label() With {.Text = "Novo Status:", .Location = New Point(15, 45), .AutoSize = True, .ForeColor = TemaEscuro.TextoMutado}
        Me.cboNovoStatus = New ComboBox()
        Me.cboNovoStatus.Location = New Point(95, 42)
        Me.cboNovoStatus.Size = New Size(180, 24)
        Me.cboNovoStatus.DropDownStyle = ComboBoxStyle.DropDownList

        Me.btnAtualizarStatus = New Button()
        Me.btnAtualizarStatus.Text = "Atualizar Status"
        Me.btnAtualizarStatus.Location = New Point(290, 40)
        Me.btnAtualizarStatus.Size = New Size(150, 30)
        Me.btnAtualizarStatus.Enabled = False

        Me.btnAbrirRastreio = New Button()
        Me.btnAbrirRastreio.Text = "Abrir Rastreio Web"
        Me.btnAbrirRastreio.Location = New Point(460, 40)
        Me.btnAbrirRastreio.Size = New Size(150, 30)
        Me.btnAbrirRastreio.Enabled = False

        Me.btnFechar = New Button()
        Me.btnFechar.Text = "Fechar"
        Me.btnFechar.Location = New Point(1110, 40)
        Me.btnFechar.Size = New Size(110, 30)
        Me.btnFechar.Anchor = AnchorStyles.Right Or AnchorStyles.Top

        pnlAcoes.Controls.AddRange({lblAcoes, lblNovoSt, cboNovoStatus,
                                     btnAtualizarStatus, btnAbrirRastreio, btnFechar})

        Me.Controls.AddRange({lblTitulo, lblTotal, dgvRemessas, lblSelecionada,
                               pnlCriar, pnlAcoes})

        Me.ResumeLayout(False)

    End Sub

End Class
