Imports System.Drawing
Imports System.Windows.Forms

Partial Class frmRemessa
    Inherits Form

    Friend WithEvents lblTitulo As Label
    Friend WithEvents lblTotal As Label
    Friend WithEvents dgvRemessas As DataGridView
    Friend WithEvents lblSelecionada As Label

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

    Friend WithEvents txtBuscarEquip As TextBox
    Friend WithEvents btnBuscarEquip As Button
    Friend WithEvents dgvResultadoBusca As DataGridView
    Friend WithEvents txtPrecoItem As TextBox
    Friend WithEvents txtNotaItem As TextBox
    Friend WithEvents btnAdicionarItem As Button
    Friend WithEvents dgvItensSelecionados As DataGridView
    Friend WithEvents btnRemoverItem As Button
    Friend WithEvents lblContadorItens As Label

    Friend WithEvents cboNovoStatus As ComboBox
    Friend WithEvents btnAtualizarStatus As Button
    Friend WithEvents btnAbrirRastreio As Button
    Friend WithEvents btnFechar As Button

    Private Sub InitializeComponent()

        Me.SuspendLayout()

        Me.Text = "Shipments - FedEx / UPS / USPS / DHL"
        Me.Size = New Size(1460, 860)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.BackColor = TemaEscuro.Fundo
        Me.Font = New Font("Segoe UI", 9)

        Me.lblTitulo = New Label()
        Me.lblTitulo.Text = "Shipment Management"
        Me.lblTitulo.Font = New Font("Segoe UI", 14, FontStyle.Bold)
        Me.lblTitulo.ForeColor = TemaEscuro.Accent
        Me.lblTitulo.Location = New Point(20, 15)
        Me.lblTitulo.AutoSize = True

        Me.lblTotal = New Label()
        Me.lblTotal.Text = "Active total: 0"
        Me.lblTotal.Location = New Point(260, 22)
        Me.lblTotal.AutoSize = True
        Me.lblTotal.ForeColor = TemaEscuro.TextoMutado

        Me.dgvRemessas = New DataGridView()
        Me.dgvRemessas.Location = New Point(15, 55)
        Me.dgvRemessas.Size = New Size(700, 430)
        Me.dgvRemessas.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Bottom

        Me.lblSelecionada = New Label()
        Me.lblSelecionada.Text = "No shipment selected"
        Me.lblSelecionada.Location = New Point(15, 495)
        Me.lblSelecionada.Size = New Size(700, 22)
        Me.lblSelecionada.ForeColor = TemaEscuro.Accent
        Me.lblSelecionada.Anchor = AnchorStyles.Left Or AnchorStyles.Bottom

        Dim pnlCriar As New Panel() With {
            .Location = New Point(730, 55),
            .Size = New Size(700, 640),
            .BackColor = TemaEscuro.Surface,
            .Anchor = AnchorStyles.Top Or AnchorStyles.Right Or AnchorStyles.Bottom
        }
        AddHandler pnlCriar.Paint, Sub(sender, e)
                                       Using p As New Pen(TemaEscuro.Accent, 2)
                                           e.Graphics.DrawRectangle(p, 0, 0, pnlCriar.Width - 1, pnlCriar.Height - 1)
                                       End Using
                                   End Sub

        Dim lblCab As New Label() With {
            .Text = "New Shipment",
            .Location = New Point(15, 12),
            .AutoSize = True,
            .ForeColor = TemaEscuro.Accent,
            .Font = New Font("Segoe UI", 11, FontStyle.Bold)
        }

        Dim lblDir As New Label() With {.Text = "Direction", .Location = New Point(15, 45), .AutoSize = True, .ForeColor = TemaEscuro.TextoMutado}
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
        Me.txtTracking.Size = New Size(350, 24)

        Dim lblRec As New Label() With {.Text = "Recipient *", .Location = New Point(15, 100), .AutoSize = True, .ForeColor = TemaEscuro.Accent}
        Me.txtRecipientName = New TextBox()
        Me.txtRecipientName.Location = New Point(15, 120)
        Me.txtRecipientName.Size = New Size(660, 24)

        Dim lblAdr As New Label() With {.Text = "Recipient Address", .Location = New Point(15, 150), .AutoSize = True, .ForeColor = TemaEscuro.TextoMutado}
        Me.txtRecipientAddress = New TextBox()
        Me.txtRecipientAddress.Location = New Point(15, 170)
        Me.txtRecipientAddress.Size = New Size(660, 24)

        Dim lblSvc As New Label() With {.Text = "Service", .Location = New Point(15, 205), .AutoSize = True, .ForeColor = TemaEscuro.TextoMutado}
        Me.txtServiceLevel = New TextBox()
        Me.txtServiceLevel.Location = New Point(15, 225)
        Me.txtServiceLevel.Size = New Size(200, 24)

        Dim lblPeso As New Label() With {.Text = "Peso (lbs)", .Location = New Point(230, 205), .AutoSize = True, .ForeColor = TemaEscuro.TextoMutado}
        Me.txtPeso = New TextBox()
        Me.txtPeso.Location = New Point(230, 225)
        Me.txtPeso.Size = New Size(120, 24)
        Me.txtPeso.TextAlign = HorizontalAlignment.Right

        Dim lblCusto As New Label() With {.Text = "Custo (USD)", .Location = New Point(365, 205), .AutoSize = True, .ForeColor = TemaEscuro.TextoMutado}
        Me.txtCustoEnvio = New TextBox()
        Me.txtCustoEnvio.Location = New Point(365, 225)
        Me.txtCustoEnvio.Size = New Size(120, 24)
        Me.txtCustoEnvio.TextAlign = HorizontalAlignment.Right

        Dim lblNt As New Label() With {.Text = "Shipment Notes", .Location = New Point(500, 205), .AutoSize = True, .ForeColor = TemaEscuro.TextoMutado}
        Me.txtNotes = New TextBox()
        Me.txtNotes.Location = New Point(500, 225)
        Me.txtNotes.Size = New Size(175, 24)

        Dim lblBusca As New Label() With {.Text = "Search equipment (UID / Serial / Model)", .Location = New Point(15, 260), .AutoSize = True, .ForeColor = TemaEscuro.TextoMutado}
        Me.txtBuscarEquip = New TextBox()
        Me.txtBuscarEquip.Location = New Point(15, 280)
        Me.txtBuscarEquip.Size = New Size(520, 24)

        Me.btnBuscarEquip = New Button()
        Me.btnBuscarEquip.Text = "Search"
        Me.btnBuscarEquip.Location = New Point(545, 278)
        Me.btnBuscarEquip.Size = New Size(130, 28)

        Me.dgvResultadoBusca = New DataGridView()
        Me.dgvResultadoBusca.Location = New Point(15, 312)
        Me.dgvResultadoBusca.Size = New Size(660, 130)
        Me.dgvResultadoBusca.ReadOnly = True
        Me.dgvResultadoBusca.MultiSelect = False
        Me.dgvResultadoBusca.SelectionMode = DataGridViewSelectionMode.FullRowSelect

        Dim lblPrecoItem As New Label() With {.Text = "Item Price (USD)", .Location = New Point(15, 448), .AutoSize = True, .ForeColor = TemaEscuro.TextoMutado}
        Me.txtPrecoItem = New TextBox()
        Me.txtPrecoItem.Location = New Point(15, 468)
        Me.txtPrecoItem.Size = New Size(130, 24)
        Me.txtPrecoItem.TextAlign = HorizontalAlignment.Right

        Dim lblNotaItem As New Label() With {.Text = "Item Note", .Location = New Point(160, 448), .AutoSize = True, .ForeColor = TemaEscuro.TextoMutado}
        Me.txtNotaItem = New TextBox()
        Me.txtNotaItem.Location = New Point(160, 468)
        Me.txtNotaItem.Size = New Size(365, 24)

        Me.btnAdicionarItem = New Button()
        Me.btnAdicionarItem.Text = "Add Item"
        Me.btnAdicionarItem.Location = New Point(535, 466)
        Me.btnAdicionarItem.Size = New Size(140, 28)

        Me.lblContadorItens = New Label()
        Me.lblContadorItens.Text = "Items: 0"
        Me.lblContadorItens.Location = New Point(15, 500)
        Me.lblContadorItens.AutoSize = True
        Me.lblContadorItens.ForeColor = TemaEscuro.Accent

        Me.dgvItensSelecionados = New DataGridView()
        Me.dgvItensSelecionados.Location = New Point(15, 520)
        Me.dgvItensSelecionados.Size = New Size(660, 70)
        Me.dgvItensSelecionados.ReadOnly = True
        Me.dgvItensSelecionados.MultiSelect = False
        Me.dgvItensSelecionados.SelectionMode = DataGridViewSelectionMode.FullRowSelect

        Me.btnRemoverItem = New Button()
        Me.btnRemoverItem.Text = "Remove Item"
        Me.btnRemoverItem.Location = New Point(385, 600)
        Me.btnRemoverItem.Size = New Size(140, 30)

        Me.btnNovaRemessa = New Button()
        Me.btnNovaRemessa.Text = "Create Shipment"
        Me.btnNovaRemessa.Location = New Point(535, 600)
        Me.btnNovaRemessa.Size = New Size(140, 30)
        Me.btnNovaRemessa.Font = New Font("Segoe UI", 10, FontStyle.Bold)

        pnlCriar.Controls.AddRange({lblCab, lblDir, cboDirection, lblCar, cboCarrier,
                                    lblTrk, txtTracking, lblRec, txtRecipientName,
                                    lblAdr, txtRecipientAddress, lblSvc, txtServiceLevel,
                                    lblPeso, txtPeso, lblCusto, txtCustoEnvio,
                                    lblNt, txtNotes, lblBusca, txtBuscarEquip, btnBuscarEquip,
                                    dgvResultadoBusca, lblPrecoItem, txtPrecoItem, lblNotaItem,
                                    txtNotaItem, btnAdicionarItem, lblContadorItens,
                                    dgvItensSelecionados, btnRemoverItem, btnNovaRemessa})

        Dim pnlAcoes As New Panel() With {
            .Location = New Point(15, 670),
            .Size = New Size(1415, 100),
            .BackColor = TemaEscuro.Surface,
            .Anchor = AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Bottom
        }
        AddHandler pnlAcoes.Paint, Sub(sender, e)
                                       Using p As New Pen(TemaEscuro.Borda, 1)
                                           e.Graphics.DrawRectangle(p, 0, 0, pnlAcoes.Width - 1, pnlAcoes.Height - 1)
                                       End Using
                                   End Sub

        Dim lblAcoes As New Label() With {
            .Text = "Actions on selected shipment",
            .Location = New Point(15, 10),
            .AutoSize = True,
            .Font = New Font("Segoe UI", 10, FontStyle.Bold),
            .ForeColor = TemaEscuro.Accent
        }

        Dim lblNovoSt As New Label() With {.Text = "New Status:", .Location = New Point(15, 45), .AutoSize = True, .ForeColor = TemaEscuro.TextoMutado}
        Me.cboNovoStatus = New ComboBox()
        Me.cboNovoStatus.Location = New Point(95, 42)
        Me.cboNovoStatus.Size = New Size(200, 24)
        Me.cboNovoStatus.DropDownStyle = ComboBoxStyle.DropDownList

        Me.btnAtualizarStatus = New Button()
        Me.btnAtualizarStatus.Text = "Update Status"
        Me.btnAtualizarStatus.Location = New Point(310, 40)
        Me.btnAtualizarStatus.Size = New Size(150, 30)
        Me.btnAtualizarStatus.Enabled = False

        Me.btnAbrirRastreio = New Button()
        Me.btnAbrirRastreio.Text = "Open Tracking Web"
        Me.btnAbrirRastreio.Location = New Point(480, 40)
        Me.btnAbrirRastreio.Size = New Size(160, 30)
        Me.btnAbrirRastreio.Enabled = False

        Me.btnFechar = New Button()
        Me.btnFechar.Text = "Close"
        Me.btnFechar.Location = New Point(1285, 40)
        Me.btnFechar.Size = New Size(110, 30)
        Me.btnFechar.Anchor = AnchorStyles.Right Or AnchorStyles.Top

        pnlAcoes.Controls.AddRange({lblAcoes, lblNovoSt, cboNovoStatus, btnAtualizarStatus, btnAbrirRastreio, btnFechar})

        Me.Controls.AddRange({lblTitulo, lblTotal, dgvRemessas, lblSelecionada, pnlCriar, pnlAcoes})

        Me.ResumeLayout(False)

    End Sub

End Class
