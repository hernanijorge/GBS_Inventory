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

    Friend WithEvents txtQuickScan       As TextBox
    Friend WithEvents lblQuickScanStatus As Label
    Friend WithEvents txtBuscarEquip     As TextBox
    Friend WithEvents btnBuscarEquip     As Button
    Friend WithEvents dgvResultadoBusca  As DataGridView
    Friend WithEvents txtPrecoItem As TextBox
    Friend WithEvents txtNotaItem As TextBox
    Friend WithEvents btnAdicionarItem As Button
    Friend WithEvents dgvItensSelecionados As DataGridView
    Friend WithEvents btnRemoverItem As Button
    Friend WithEvents lblContadorItens As Label

    Friend WithEvents cboNovoStatus As ComboBox
    Friend WithEvents btnAtualizarStatus As Button
    Friend WithEvents btnAbrirRastreio As Button
    Friend WithEvents btnRelatorio As Button
    Friend WithEvents btnFechar As Button

    Private Sub InitializeComponent()

        Me.SuspendLayout()

        Me.Text        = "Shipments - FedEx / UPS / USPS / DHL"
        Me.Size        = New Size(1400, 900)
        Me.MinimumSize = New Size(1200, 700)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.BackColor   = TemaEscuro.Fundo
        Me.Font        = New Font("Segoe UI", 9)

        ' ── pnlTop — title bar ────────────────────────────────────────────────
        Dim pnlTop As New Panel() With {
            .Dock      = DockStyle.Top,
            .Height    = 48,
            .BackColor = TemaEscuro.Fundo
        }
        Me.lblTitulo = New Label() With {
            .Text      = "Shipment Management",
            .Font      = New Font("Segoe UI", 14, FontStyle.Bold),
            .ForeColor = TemaEscuro.Accent,
            .Location  = New Point(15, 10),
            .AutoSize  = True
        }
        Me.lblTotal = New Label() With {
            .Text      = "Active total: 0",
            .Location  = New Point(260, 17),
            .AutoSize  = True,
            .ForeColor = TemaEscuro.TextoMutado
        }
        pnlTop.Controls.AddRange({lblTitulo, lblTotal})

        ' ── pnlAcoes — bottom action bar ─────────────────────────────────────
        Dim pnlAcoes As New Panel() With {
            .Dock      = DockStyle.Bottom,
            .Height    = 105,
            .BackColor = TemaEscuro.Surface
        }
        AddHandler pnlAcoes.Paint, Sub(s, e)
            Using p As New Pen(TemaEscuro.Borda, 1)
                e.Graphics.DrawRectangle(p, 0, 0, pnlAcoes.Width - 1, pnlAcoes.Height - 1)
            End Using
        End Sub

        Dim lblAcoes As New Label() With {
            .Text      = "Actions on selected shipment",
            .Location  = New Point(15, 10),
            .AutoSize  = True,
            .Font      = New Font("Segoe UI", 10, FontStyle.Bold),
            .ForeColor = TemaEscuro.Accent
        }
        Dim lblNovoSt As New Label() With {
            .Text      = "New Status:",
            .Location  = New Point(15, 48),
            .AutoSize  = True,
            .ForeColor = TemaEscuro.TextoMutado
        }
        Me.cboNovoStatus = New ComboBox() With {
            .Location      = New Point(95, 45),
            .Size          = New Size(200, 24),
            .DropDownStyle = ComboBoxStyle.DropDownList
        }
        Me.btnAtualizarStatus = New Button() With {
            .Text     = "Update Status",
            .Location = New Point(310, 43),
            .Size     = New Size(150, 30),
            .Enabled  = False
        }
        Me.btnAbrirRastreio = New Button() With {
            .Text     = "Open Tracking Web",
            .Location = New Point(475, 43),
            .Size     = New Size(160, 30),
            .Enabled  = False
        }
        Me.btnRelatorio = New Button() With {
            .Text     = "Shipment Report",
            .Location = New Point(650, 43),
            .Size     = New Size(150, 30),
            .Enabled  = False
        }
        Me.btnFechar = New Button() With {
            .Text     = "Close",
            .Location = New Point(1255, 43),
            .Size     = New Size(110, 30),
            .Anchor   = AnchorStyles.Top Or AnchorStyles.Right
        }
        pnlAcoes.Controls.AddRange({lblAcoes, lblNovoSt, cboNovoStatus,
                                     btnAtualizarStatus, btnAbrirRastreio,
                                     btnRelatorio, btnFechar})

        ' ── pnlCriar — New Shipment form (Dock=Right, fixed width=730) ────────
        ' Width is fixed — controls inside use FIXED positions, NO Anchor=Right,
        ' to avoid the Width=0-at-init-time anchor-offset bug.
        Const PANEL_W As Integer = 730

        Dim pnlCriar As New Panel() With {
            .Dock      = DockStyle.Right,
            .Width     = PANEL_W,
            .BackColor = TemaEscuro.Surface
        }
        AddHandler pnlCriar.Paint, Sub(s, e)
            Using p As New Pen(TemaEscuro.Accent, 2)
                e.Graphics.DrawRectangle(p, 0, 0, pnlCriar.Width - 1, pnlCriar.Height - 1)
            End Using
        End Sub

        ' ── pnlCriarButtons — docked to bottom of pnlCriar ───────────────────
        ' Width set explicitly so fixed-position buttons render correctly.
        Dim pnlCriarButtons As New Panel() With {
            .Dock      = DockStyle.Bottom,
            .Height    = 50,
            .Width     = PANEL_W,
            .BackColor = TemaEscuro.Surface
        }
        Me.btnRemoverItem = New Button() With {
            .Text     = "Remove Item",
            .Location = New Point(430, 10),
            .Size     = New Size(135, 30)
        }
        Me.btnNovaRemessa = New Button() With {
            .Text     = "Create Shipment",
            .Location = New Point(575, 10),
            .Size     = New Size(140, 30),
            .Font     = New Font("Segoe UI", 10, FontStyle.Bold)
        }
        pnlCriarButtons.Controls.AddRange({btnRemoverItem, btnNovaRemessa})

        ' ── pnlContent — scrollable form fields ───────────────────────────────
        ' All controls use fixed sizes (no Anchor=Right) because panel width
        ' is fixed at PANEL_W and never changes.
        Const CTL_W  As Integer = 700   ' usable width inside 730px panel
        Const LEFT   As Integer = 15    ' left margin

        Dim pnlContent As New Panel() With {
            .Dock       = DockStyle.Fill,
            .AutoScroll = True,
            .BackColor  = TemaEscuro.Surface,
            .Padding    = New Padding(0)
        }

        Dim lblCab As New Label() With {
            .Text      = "New Shipment",
            .Location  = New Point(LEFT, 12),
            .AutoSize  = True,
            .ForeColor = TemaEscuro.Accent,
            .Font      = New Font("Segoe UI", 11, FontStyle.Bold)
        }

        ' Row 1 — Direction | Carrier | Tracking Number
        Dim lblDir As New Label() With {.Text = "Direction",         .Location = New Point(LEFT, 45),  .AutoSize = True, .ForeColor = TemaEscuro.TextoMutado}
        Me.cboDirection = New ComboBox() With {
            .Location      = New Point(LEFT, 63),
            .Size          = New Size(140, 24),
            .DropDownStyle = ComboBoxStyle.DropDownList
        }
        Dim lblCar As New Label() With {.Text = "Carrier",           .Location = New Point(170, 45), .AutoSize = True, .ForeColor = TemaEscuro.TextoMutado}
        Me.cboCarrier = New ComboBox() With {
            .Location      = New Point(170, 63),
            .Size          = New Size(140, 24),
            .DropDownStyle = ComboBoxStyle.DropDownList
        }
        Dim lblTrk As New Label() With {.Text = "Tracking Number *", .Location = New Point(325, 45), .AutoSize = True, .ForeColor = TemaEscuro.Accent}
        Me.txtTracking = New TextBox() With {
            .Location = New Point(325, 63),
            .Size     = New Size(390, 24)
        }

        ' Row 2 — Recipient
        Dim lblRec As New Label() With {.Text = "Recipient *",       .Location = New Point(LEFT, 98),  .AutoSize = True, .ForeColor = TemaEscuro.Accent}
        Me.txtRecipientName = New TextBox() With {
            .Location = New Point(LEFT, 116),
            .Size     = New Size(CTL_W, 24)
        }

        ' Row 3 — Address
        Dim lblAdr As New Label() With {.Text = "Recipient Address", .Location = New Point(LEFT, 148), .AutoSize = True, .ForeColor = TemaEscuro.TextoMutado}
        Me.txtRecipientAddress = New TextBox() With {
            .Location = New Point(LEFT, 166),
            .Size     = New Size(CTL_W, 24)
        }

        ' Row 4 — Service | Weight | Cost | Notes
        Dim lblSvc As New Label() With {.Text = "Service",      .Location = New Point(LEFT, 200),  .AutoSize = True, .ForeColor = TemaEscuro.TextoMutado}
        Me.txtServiceLevel = New TextBox() With {.Location = New Point(LEFT, 218),  .Size = New Size(180, 24)}

        Dim lblPeso As New Label() With {.Text = "Weight (lbs)", .Location = New Point(210, 200), .AutoSize = True, .ForeColor = TemaEscuro.TextoMutado}
        Me.txtPeso = New TextBox() With {.Location = New Point(210, 218), .Size = New Size(110, 24), .TextAlign = HorizontalAlignment.Right}

        Dim lblCusto As New Label() With {.Text = "Cost (USD)",   .Location = New Point(335, 200), .AutoSize = True, .ForeColor = TemaEscuro.TextoMutado}
        Me.txtCustoEnvio = New TextBox() With {.Location = New Point(335, 218), .Size = New Size(110, 24), .TextAlign = HorizontalAlignment.Right}

        Dim lblNt As New Label() With {.Text = "Shipment Notes", .Location = New Point(460, 200), .AutoSize = True, .ForeColor = TemaEscuro.TextoMutado}
        Me.txtNotes = New TextBox() With {
            .Location = New Point(460, 218),
            .Size     = New Size(255, 24)
        }

        ' Row 5 — Quick Scan
        Dim lblQS As New Label() With {.Text = "Quick Scan (UID / Serial):", .Location = New Point(LEFT, 255), .AutoSize = True, .ForeColor = TemaEscuro.Accent}
        Me.txtQuickScan = New TextBox() With {
            .Location = New Point(LEFT, 273),
            .Size     = New Size(490, 30),
            .Font     = New Font("Segoe UI", 11)
        }
        Me.lblQuickScanStatus = New Label() With {
            .Location  = New Point(515, 279),
            .Size      = New Size(195, 22),
            .ForeColor = TemaEscuro.Verde,
            .Text      = ""
        }

        ' Row 6 — Search
        Dim lblBusca As New Label() With {.Text = "Search equipment (UID / Serial / Model)", .Location = New Point(LEFT, 310), .AutoSize = True, .ForeColor = TemaEscuro.TextoMutado}
        Me.txtBuscarEquip = New TextBox() With {
            .Location = New Point(LEFT, 328),
            .Size     = New Size(530, 24)
        }
        Me.btnBuscarEquip = New Button() With {
            .Text     = "Search",
            .Location = New Point(555, 326),
            .Size     = New Size(155, 28)
        }

        Me.dgvResultadoBusca = New DataGridView() With {
            .Location      = New Point(LEFT, 362),
            .Size          = New Size(CTL_W, 110),
            .ReadOnly      = True,
            .MultiSelect   = False,
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect
        }

        ' Row 7 — Item price + note + Add button
        Dim lblPI As New Label() With {.Text = "Item Price (USD)", .Location = New Point(LEFT, 482), .AutoSize = True, .ForeColor = TemaEscuro.TextoMutado}
        Me.txtPrecoItem = New TextBox() With {
            .Location  = New Point(LEFT, 500),
            .Size      = New Size(130, 24),
            .TextAlign = HorizontalAlignment.Right
        }
        Dim lblNI As New Label() With {.Text = "Item Note", .Location = New Point(160, 482), .AutoSize = True, .ForeColor = TemaEscuro.TextoMutado}
        Me.txtNotaItem = New TextBox() With {
            .Location = New Point(160, 500),
            .Size     = New Size(385, 24)
        }
        Me.btnAdicionarItem = New Button() With {
            .Text     = "Add Item",
            .Location = New Point(555, 498),
            .Size     = New Size(155, 28)
        }

        Me.lblContadorItens = New Label() With {
            .Text      = "Items: 0",
            .Location  = New Point(LEFT, 536),
            .AutoSize  = True,
            .ForeColor = TemaEscuro.Accent
        }
        Me.dgvItensSelecionados = New DataGridView() With {
            .Location      = New Point(LEFT, 556),
            .Size          = New Size(CTL_W, 90),
            .ReadOnly      = True,
            .MultiSelect   = False,
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect
        }

        pnlContent.Controls.AddRange({
            lblCab,
            lblDir, cboDirection, lblCar, cboCarrier, lblTrk, txtTracking,
            lblRec, txtRecipientName,
            lblAdr, txtRecipientAddress,
            lblSvc, txtServiceLevel, lblPeso, txtPeso, lblCusto, txtCustoEnvio, lblNt, txtNotes,
            lblQS, txtQuickScan, lblQuickScanStatus,
            lblBusca, txtBuscarEquip, btnBuscarEquip,
            dgvResultadoBusca,
            lblPI, txtPrecoItem, lblNI, txtNotaItem, btnAdicionarItem,
            lblContadorItens, dgvItensSelecionados})

        ' Add Fill panel first, then Bottom panel — WinForms processes in reverse order
        pnlCriar.Controls.Add(pnlContent)
        pnlCriar.Controls.Add(pnlCriarButtons)

        ' ── pnlLeft — shipments grid (Dock=Fill) ─────────────────────────────
        Dim pnlLeft As New Panel() With {
            .Dock      = DockStyle.Fill,
            .BackColor = TemaEscuro.Fundo
        }
        Me.lblSelecionada = New Label() With {
            .Text      = "No shipment selected",
            .Dock      = DockStyle.Bottom,
            .Height    = 26,
            .ForeColor = TemaEscuro.Accent,
            .TextAlign = ContentAlignment.MiddleLeft,
            .Padding   = New Padding(5, 0, 0, 0)
        }
        Me.dgvRemessas = New DataGridView() With {
            .Dock = DockStyle.Fill
        }
        pnlLeft.Controls.Add(dgvRemessas)
        pnlLeft.Controls.Add(lblSelecionada)

        ' ── Assemble form (dock order: Fill → Right → Bottom → Top) ──────────
        Me.Controls.Add(pnlLeft)
        Me.Controls.Add(pnlCriar)
        Me.Controls.Add(pnlAcoes)
        Me.Controls.Add(pnlTop)

        Me.ResumeLayout(False)

    End Sub

End Class
