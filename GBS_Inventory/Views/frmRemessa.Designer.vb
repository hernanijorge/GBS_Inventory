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

        Me.Text            = "Shipments - FedEx / UPS / USPS / DHL"
        Me.Size            = New Size(1400, 900)
        Me.MinimumSize     = New Size(1100, 720)
        Me.StartPosition   = FormStartPosition.CenterParent
        Me.BackColor       = TemaEscuro.Fundo
        Me.Font            = New Font("Segoe UI", 9)

        ' ── Top bar (title + total) ───────────────────────────────────────────
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

        ' ── Bottom actions panel ──────────────────────────────────────────────
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
            .Location = New Point(15, 43),
            .Size     = New Size(110, 30),
            .Anchor   = AnchorStyles.Right Or AnchorStyles.Top
        }
        ' btnFechar will be repositioned via Anchor — set initial x wide enough
        Me.btnFechar.Location = New Point(1250, 43)

        pnlAcoes.Controls.AddRange({lblAcoes, lblNovoSt, cboNovoStatus,
                                     btnAtualizarStatus, btnAbrirRastreio, btnRelatorio, btnFechar})

        ' ── Right panel — New Shipment form (fixed width, Dock=Right) ─────────
        Dim pnlCriar As New Panel() With {
            .Dock      = DockStyle.Right,
            .Width     = 725,
            .BackColor = TemaEscuro.Surface
        }
        AddHandler pnlCriar.Paint, Sub(s, e)
            Using p As New Pen(TemaEscuro.Accent, 2)
                e.Graphics.DrawRectangle(p, 0, 0, pnlCriar.Width - 1, pnlCriar.Height - 1)
            End Using
        End Sub

        Dim lblCab As New Label() With {
            .Text      = "New Shipment",
            .Location  = New Point(15, 12),
            .AutoSize  = True,
            .ForeColor = TemaEscuro.Accent,
            .Font      = New Font("Segoe UI", 11, FontStyle.Bold)
        }

        ' Row 1 — Direction | Carrier | Tracking Number
        Dim lblDir As New Label() With {.Text = "Direction",         .Location = New Point(15, 45),  .AutoSize = True, .ForeColor = TemaEscuro.TextoMutado}
        Me.cboDirection = New ComboBox() With {
            .Location      = New Point(15, 63),
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
            .Size     = New Size(380, 24),
            .Anchor   = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        }

        ' Row 2 — Recipient
        Dim lblRec As New Label() With {.Text = "Recipient *",       .Location = New Point(15, 98),  .AutoSize = True, .ForeColor = TemaEscuro.Accent}
        Me.txtRecipientName = New TextBox() With {
            .Location = New Point(15, 116),
            .Size     = New Size(695, 24),
            .Anchor   = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        }

        ' Row 3 — Address
        Dim lblAdr As New Label() With {.Text = "Recipient Address", .Location = New Point(15, 148), .AutoSize = True, .ForeColor = TemaEscuro.TextoMutado}
        Me.txtRecipientAddress = New TextBox() With {
            .Location = New Point(15, 166),
            .Size     = New Size(695, 24),
            .Anchor   = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        }

        ' Row 4 — Service | Weight | Cost | Notes
        Dim lblSvc As New Label() With {.Text = "Service",           .Location = New Point(15, 200),  .AutoSize = True, .ForeColor = TemaEscuro.TextoMutado}
        Me.txtServiceLevel = New TextBox() With {.Location = New Point(15, 218),  .Size = New Size(180, 24)}

        Dim lblPeso As New Label() With {.Text = "Weight (lbs)",     .Location = New Point(210, 200), .AutoSize = True, .ForeColor = TemaEscuro.TextoMutado}
        Me.txtPeso = New TextBox() With {.Location = New Point(210, 218), .Size = New Size(110, 24), .TextAlign = HorizontalAlignment.Right}

        Dim lblCusto As New Label() With {.Text = "Cost (USD)",      .Location = New Point(335, 200), .AutoSize = True, .ForeColor = TemaEscuro.TextoMutado}
        Me.txtCustoEnvio = New TextBox() With {.Location = New Point(335, 218), .Size = New Size(110, 24), .TextAlign = HorizontalAlignment.Right}

        Dim lblNt As New Label() With {.Text = "Shipment Notes",     .Location = New Point(460, 200), .AutoSize = True, .ForeColor = TemaEscuro.TextoMutado}
        Me.txtNotes = New TextBox() With {
            .Location = New Point(460, 218),
            .Size     = New Size(250, 24),
            .Anchor   = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        }

        ' Row 5 — Quick Scan
        Dim lblQS As New Label() With {.Text = "Quick Scan (UID / Serial):", .Location = New Point(15, 255), .AutoSize = True, .ForeColor = TemaEscuro.Accent}
        Me.txtQuickScan = New TextBox() With {
            .Location = New Point(15, 273),
            .Size     = New Size(500, 30),
            .Font     = New Font("Segoe UI", 11),
            .Anchor   = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        }
        Me.lblQuickScanStatus = New Label() With {
            .Location  = New Point(525, 277),
            .Size      = New Size(180, 22),
            .ForeColor = TemaEscuro.Verde,
            .Text      = "",
            .Anchor    = AnchorStyles.Top Or AnchorStyles.Right
        }

        ' Row 6 — Search equipment
        Dim lblBusca As New Label() With {.Text = "Search equipment (UID / Serial / Model)", .Location = New Point(15, 310), .AutoSize = True, .ForeColor = TemaEscuro.TextoMutado}
        Me.txtBuscarEquip = New TextBox() With {
            .Location = New Point(15, 328),
            .Size     = New Size(550, 24),
            .Anchor   = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        }
        Me.btnBuscarEquip = New Button() With {
            .Text     = "Search",
            .Location = New Point(578, 326),
            .Size     = New Size(130, 28),
            .Anchor   = AnchorStyles.Top Or AnchorStyles.Right
        }

        ' Search results grid
        Me.dgvResultadoBusca = New DataGridView() With {
            .Location      = New Point(15, 360),
            .Size          = New Size(695, 110),
            .ReadOnly      = True,
            .MultiSelect   = False,
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            .Anchor        = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        }

        ' Row 7 — Item price + note + add
        Dim lblPI As New Label() With {.Text = "Item Price (USD)", .Location = New Point(15, 480), .AutoSize = True, .ForeColor = TemaEscuro.TextoMutado}
        Me.txtPrecoItem = New TextBox() With {
            .Location  = New Point(15, 498),
            .Size      = New Size(130, 24),
            .TextAlign = HorizontalAlignment.Right
        }

        Dim lblNI As New Label() With {.Text = "Item Note", .Location = New Point(160, 480), .AutoSize = True, .ForeColor = TemaEscuro.TextoMutado}
        Me.txtNotaItem = New TextBox() With {
            .Location = New Point(160, 498),
            .Size     = New Size(415, 24),
            .Anchor   = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        }

        Me.btnAdicionarItem = New Button() With {
            .Text     = "Add Item",
            .Location = New Point(588, 496),
            .Size     = New Size(122, 28),
            .Anchor   = AnchorStyles.Top Or AnchorStyles.Right
        }

        ' Items counter + selected items grid
        Me.lblContadorItens = New Label() With {
            .Text      = "Items: 0",
            .Location  = New Point(15, 532),
            .AutoSize  = True,
            .ForeColor = TemaEscuro.Accent
        }
        Me.dgvItensSelecionados = New DataGridView() With {
            .Location      = New Point(15, 550),
            .Size          = New Size(695, 90),
            .ReadOnly      = True,
            .MultiSelect   = False,
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            .Anchor        = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right Or AnchorStyles.Bottom
        }

        ' Buttons — anchored to bottom-right
        Me.btnRemoverItem = New Button() With {
            .Text     = "Remove Item",
            .Location = New Point(420, 658),
            .Size     = New Size(140, 30),
            .Anchor   = AnchorStyles.Bottom Or AnchorStyles.Right
        }
        Me.btnNovaRemessa = New Button() With {
            .Text     = "Create Shipment",
            .Location = New Point(570, 658),
            .Size     = New Size(140, 30),
            .Font     = New Font("Segoe UI", 10, FontStyle.Bold),
            .Anchor   = AnchorStyles.Bottom Or AnchorStyles.Right
        }

        pnlCriar.Controls.AddRange({
            lblCab,
            lblDir, cboDirection, lblCar, cboCarrier, lblTrk, txtTracking,
            lblRec, txtRecipientName,
            lblAdr, txtRecipientAddress,
            lblSvc, txtServiceLevel, lblPeso, txtPeso, lblCusto, txtCustoEnvio, lblNt, txtNotes,
            lblQS, txtQuickScan, lblQuickScanStatus,
            lblBusca, txtBuscarEquip, btnBuscarEquip,
            dgvResultadoBusca,
            lblPI, txtPrecoItem, lblNI, txtNotaItem, btnAdicionarItem,
            lblContadorItens, dgvItensSelecionados,
            btnRemoverItem, btnNovaRemessa})

        ' ── Left panel — shipments grid (Dock=Fill) ───────────────────────────
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

        ' Add to pnlLeft in correct dock order (Bottom first, then Fill)
        pnlLeft.Controls.Add(dgvRemessas)
        pnlLeft.Controls.Add(lblSelecionada)

        ' ── Assemble form — order matters for Dock layout ─────────────────────
        ' WinForms docks in reverse Controls order: last added = first docked
        ' Order to add: Fill first, then Right, then Bottom, then Top
        Me.Controls.Add(pnlLeft)
        Me.Controls.Add(pnlCriar)
        Me.Controls.Add(pnlAcoes)
        Me.Controls.Add(pnlTop)

        Me.ResumeLayout(False)

    End Sub

End Class
