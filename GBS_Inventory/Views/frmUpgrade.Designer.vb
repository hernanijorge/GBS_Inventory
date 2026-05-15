Imports System.Drawing
Imports System.Windows.Forms

Partial Class frmUpgrade
    Inherits Form

    Friend WithEvents lblTitulo          As Label
    Friend WithEvents lblUID             As Label

    ' Action type
    Friend WithEvents lblAction          As Label
    Friend WithEvents pnlAction          As Panel
    Friend WithEvents rbInstall          As RadioButton
    Friend WithEvents rbRemove           As RadioButton
    Friend WithEvents rbExtract          As RadioButton
    Friend WithEvents rbManual           As RadioButton

    ' Component type + selectors
    Friend WithEvents lblComponente      As Label
    Friend WithEvents cboComponente      As ComboBox
    Friend WithEvents lblOrigem          As Label
    Friend WithEvents cboOrigem          As ComboBox
    Friend WithEvents lblCompStock       As Label
    Friend WithEvents cboComponentStock  As ComboBox

    ' Extract to Stock panel
    Friend WithEvents pnlExtract         As Panel
    Friend WithEvents lblExtBrand        As Label
    Friend WithEvents txtExtractBrand    As TextBox
    Friend WithEvents lblExtCap          As Label
    Friend WithEvents nudExtractCapacity As NumericUpDown
    Friend WithEvents lblExtQty          As Label
    Friend WithEvents nudExtractQty      As NumericUpDown
    Friend WithEvents lblExtGen          As Label
    Friend WithEvents txtExtractGeneration As TextBox
    Friend WithEvents lblExtSpeed        As Label
    Friend WithEvents nudExtractSpeed    As NumericUpDown

    ' Disposition panel (REMOVE only)
    Friend WithEvents pnlDisposition     As Panel
    Friend WithEvents lblDisposition     As Label
    Friend WithEvents rbReturnToStock    As RadioButton
    Friend WithEvents rbScrap            As RadioButton

    ' Value before / after
    Friend WithEvents lblValorAntes      As Label
    Friend WithEvents txtValorAntes      As TextBox
    Friend WithEvents lblValorDepois     As Label
    Friend WithEvents txtValorDepois     As TextBox

    ' Part serial / technician / cost
    Friend WithEvents lblPartSerial      As Label
    Friend WithEvents txtPartSerial      As TextBox
    Friend WithEvents lblTecnico         As Label
    Friend WithEvents txtTecnico         As TextBox
    Friend WithEvents lblCusto           As Label
    Friend WithEvents txtCusto           As TextBox

    ' Notes + buttons
    Friend WithEvents lblNotas           As Label
    Friend WithEvents txtNotas           As TextBox
    Friend WithEvents btnSalvar          As Button
    Friend WithEvents btnCancelar        As Button

    Private Sub InitializeComponent()

        Me.SuspendLayout()

        Me.Text            = "Register Upgrade"
        Me.Size            = New Size(600, 630)
        Me.StartPosition   = FormStartPosition.CenterParent
        Me.BackColor       = TemaEscuro.Fundo
        Me.Font            = New Font("Segoe UI", 9)
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox     = False
        Me.MinimizeBox     = False

        ' ── Header ────────────────────────────────────────────────────────────
        Me.lblTitulo = New Label() With {
            .Text      = "New Component Upgrade",
            .Font      = New Font("Segoe UI", 13, FontStyle.Bold),
            .ForeColor = TemaEscuro.Accent,
            .Location  = New Point(20, 15),
            .AutoSize  = True
        }
        Me.lblUID = New Label() With {
            .Text      = "UID:",
            .Location  = New Point(20, 50),
            .Size      = New Size(550, 20),
            .ForeColor = TemaEscuro.TextoMutado,
            .Font      = New Font("Consolas", 10, FontStyle.Bold)
        }

        ' ── Action type (y=80) ────────────────────────────────────────────────
        Me.lblAction = New Label() With {
            .Text      = "Action Type *",
            .Location  = New Point(20, 82),
            .AutoSize  = True,
            .ForeColor = TemaEscuro.TextoMutado
        }

        Me.rbInstall = New RadioButton() With {
            .Text      = "Install from Inventory",
            .Location  = New Point(0, 5),
            .Size      = New Size(158, 20),
            .ForeColor = TemaEscuro.Texto,
            .Checked   = True
        }
        Me.rbRemove = New RadioButton() With {
            .Text      = "Remove Component",
            .Location  = New Point(161, 5),
            .Size      = New Size(140, 20),
            .ForeColor = TemaEscuro.Texto
        }
        Me.rbExtract = New RadioButton() With {
            .Text      = "Extract to Stock",
            .Location  = New Point(304, 5),
            .Size      = New Size(125, 20),
            .ForeColor = Color.FromArgb(100, 200, 120)
        }
        Me.rbManual = New RadioButton() With {
            .Text      = "Manual Entry",
            .Location  = New Point(432, 5),
            .Size      = New Size(105, 20),
            .ForeColor = TemaEscuro.Texto
        }

        Me.pnlAction = New Panel() With {
            .Location  = New Point(20, 98),
            .Size      = New Size(540, 30),
            .BackColor = Color.Transparent
        }
        Me.pnlAction.Controls.AddRange({rbInstall, rbRemove, rbExtract, rbManual})

        ' ── Component type (y=142) ────────────────────────────────────────────
        Me.lblComponente = New Label() With {
            .Text      = "Component *",
            .Location  = New Point(20, 142),
            .AutoSize  = True,
            .ForeColor = TemaEscuro.TextoMutado
        }
        Me.cboComponente = New ComboBox() With {
            .Location      = New Point(20, 162),
            .Size          = New Size(200, 24),
            .DropDownStyle = ComboBoxStyle.DropDownList
        }

        ' Part Source — MANUAL mode only
        Me.lblOrigem = New Label() With {
            .Text      = "Part Source",
            .Location  = New Point(240, 142),
            .AutoSize  = True,
            .ForeColor = TemaEscuro.TextoMutado,
            .Visible   = False
        }
        Me.cboOrigem = New ComboBox() With {
            .Location      = New Point(240, 162),
            .Size          = New Size(320, 24),
            .DropDownStyle = ComboBoxStyle.DropDown,
            .Visible       = False
        }

        ' Stock selector — INSTALL / REMOVE mode
        Me.lblCompStock = New Label() With {
            .Text      = "Select from Inventory *",
            .Location  = New Point(240, 142),
            .AutoSize  = True,
            .ForeColor = TemaEscuro.Accent
        }
        Me.cboComponentStock = New ComboBox() With {
            .Location      = New Point(240, 162),
            .Size          = New Size(320, 24),
            .DropDownStyle = ComboBoxStyle.DropDownList
        }

        ' ── Extract to Stock panel (x=240, y=142, EXTRACT mode only) ─────────
        Me.lblExtBrand = New Label() With {
            .Text      = "Brand",
            .Location  = New Point(0, 0),
            .AutoSize  = True,
            .ForeColor = TemaEscuro.TextoMutado
        }
        Me.txtExtractBrand = New TextBox() With {
            .Location  = New Point(0, 18),
            .Size      = New Size(120, 24)
        }

        Me.lblExtCap = New Label() With {
            .Text      = "Capacity (GB) *",
            .Location  = New Point(130, 0),
            .AutoSize  = True,
            .ForeColor = TemaEscuro.Accent
        }
        Me.nudExtractCapacity = New NumericUpDown() With {
            .Location  = New Point(130, 18),
            .Size      = New Size(75, 24),
            .Minimum   = 1,
            .Maximum   = 9999,
            .Value     = 8
        }

        Me.lblExtQty = New Label() With {
            .Text      = "Qty *",
            .Location  = New Point(215, 0),
            .AutoSize  = True,
            .ForeColor = TemaEscuro.Accent
        }
        Me.nudExtractQty = New NumericUpDown() With {
            .Location  = New Point(215, 18),
            .Size      = New Size(60, 24),
            .Minimum   = 1,
            .Maximum   = 999,
            .Value     = 1
        }

        Me.lblExtGen = New Label() With {
            .Text      = "Generation",
            .Location  = New Point(0, 52),
            .AutoSize  = True,
            .ForeColor = TemaEscuro.TextoMutado
        }
        Me.txtExtractGeneration = New TextBox() With {
            .Location  = New Point(0, 70),
            .Size      = New Size(120, 24)
        }

        Me.lblExtSpeed = New Label() With {
            .Text      = "Speed (MHz)",
            .Location  = New Point(130, 52),
            .AutoSize  = True,
            .ForeColor = TemaEscuro.TextoMutado
        }
        Me.nudExtractSpeed = New NumericUpDown() With {
            .Location  = New Point(130, 70),
            .Size      = New Size(80, 24),
            .Minimum   = 0,
            .Maximum   = 99999,
            .Value     = 0
        }

        Me.pnlExtract = New Panel() With {
            .Location  = New Point(240, 142),
            .Size      = New Size(320, 100),
            .BackColor = Color.Transparent,
            .Visible   = False
        }
        Me.pnlExtract.Controls.AddRange({
            lblExtBrand, txtExtractBrand,
            lblExtCap, nudExtractCapacity,
            lblExtQty, nudExtractQty,
            lblExtGen, txtExtractGeneration,
            lblExtSpeed, nudExtractSpeed})

        ' ── Disposition panel — REMOVE mode only (y=196) ─────────────────────
        Me.lblDisposition = New Label() With {
            .Text      = "Disposition:",
            .Location  = New Point(2, 5),
            .AutoSize  = True,
            .ForeColor = TemaEscuro.TextoMutado
        }
        Me.rbReturnToStock = New RadioButton() With {
            .Text      = "Return to Stock",
            .Location  = New Point(80, 3),
            .Size      = New Size(130, 20),
            .ForeColor = TemaEscuro.Texto,
            .Checked   = True
        }
        Me.rbScrap = New RadioButton() With {
            .Text      = "Scrap",
            .Location  = New Point(220, 3),
            .Size      = New Size(80, 20),
            .ForeColor = Color.FromArgb(210, 90, 60)
        }
        Me.pnlDisposition = New Panel() With {
            .Location  = New Point(240, 196),
            .Size      = New Size(320, 28),
            .BackColor = Color.Transparent,
            .Visible   = False
        }
        Me.pnlDisposition.Controls.AddRange({lblDisposition, rbReturnToStock, rbScrap})

        ' ── Value Before / After (y=250) ──────────────────────────────────────
        ' y=250 keeps both rows below pnlExtract bottom (y=142+100=242)
        Me.lblValorAntes = New Label() With {
            .Text      = "Value BEFORE (e.g.: 32 GB)",
            .Location  = New Point(20, 250),
            .AutoSize  = True,
            .ForeColor = TemaEscuro.TextoMutado
        }
        Me.txtValorAntes = New TextBox() With {
            .Location  = New Point(20, 270),
            .Size      = New Size(260, 24)
        }

        Me.lblValorDepois = New Label() With {
            .Text      = "Value AFTER (e.g.: 8 GB) *",
            .Location  = New Point(300, 250),
            .AutoSize  = True,
            .ForeColor = TemaEscuro.Accent
        }
        Me.txtValorDepois = New TextBox() With {
            .Location  = New Point(300, 270),
            .Size      = New Size(260, 24)
        }

        ' ── Part Serial / Technician (y=294) ──────────────────────────────────
        Me.lblPartSerial = New Label() With {
            .Text      = "Part Serial (if applicable)",
            .Location  = New Point(20, 294),
            .AutoSize  = True,
            .ForeColor = TemaEscuro.TextoMutado
        }
        Me.txtPartSerial = New TextBox() With {
            .Location  = New Point(20, 314),
            .Size      = New Size(260, 24)
        }

        Me.lblTecnico = New Label() With {
            .Text      = "Technician",
            .Location  = New Point(300, 294),
            .AutoSize  = True,
            .ForeColor = TemaEscuro.TextoMutado
        }
        Me.txtTecnico = New TextBox() With {
            .Location  = New Point(300, 314),
            .Size      = New Size(260, 24)
        }

        ' ── Cost (y=354) ──────────────────────────────────────────────────────
        Me.lblCusto = New Label() With {
            .Text      = "Cost (USD)",
            .Location  = New Point(20, 354),
            .AutoSize  = True,
            .ForeColor = TemaEscuro.TextoMutado
        }
        Me.txtCusto = New TextBox() With {
            .Location  = New Point(20, 374),
            .Size      = New Size(120, 24),
            .Text      = "0.00",
            .TextAlign = HorizontalAlignment.Right
        }

        ' ── Notes (y=410) ─────────────────────────────────────────────────────
        Me.lblNotas = New Label() With {
            .Text      = "Notes",
            .Location  = New Point(20, 410),
            .AutoSize  = True,
            .ForeColor = TemaEscuro.TextoMutado
        }
        Me.txtNotas = New TextBox() With {
            .Location   = New Point(20, 430),
            .Size       = New Size(540, 90),
            .Multiline  = True,
            .ScrollBars = ScrollBars.Vertical
        }

        ' ── Buttons (y=540) ───────────────────────────────────────────────────
        Me.btnCancelar = New Button() With {
            .Text     = "Cancel",
            .Location = New Point(320, 540),
            .Size     = New Size(110, 35)
        }
        Me.btnSalvar = New Button() With {
            .Text     = "Save Upgrade",
            .Location = New Point(440, 540),
            .Size     = New Size(120, 35),
            .Font     = New Font("Segoe UI", 9, FontStyle.Bold)
        }

        Me.Controls.AddRange({
            lblTitulo, lblUID,
            lblAction, pnlAction,
            lblComponente, cboComponente,
            lblOrigem, cboOrigem,
            lblCompStock, cboComponentStock,
            pnlExtract,
            pnlDisposition,
            lblValorAntes, txtValorAntes,
            lblValorDepois, txtValorDepois,
            lblPartSerial, txtPartSerial,
            lblTecnico, txtTecnico,
            lblCusto, txtCusto,
            lblNotas, txtNotas,
            btnCancelar, btnSalvar})

        Me.ResumeLayout(False)

    End Sub

End Class
