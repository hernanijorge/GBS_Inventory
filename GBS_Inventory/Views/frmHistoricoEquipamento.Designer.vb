Partial Class frmHistoricoEquipamento
    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer

    Protected Overrides Sub Dispose(disposing As Boolean)
        If disposing AndAlso (components IsNot Nothing) Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    Private Sub InitializeComponent()

        ' ── Controls ──────────────────────────────────────────────────────────
        Me.pnlHeader        = New System.Windows.Forms.Panel()
        Me.lblUID           = New System.Windows.Forms.Label()
        Me.lblHeaderInfo    = New System.Windows.Forms.Label()
        Me.tabControl       = New System.Windows.Forms.TabControl()
        Me.tabArrival       = New System.Windows.Forms.TabPage()
        Me.tabUpgrades      = New System.Windows.Forms.TabPage()
        Me.tabShipment      = New System.Windows.Forms.TabPage()
        Me.pnlFooter        = New System.Windows.Forms.Panel()
        Me.btnHistoryReport = New System.Windows.Forms.Button()
        Me.btnClose         = New System.Windows.Forms.Button()

        ' ── Arrival tab controls ───────────────────────────────────────────
        Me.pnlArrival       = New System.Windows.Forms.Panel()
        Me.lblArrDate       = New System.Windows.Forms.Label()
        Me.lblArrDateVal    = New System.Windows.Forms.Label()
        Me.lblArrBatch      = New System.Windows.Forms.Label()
        Me.lblArrBatchVal   = New System.Windows.Forms.Label()
        Me.lblArrCond       = New System.Windows.Forms.Label()
        Me.lblArrCondVal    = New System.Windows.Forms.Label()
        Me.lblArrStatus     = New System.Windows.Forms.Label()
        Me.lblArrStatusVal  = New System.Windows.Forms.Label()
        Me.lblArrSerial     = New System.Windows.Forms.Label()
        Me.lblArrSerialVal  = New System.Windows.Forms.Label()
        Me.lblArrObs        = New System.Windows.Forms.Label()
        Me.txtArrObs        = New System.Windows.Forms.TextBox()

        ' ── Upgrades tab controls ──────────────────────────────────────────
        Me.pnlUpgrades      = New System.Windows.Forms.Panel()
        Me.dgvUpgrades      = New System.Windows.Forms.DataGridView()
        Me.pnlUpgradeBar    = New System.Windows.Forms.Panel()
        Me.lblUpgradeTotal  = New System.Windows.Forms.Label()
        Me.btnUpgradeReport = New System.Windows.Forms.Button()

        ' ── Shipment tab controls ──────────────────────────────────────────
        Me.pnlShipment      = New System.Windows.Forms.Panel()
        Me.dgvShipment      = New System.Windows.Forms.DataGridView()
        Me.lblShipmentTotal = New System.Windows.Forms.Label()

        ' ── Begin Init ────────────────────────────────────────────────────
        Me.pnlHeader.SuspendLayout()
        Me.tabControl.SuspendLayout()
        Me.tabArrival.SuspendLayout()
        Me.tabUpgrades.SuspendLayout()
        Me.tabShipment.SuspendLayout()
        Me.pnlArrival.SuspendLayout()
        Me.pnlUpgrades.SuspendLayout()
        Me.pnlUpgradeBar.SuspendLayout()
        Me.pnlShipment.SuspendLayout()
        CType(Me.dgvUpgrades, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvShipment, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlFooter.SuspendLayout()
        Me.SuspendLayout()

        ' ── frmHistoricoEquipamento ────────────────────────────────────────
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0F, 15.0F)
        Me.AutoScaleMode       = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize          = New System.Drawing.Size(1080, 680)
        Me.MinimumSize         = New System.Drawing.Size(900, 580)
        Me.StartPosition       = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text                = "Equipment History"
        Me.Font                = New System.Drawing.Font("Segoe UI", 9.0F)

        ' ── pnlHeader ─────────────────────────────────────────────────────
        Me.pnlHeader.Dock      = System.Windows.Forms.DockStyle.Top
        Me.pnlHeader.Height    = 64
        Me.pnlHeader.Padding   = New System.Windows.Forms.Padding(12, 8, 12, 8)

        Me.lblUID.AutoSize  = True
        Me.lblUID.Font      = New System.Drawing.Font("Segoe UI", 14.0F, System.Drawing.FontStyle.Bold)
        Me.lblUID.Location  = New System.Drawing.Point(12, 8)
        Me.lblUID.Text      = "UID"

        Me.lblHeaderInfo.AutoSize  = False
        Me.lblHeaderInfo.Dock      = System.Windows.Forms.DockStyle.Bottom
        Me.lblHeaderInfo.Height    = 20
        Me.lblHeaderInfo.Font      = New System.Drawing.Font("Segoe UI", 8.5F)
        Me.lblHeaderInfo.Text      = ""
        Me.lblHeaderInfo.Padding   = New System.Windows.Forms.Padding(12, 0, 0, 4)

        Me.pnlHeader.Controls.Add(Me.lblUID)
        Me.pnlHeader.Controls.Add(Me.lblHeaderInfo)

        ' ── pnlFooter ─────────────────────────────────────────────────────
        Me.pnlFooter.Dock   = System.Windows.Forms.DockStyle.Bottom
        Me.pnlFooter.Height = 48
        Me.pnlFooter.Padding = New System.Windows.Forms.Padding(8, 8, 8, 8)

        Me.btnHistoryReport.Text      = "Generate History Report"
        Me.btnHistoryReport.Size      = New System.Drawing.Size(180, 30)
        Me.btnHistoryReport.Location  = New System.Drawing.Point(8, 9)
        Me.btnHistoryReport.FlatStyle = System.Windows.Forms.FlatStyle.Flat

        Me.btnClose.Text      = "Close"
        Me.btnClose.Size      = New System.Drawing.Size(90, 30)
        Me.btnClose.Location  = New System.Drawing.Point(196, 9)
        Me.btnClose.FlatStyle = System.Windows.Forms.FlatStyle.Flat

        Me.pnlFooter.Controls.Add(Me.btnHistoryReport)
        Me.pnlFooter.Controls.Add(Me.btnClose)

        ' ── tabControl ────────────────────────────────────────────────────
        Me.tabControl.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tabControl.TabPages.Add(Me.tabArrival)
        Me.tabControl.TabPages.Add(Me.tabUpgrades)
        Me.tabControl.TabPages.Add(Me.tabShipment)

        Me.tabArrival.Text    = "Arrival"
        Me.tabUpgrades.Text   = "Upgrades"
        Me.tabShipment.Text   = "Shipment"

        ' ── ABA Arrival ───────────────────────────────────────────────────
        Me.pnlArrival.Dock    = System.Windows.Forms.DockStyle.Fill
        Me.pnlArrival.Padding = New System.Windows.Forms.Padding(16, 16, 16, 16)

        Dim rowY As Integer = 16

        BuildLabel(Me.lblArrDate,      "Date Registered:",  16, rowY)
        BuildValue(Me.lblArrDateVal,   200, rowY, 160)
        rowY += 30

        BuildLabel(Me.lblArrBatch,     "Source Batch:",     16, rowY)
        BuildValue(Me.lblArrBatchVal,  200, rowY, 200)
        rowY += 30

        BuildLabel(Me.lblArrCond,      "Condition Status:", 16, rowY)
        BuildValue(Me.lblArrCondVal,   200, rowY, 160)
        rowY += 30

        BuildLabel(Me.lblArrStatus,    "Status:",           16, rowY)
        BuildValue(Me.lblArrStatusVal, 200, rowY, 160)
        rowY += 30

        BuildLabel(Me.lblArrSerial,    "Serial Number:",    16, rowY)
        BuildValue(Me.lblArrSerialVal, 200, rowY, 200)
        rowY += 36

        BuildLabel(Me.lblArrObs,       "Notes / Arrival observation:", 16, rowY)
        rowY += 22

        Me.txtArrObs.Location   = New System.Drawing.Point(16, rowY)
        Me.txtArrObs.Size       = New System.Drawing.Size(700, 80)
        Me.txtArrObs.Multiline  = True
        Me.txtArrObs.ReadOnly   = True
        Me.txtArrObs.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtArrObs.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle

        Me.pnlArrival.Controls.AddRange(New System.Windows.Forms.Control() {
            Me.lblArrDate, Me.lblArrDateVal,
            Me.lblArrBatch, Me.lblArrBatchVal,
            Me.lblArrCond, Me.lblArrCondVal,
            Me.lblArrStatus, Me.lblArrStatusVal,
            Me.lblArrSerial, Me.lblArrSerialVal,
            Me.lblArrObs, Me.txtArrObs
        })
        Me.tabArrival.Controls.Add(Me.pnlArrival)

        ' ── ABA Upgrades ──────────────────────────────────────────────────
        Me.pnlUpgrades.Dock    = System.Windows.Forms.DockStyle.Fill
        Me.pnlUpgrades.Padding = New System.Windows.Forms.Padding(8)

        Me.dgvUpgrades.Dock            = System.Windows.Forms.DockStyle.Fill
        Me.dgvUpgrades.ReadOnly        = True
        Me.dgvUpgrades.MultiSelect     = False
        Me.dgvUpgrades.SelectionMode   = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvUpgrades.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.None
        Me.dgvUpgrades.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvUpgrades.AllowUserToAddRows    = False
        Me.dgvUpgrades.AllowUserToDeleteRows = False
        Me.dgvUpgrades.RowHeadersVisible     = False

        Me.pnlUpgradeBar.Dock    = System.Windows.Forms.DockStyle.Bottom
        Me.pnlUpgradeBar.Height  = 38
        Me.pnlUpgradeBar.Padding = New System.Windows.Forms.Padding(4, 4, 4, 4)

        Me.lblUpgradeTotal.Dock      = System.Windows.Forms.DockStyle.Fill
        Me.lblUpgradeTotal.Font      = New System.Drawing.Font("Segoe UI", 9.0F, System.Drawing.FontStyle.Bold)
        Me.lblUpgradeTotal.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.lblUpgradeTotal.Text      = "Total upgrade cost: $ 0.00"
        Me.lblUpgradeTotal.Padding   = New System.Windows.Forms.Padding(4, 0, 0, 0)

        Me.btnUpgradeReport.Text      = "Upgrade Report (Customer)"
        Me.btnUpgradeReport.Size      = New System.Drawing.Size(190, 28)
        Me.btnUpgradeReport.Dock      = System.Windows.Forms.DockStyle.Right
        Me.btnUpgradeReport.FlatStyle = System.Windows.Forms.FlatStyle.Flat

        Me.pnlUpgradeBar.Controls.Add(Me.lblUpgradeTotal)
        Me.pnlUpgradeBar.Controls.Add(Me.btnUpgradeReport)

        Me.pnlUpgrades.Controls.Add(Me.dgvUpgrades)
        Me.pnlUpgrades.Controls.Add(Me.pnlUpgradeBar)
        Me.tabUpgrades.Controls.Add(Me.pnlUpgrades)

        ' ── ABA Shipment ──────────────────────────────────────────────────
        Me.pnlShipment.Dock    = System.Windows.Forms.DockStyle.Fill
        Me.pnlShipment.Padding = New System.Windows.Forms.Padding(8)

        Me.dgvShipment.Dock            = System.Windows.Forms.DockStyle.Fill
        Me.dgvShipment.ReadOnly        = True
        Me.dgvShipment.MultiSelect     = False
        Me.dgvShipment.SelectionMode   = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvShipment.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.None
        Me.dgvShipment.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvShipment.AllowUserToAddRows    = False
        Me.dgvShipment.AllowUserToDeleteRows = False
        Me.dgvShipment.RowHeadersVisible     = False

        Me.lblShipmentTotal.Dock      = System.Windows.Forms.DockStyle.Bottom
        Me.lblShipmentTotal.Height    = 24
        Me.lblShipmentTotal.Font      = New System.Drawing.Font("Segoe UI", 9.0F, System.Drawing.FontStyle.Bold)
        Me.lblShipmentTotal.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.lblShipmentTotal.Text      = "Sale price: $ 0.00  |  Shipping cost: $ 0.00"
        Me.lblShipmentTotal.Padding   = New System.Windows.Forms.Padding(4, 0, 0, 0)

        Me.pnlShipment.Controls.Add(Me.dgvShipment)
        Me.pnlShipment.Controls.Add(Me.lblShipmentTotal)
        Me.tabShipment.Controls.Add(Me.pnlShipment)

        ' ── Form layout (Dock order: Bottom first, then Top, then Fill) ──
        Me.Controls.Add(Me.tabControl)
        Me.Controls.Add(Me.pnlFooter)
        Me.Controls.Add(Me.pnlHeader)

        ' ── Resume ────────────────────────────────────────────────────────
        Me.pnlHeader.ResumeLayout(False)
        Me.tabControl.ResumeLayout(False)
        Me.tabArrival.ResumeLayout(False)
        Me.tabUpgrades.ResumeLayout(False)
        Me.tabShipment.ResumeLayout(False)
        Me.pnlArrival.ResumeLayout(False)
        Me.pnlUpgrades.ResumeLayout(False)
        Me.pnlUpgradeBar.ResumeLayout(False)
        Me.pnlShipment.ResumeLayout(False)
        CType(Me.dgvUpgrades, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvShipment, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlFooter.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    ' ── Helper para labels fixos ──────────────────────────────────────────────
    Private Sub BuildLabel(lbl As System.Windows.Forms.Label, texto As String, x As Integer, y As Integer)
        lbl.AutoSize  = True
        lbl.Font      = New System.Drawing.Font("Segoe UI", 9.0F, System.Drawing.FontStyle.Bold)
        lbl.Location  = New System.Drawing.Point(x, y + 2)
        lbl.Text      = texto
    End Sub

    Private Sub BuildValue(lbl As System.Windows.Forms.Label, x As Integer, y As Integer, w As Integer)
        lbl.AutoSize  = False
        lbl.Width     = w
        lbl.Height    = 22
        lbl.Font      = New System.Drawing.Font("Segoe UI", 9.0F)
        lbl.Location  = New System.Drawing.Point(x, y + 2)
        lbl.Text      = "—"
        lbl.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
    End Sub

    ' ── Declarações ───────────────────────────────────────────────────────────
    Friend WithEvents pnlHeader        As System.Windows.Forms.Panel
    Friend WithEvents lblUID           As System.Windows.Forms.Label
    Friend WithEvents lblHeaderInfo    As System.Windows.Forms.Label
    Friend WithEvents tabControl       As System.Windows.Forms.TabControl
    Friend WithEvents tabArrival       As System.Windows.Forms.TabPage
    Friend WithEvents tabUpgrades      As System.Windows.Forms.TabPage
    Friend WithEvents tabShipment      As System.Windows.Forms.TabPage
    Friend WithEvents pnlFooter        As System.Windows.Forms.Panel
    Friend WithEvents btnHistoryReport As System.Windows.Forms.Button
    Friend WithEvents btnClose         As System.Windows.Forms.Button

    Friend WithEvents pnlArrival       As System.Windows.Forms.Panel
    Friend WithEvents lblArrDate       As System.Windows.Forms.Label
    Friend WithEvents lblArrDateVal    As System.Windows.Forms.Label
    Friend WithEvents lblArrBatch      As System.Windows.Forms.Label
    Friend WithEvents lblArrBatchVal   As System.Windows.Forms.Label
    Friend WithEvents lblArrCond       As System.Windows.Forms.Label
    Friend WithEvents lblArrCondVal    As System.Windows.Forms.Label
    Friend WithEvents lblArrStatus     As System.Windows.Forms.Label
    Friend WithEvents lblArrStatusVal  As System.Windows.Forms.Label
    Friend WithEvents lblArrSerial     As System.Windows.Forms.Label
    Friend WithEvents lblArrSerialVal  As System.Windows.Forms.Label
    Friend WithEvents lblArrObs        As System.Windows.Forms.Label
    Friend WithEvents txtArrObs        As System.Windows.Forms.TextBox

    Friend WithEvents pnlUpgrades      As System.Windows.Forms.Panel
    Friend WithEvents dgvUpgrades      As System.Windows.Forms.DataGridView
    Friend WithEvents pnlUpgradeBar    As System.Windows.Forms.Panel
    Friend WithEvents lblUpgradeTotal  As System.Windows.Forms.Label
    Friend WithEvents btnUpgradeReport As System.Windows.Forms.Button

    Friend WithEvents pnlShipment      As System.Windows.Forms.Panel
    Friend WithEvents dgvShipment      As System.Windows.Forms.DataGridView
    Friend WithEvents lblShipmentTotal As System.Windows.Forms.Label

End Class
