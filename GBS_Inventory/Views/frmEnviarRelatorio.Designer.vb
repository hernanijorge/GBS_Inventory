Imports System.Drawing
Imports System.Windows.Forms

Partial Class frmEnviarRelatorio
    Inherits Form

    ' ── Section 1: Generated Files
    Friend WithEvents lblSection1    As Label
    Friend WithEvents lblWordIcon    As Label
    Friend WithEvents lblWordPath    As Label
    Friend WithEvents btnOpenWord    As Button
    Friend WithEvents lblPdfIcon     As Label
    Friend WithEvents lblPdfPath     As Label
    Friend WithEvents btnOpenPdf     As Button
    Friend WithEvents lblExcelIcon   As Label
    Friend WithEvents lblExcelPath   As Label
    Friend WithEvents btnOpenExcel   As Button

    ' ── Section 2: Send by Email
    Friend WithEvents lblSection2    As Label
    Friend WithEvents lblTo          As Label
    Friend WithEvents txtTo          As TextBox
    Friend WithEvents lblCC          As Label
    Friend WithEvents txtCC          As TextBox
    Friend WithEvents lblSubject     As Label
    Friend WithEvents txtSubject     As TextBox
    Friend WithEvents lblBody        As Label
    Friend WithEvents txtBody        As TextBox

    ' ── Section 3: Attachments
    Friend WithEvents lblSection3    As Label
    Friend WithEvents chkAttachWord  As CheckBox
    Friend WithEvents chkAttachPdf   As CheckBox
    Friend WithEvents chkAttachExcel As CheckBox

    ' ── Buttons
    Friend WithEvents btnSendEmail   As Button
    Friend WithEvents btnClose       As Button

    Private Sub InitializeComponent()

        Me.SuspendLayout()

        Me.Text            = "Report Ready"
        Me.Size            = New Size(660, 560)
        Me.StartPosition   = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox     = False
        Me.MinimizeBox     = False
        Me.BackColor       = TemaEscuro.Fundo
        Me.Font            = New System.Drawing.Font("Segoe UI", 9)

        ' ── Section 1: Generated Files ───────────────────────────────────────
        Me.lblSection1           = New Label()
        Me.lblSection1.Text      = "Generated Files"
        Me.lblSection1.Font      = New System.Drawing.Font("Segoe UI", 10, System.Drawing.FontStyle.Bold)
        Me.lblSection1.ForeColor = TemaEscuro.Accent
        Me.lblSection1.Location  = New Point(20, 14)
        Me.lblSection1.AutoSize  = True

        ' Word row (y=40)
        Me.lblWordIcon           = New Label()
        Me.lblWordIcon.Text      = "Word"
        Me.lblWordIcon.ForeColor = TemaEscuro.TextoMutado
        Me.lblWordIcon.Location  = New Point(20, 42)
        Me.lblWordIcon.Size      = New Size(44, 20)
        Me.lblWordIcon.TextAlign = ContentAlignment.MiddleLeft

        Me.lblWordPath              = New Label()
        Me.lblWordPath.Text         = ""
        Me.lblWordPath.ForeColor    = TemaEscuro.Texto
        Me.lblWordPath.Location     = New Point(68, 42)
        Me.lblWordPath.Size         = New Size(472, 20)
        Me.lblWordPath.AutoEllipsis = True

        Me.btnOpenWord          = New Button()
        Me.btnOpenWord.Text     = "Open"
        Me.btnOpenWord.Location = New Point(548, 38)
        Me.btnOpenWord.Size     = New Size(72, 24)
        Me.btnOpenWord.FlatStyle = FlatStyle.Flat

        ' PDF row (y=68)
        Me.lblPdfIcon           = New Label()
        Me.lblPdfIcon.Text      = "PDF"
        Me.lblPdfIcon.ForeColor = TemaEscuro.TextoMutado
        Me.lblPdfIcon.Location  = New Point(20, 68)
        Me.lblPdfIcon.Size      = New Size(44, 20)
        Me.lblPdfIcon.TextAlign = ContentAlignment.MiddleLeft

        Me.lblPdfPath              = New Label()
        Me.lblPdfPath.Text         = ""
        Me.lblPdfPath.ForeColor    = TemaEscuro.Texto
        Me.lblPdfPath.Location     = New Point(68, 68)
        Me.lblPdfPath.Size         = New Size(472, 20)
        Me.lblPdfPath.AutoEllipsis = True

        Me.btnOpenPdf          = New Button()
        Me.btnOpenPdf.Text     = "Open"
        Me.btnOpenPdf.Location = New Point(548, 64)
        Me.btnOpenPdf.Size     = New Size(72, 24)
        Me.btnOpenPdf.FlatStyle = FlatStyle.Flat

        ' Excel row (y=94)
        Me.lblExcelIcon           = New Label()
        Me.lblExcelIcon.Text      = "Excel"
        Me.lblExcelIcon.ForeColor = TemaEscuro.TextoMutado
        Me.lblExcelIcon.Location  = New Point(20, 94)
        Me.lblExcelIcon.Size      = New Size(44, 20)
        Me.lblExcelIcon.TextAlign = ContentAlignment.MiddleLeft

        Me.lblExcelPath              = New Label()
        Me.lblExcelPath.Text         = ""
        Me.lblExcelPath.ForeColor    = TemaEscuro.Texto
        Me.lblExcelPath.Location     = New Point(68, 94)
        Me.lblExcelPath.Size         = New Size(472, 20)
        Me.lblExcelPath.AutoEllipsis = True

        Me.btnOpenExcel          = New Button()
        Me.btnOpenExcel.Text     = "Open"
        Me.btnOpenExcel.Location = New Point(548, 90)
        Me.btnOpenExcel.Size     = New Size(72, 24)
        Me.btnOpenExcel.FlatStyle = FlatStyle.Flat

        ' ── Section 2: Send by Email ─────────────────────────────────────────
        Me.lblSection2           = New Label()
        Me.lblSection2.Text      = "Send by Email"
        Me.lblSection2.Font      = New System.Drawing.Font("Segoe UI", 10, System.Drawing.FontStyle.Bold)
        Me.lblSection2.ForeColor = TemaEscuro.Accent
        Me.lblSection2.Location  = New Point(20, 126)
        Me.lblSection2.AutoSize  = True

        Me.lblTo          = New Label()
        Me.lblTo.Text     = "To (required) *"
        Me.lblTo.Location = New Point(20, 150)
        Me.lblTo.AutoSize = True
        Me.lblTo.ForeColor = TemaEscuro.Accent

        Me.txtTo          = New TextBox()
        Me.txtTo.Location = New Point(20, 168)
        Me.txtTo.Size     = New Size(614, 23)

        Me.lblCC          = New Label()
        Me.lblCC.Text     = "CC (optional)"
        Me.lblCC.Location = New Point(20, 200)
        Me.lblCC.AutoSize = True
        Me.lblCC.ForeColor = TemaEscuro.TextoMutado

        Me.txtCC          = New TextBox()
        Me.txtCC.Location = New Point(20, 218)
        Me.txtCC.Size     = New Size(614, 23)

        Me.lblSubject          = New Label()
        Me.lblSubject.Text     = "Subject"
        Me.lblSubject.Location = New Point(20, 250)
        Me.lblSubject.AutoSize = True
        Me.lblSubject.ForeColor = TemaEscuro.TextoMutado

        Me.txtSubject          = New TextBox()
        Me.txtSubject.Location = New Point(20, 268)
        Me.txtSubject.Size     = New Size(614, 23)

        Me.lblBody          = New Label()
        Me.lblBody.Text     = "Body"
        Me.lblBody.Location = New Point(20, 300)
        Me.lblBody.AutoSize = True
        Me.lblBody.ForeColor = TemaEscuro.TextoMutado

        Me.txtBody           = New TextBox()
        Me.txtBody.Location  = New Point(20, 318)
        Me.txtBody.Size      = New Size(614, 90)
        Me.txtBody.Multiline = True
        Me.txtBody.ScrollBars = ScrollBars.Vertical

        ' ── Section 3: Attachments ───────────────────────────────────────────
        Me.lblSection3           = New Label()
        Me.lblSection3.Text      = "Attachments"
        Me.lblSection3.Font      = New System.Drawing.Font("Segoe UI", 10, System.Drawing.FontStyle.Bold)
        Me.lblSection3.ForeColor = TemaEscuro.Accent
        Me.lblSection3.Location  = New Point(20, 422)
        Me.lblSection3.AutoSize  = True

        Me.chkAttachWord           = New CheckBox()
        Me.chkAttachWord.Text      = "Word (.doc)"
        Me.chkAttachWord.Location  = New Point(20, 446)
        Me.chkAttachWord.AutoSize  = True
        Me.chkAttachWord.Checked   = True
        Me.chkAttachWord.ForeColor = TemaEscuro.Texto

        Me.chkAttachPdf           = New CheckBox()
        Me.chkAttachPdf.Text      = "PDF"
        Me.chkAttachPdf.Location  = New Point(160, 446)
        Me.chkAttachPdf.AutoSize  = True
        Me.chkAttachPdf.Checked   = True
        Me.chkAttachPdf.ForeColor = TemaEscuro.Texto

        Me.chkAttachExcel           = New CheckBox()
        Me.chkAttachExcel.Text      = "Excel (.xlsx)"
        Me.chkAttachExcel.Location  = New Point(254, 446)
        Me.chkAttachExcel.AutoSize  = True
        Me.chkAttachExcel.Checked   = False
        Me.chkAttachExcel.ForeColor = TemaEscuro.Texto

        ' ── Buttons ──────────────────────────────────────────────────────────
        Me.btnSendEmail           = New Button()
        Me.btnSendEmail.Text      = "Send Email"
        Me.btnSendEmail.Location  = New Point(410, 482)
        Me.btnSendEmail.Size      = New Size(120, 34)
        Me.btnSendEmail.BackColor = TemaEscuro.Accent
        Me.btnSendEmail.ForeColor = TemaEscuro.Fundo
        Me.btnSendEmail.FlatStyle = FlatStyle.Flat
        Me.btnSendEmail.Font      = New System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Bold)

        Me.btnClose          = New Button()
        Me.btnClose.Text     = "Close"
        Me.btnClose.Location = New Point(542, 482)
        Me.btnClose.Size     = New Size(80, 34)
        Me.btnClose.FlatStyle = FlatStyle.Flat

        Me.Controls.AddRange({
            lblSection1,
            lblWordIcon, lblWordPath, btnOpenWord,
            lblPdfIcon, lblPdfPath, btnOpenPdf,
            lblExcelIcon, lblExcelPath, btnOpenExcel,
            lblSection2,
            lblTo, txtTo,
            lblCC, txtCC,
            lblSubject, txtSubject,
            lblBody, txtBody,
            lblSection3,
            chkAttachWord, chkAttachPdf, chkAttachExcel,
            btnSendEmail, btnClose})

        Me.ResumeLayout(False)

    End Sub

End Class
