Imports System.Drawing
Imports System.Windows.Forms

Partial Class frmUpgrade
    Inherits Form

    Friend WithEvents lblTitulo As Label
    Friend WithEvents lblUID As Label
    Friend WithEvents lblComponente As Label
    Friend WithEvents cboComponente As ComboBox
    Friend WithEvents lblValorAntes As Label
    Friend WithEvents txtValorAntes As TextBox
    Friend WithEvents lblValorDepois As Label
    Friend WithEvents txtValorDepois As TextBox
    Friend WithEvents lblPartSerial As Label
    Friend WithEvents txtPartSerial As TextBox
    Friend WithEvents lblOrigem As Label
    Friend WithEvents cboOrigem As ComboBox
    Friend WithEvents lblCusto As Label
    Friend WithEvents txtCusto As TextBox
    Friend WithEvents lblTecnico As Label
    Friend WithEvents txtTecnico As TextBox
    Friend WithEvents lblNotas As Label
    Friend WithEvents txtNotas As TextBox
    Friend WithEvents btnSalvar As Button
    Friend WithEvents btnCancelar As Button

    Private Sub InitializeComponent()

        Me.SuspendLayout()

        Me.Text = "Registrar Upgrade"
        Me.Size = New Size(600, 560)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.BackColor = TemaEscuro.Fundo
        Me.Font = New Font("Segoe UI", 9)
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False

        ' Título
        Me.lblTitulo = New Label()
        Me.lblTitulo.Text = "Novo Upgrade de Componente"
        Me.lblTitulo.Font = New Font("Segoe UI", 13, FontStyle.Bold)
        Me.lblTitulo.ForeColor = TemaEscuro.Accent
        Me.lblTitulo.Location = New Point(20, 15)
        Me.lblTitulo.AutoSize = True

        Me.lblUID = New Label()
        Me.lblUID.Text = "UID:"
        Me.lblUID.Location = New Point(20, 50)
        Me.lblUID.Size = New Size(550, 20)
        Me.lblUID.ForeColor = TemaEscuro.TextoMutado
        Me.lblUID.Font = New Font("Consolas", 10, FontStyle.Bold)

        ' Componente
        Me.lblComponente = New Label()
        Me.lblComponente.Text = "Componente *"
        Me.lblComponente.Location = New Point(20, 85)
        Me.lblComponente.AutoSize = True
        Me.lblComponente.ForeColor = TemaEscuro.TextoMutado

        Me.cboComponente = New ComboBox()
        Me.cboComponente.Location = New Point(20, 105)
        Me.cboComponente.Size = New Size(200, 24)
        Me.cboComponente.DropDownStyle = ComboBoxStyle.DropDownList

        ' Origem
        Me.lblOrigem = New Label()
        Me.lblOrigem.Text = "Origem da peça *"
        Me.lblOrigem.Location = New Point(240, 85)
        Me.lblOrigem.AutoSize = True
        Me.lblOrigem.ForeColor = TemaEscuro.TextoMutado

        Me.cboOrigem = New ComboBox()
        Me.cboOrigem.Location = New Point(240, 105)
        Me.cboOrigem.Size = New Size(200, 24)
        Me.cboOrigem.DropDownStyle = ComboBoxStyle.DropDownList

        ' Custo
        Me.lblCusto = New Label()
        Me.lblCusto.Text = "Custo (USD)"
        Me.lblCusto.Location = New Point(460, 85)
        Me.lblCusto.AutoSize = True
        Me.lblCusto.ForeColor = TemaEscuro.TextoMutado

        Me.txtCusto = New TextBox()
        Me.txtCusto.Location = New Point(460, 105)
        Me.txtCusto.Size = New Size(100, 24)
        Me.txtCusto.Text = "0.00"
        Me.txtCusto.TextAlign = HorizontalAlignment.Right

        ' Valor antes
        Me.lblValorAntes = New Label()
        Me.lblValorAntes.Text = "Valor ANTES (ex: 8 GB)"
        Me.lblValorAntes.Location = New Point(20, 145)
        Me.lblValorAntes.AutoSize = True
        Me.lblValorAntes.ForeColor = TemaEscuro.TextoMutado

        Me.txtValorAntes = New TextBox()
        Me.txtValorAntes.Location = New Point(20, 165)
        Me.txtValorAntes.Size = New Size(260, 24)

        ' Valor depois
        Me.lblValorDepois = New Label()
        Me.lblValorDepois.Text = "Valor DEPOIS (ex: 16 GB) *"
        Me.lblValorDepois.Location = New Point(300, 145)
        Me.lblValorDepois.AutoSize = True
        Me.lblValorDepois.ForeColor = TemaEscuro.Accent

        Me.txtValorDepois = New TextBox()
        Me.txtValorDepois.Location = New Point(300, 165)
        Me.txtValorDepois.Size = New Size(260, 24)

        ' Part Serial
        Me.lblPartSerial = New Label()
        Me.lblPartSerial.Text = "Serial da peça (se aplicável)"
        Me.lblPartSerial.Location = New Point(20, 205)
        Me.lblPartSerial.AutoSize = True
        Me.lblPartSerial.ForeColor = TemaEscuro.TextoMutado

        Me.txtPartSerial = New TextBox()
        Me.txtPartSerial.Location = New Point(20, 225)
        Me.txtPartSerial.Size = New Size(260, 24)

        ' Técnico
        Me.lblTecnico = New Label()
        Me.lblTecnico.Text = "Técnico responsável"
        Me.lblTecnico.Location = New Point(300, 205)
        Me.lblTecnico.AutoSize = True
        Me.lblTecnico.ForeColor = TemaEscuro.TextoMutado

        Me.txtTecnico = New TextBox()
        Me.txtTecnico.Location = New Point(300, 225)
        Me.txtTecnico.Size = New Size(260, 24)

        ' Notas
        Me.lblNotas = New Label()
        Me.lblNotas.Text = "Observações"
        Me.lblNotas.Location = New Point(20, 265)
        Me.lblNotas.AutoSize = True
        Me.lblNotas.ForeColor = TemaEscuro.TextoMutado

        Me.txtNotas = New TextBox()
        Me.txtNotas.Location = New Point(20, 285)
        Me.txtNotas.Size = New Size(540, 130)
        Me.txtNotas.Multiline = True
        Me.txtNotas.ScrollBars = ScrollBars.Vertical

        ' Botões
        Me.btnCancelar = New Button()
        Me.btnCancelar.Text = "Cancelar"
        Me.btnCancelar.Location = New Point(320, 440)
        Me.btnCancelar.Size = New Size(110, 35)

        Me.btnSalvar = New Button()
        Me.btnSalvar.Text = "Salvar Upgrade"
        Me.btnSalvar.Location = New Point(440, 440)
        Me.btnSalvar.Size = New Size(120, 35)
        Me.btnSalvar.Font = New Font("Segoe UI", 9, FontStyle.Bold)

        Me.Controls.AddRange({lblTitulo, lblUID,
                               lblComponente, cboComponente,
                               lblOrigem, cboOrigem,
                               lblCusto, txtCusto,
                               lblValorAntes, txtValorAntes,
                               lblValorDepois, txtValorDepois,
                               lblPartSerial, txtPartSerial,
                               lblTecnico, txtTecnico,
                               lblNotas, txtNotas,
                               btnCancelar, btnSalvar})

        Me.ResumeLayout(False)

    End Sub

End Class
