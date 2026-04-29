Imports System.Drawing
Imports System.Windows.Forms

''' <summary>
''' Helper de tema escuro — paleta GBS Inventory
''' Aplica estilo consistente em todos os formulários
''' </summary>
Public Class TemaEscuro

#Region "Paleta de Cores"

    ''' <summary>Fundo principal (quase preto)</summary>
    Public Shared ReadOnly Fundo         As Color = Color.FromArgb(8, 12, 16)

    ''' <summary>Fundo dos painéis/surface</summary>
    Public Shared ReadOnly Surface       As Color = Color.FromArgb(19, 25, 32)

    ''' <summary>Fundo dos controles (inputs/grid)</summary>
    Public Shared ReadOnly SurfaceClaro  As Color = Color.FromArgb(28, 36, 46)

    ''' <summary>Borda sutil</summary>
    Public Shared ReadOnly Borda         As Color = Color.FromArgb(38, 48, 60)

    ''' <summary>Accent ciano/teal (destaque principal)</summary>
    Public Shared ReadOnly Accent        As Color = Color.FromArgb(0, 229, 160)

    ''' <summary>Accent hover/pressed</summary>
    Public Shared ReadOnly AccentHover   As Color = Color.FromArgb(0, 200, 140)

    ''' <summary>Texto claro (padrão)</summary>
    Public Shared ReadOnly Texto         As Color = Color.FromArgb(232, 237, 242)

    ''' <summary>Texto mutado (labels secundários)</summary>
    Public Shared ReadOnly TextoMutado   As Color = Color.FromArgb(140, 155, 170)

    ''' <summary>Destaque amarelo (títulos)</summary>
    Public Shared ReadOnly Destaque      As Color = Color.FromArgb(255, 206, 84)

    ''' <summary>Verde power button</summary>
    Public Shared ReadOnly Verde         As Color = Color.FromArgb(46, 204, 113)

    ''' <summary>Vermelho alerta/excluir</summary>
    Public Shared ReadOnly Vermelho      As Color = Color.FromArgb(231, 76, 60)

    ''' <summary>Azul ação</summary>
    Public Shared ReadOnly Azul          As Color = Color.FromArgb(52, 152, 219)

#End Region

#Region "Paleta Helius"

    Public Shared ReadOnly HeliusFundo        As Color = Color.FromArgb(60, 62, 66)
    Public Shared ReadOnly HeliusSurface      As Color = Color.FromArgb(68, 70, 74)
    Public Shared ReadOnly HeliusSurfaceClaro As Color = Color.FromArgb(78, 80, 84)
    Public Shared ReadOnly HeliusBorda        As Color = Color.FromArgb(126, 130, 136)
    Public Shared ReadOnly HeliusTexto        As Color = Color.FromArgb(230, 232, 235)
    Public Shared ReadOnly HeliusTextoMutado  As Color = Color.FromArgb(196, 199, 203)

#End Region

#Region "Aplicar tema ao formulário"

    ''' <summary>
    ''' Aplica o tema escuro recursivamente a todos os controles de um formulário.
    ''' </summary>
    Public Shared Sub aplicar(pForm As Form)

        pForm.BackColor = Fundo
        pForm.ForeColor = Texto
        pForm.Font      = New Font("Segoe UI", 9)

        aplicarControles(pForm.Controls)

    End Sub

    ''' <summary>
    ''' Aplica variante visual inspirada no Helius (cinza industrial).
    ''' </summary>
    Public Shared Sub aplicarHelius(pForm As Form)

        pForm.BackColor = HeliusFundo
        pForm.ForeColor = HeliusTexto
        pForm.Font = New Font("Segoe UI", 9)

        aplicarControlesHelius(pForm.Controls)

    End Sub

    Private Shared Sub aplicarControles(pCtrls As Control.ControlCollection)

        For Each ctrl As Control In pCtrls

            Select Case True

                Case TypeOf ctrl Is Panel
                    ctrl.BackColor = Surface
                    ctrl.ForeColor = Texto

                Case TypeOf ctrl Is Label
                    ctrl.BackColor = Color.Transparent
                    ctrl.ForeColor = Texto

                Case TypeOf ctrl Is Button
                    estilizarBotao(CType(ctrl, Button))

                Case TypeOf ctrl Is TextBox
                    estilizarTextBox(CType(ctrl, TextBox))

                Case TypeOf ctrl Is ComboBox
                    estilizarComboBox(CType(ctrl, ComboBox))

                Case TypeOf ctrl Is DataGridView
                    estilizarGrid(CType(ctrl, DataGridView))

                Case TypeOf ctrl Is TabControl
                    estilizarTabs(CType(ctrl, TabControl))

                Case TypeOf ctrl Is GroupBox
                    ctrl.BackColor = Surface
                    ctrl.ForeColor = Accent

                Case TypeOf ctrl Is CheckBox, TypeOf ctrl Is RadioButton
                    ctrl.BackColor = Color.Transparent
                    ctrl.ForeColor = Texto

            End Select

            If ctrl.HasChildren Then
                aplicarControles(ctrl.Controls)
            End If

        Next

    End Sub

    Private Shared Sub aplicarControlesHelius(pCtrls As Control.ControlCollection)

        For Each ctrl As Control In pCtrls

            Select Case True

                Case TypeOf ctrl Is Panel
                    ctrl.BackColor = HeliusSurface
                    ctrl.ForeColor = HeliusTexto

                Case TypeOf ctrl Is Label
                    ctrl.BackColor = Color.Transparent
                    ctrl.ForeColor = HeliusTexto

                Case TypeOf ctrl Is Button
                    estilizarBotaoHelius(CType(ctrl, Button))

                Case TypeOf ctrl Is TextBox
                    estilizarTextBoxHelius(CType(ctrl, TextBox))

                Case TypeOf ctrl Is ComboBox
                    estilizarComboBoxHelius(CType(ctrl, ComboBox))

                Case TypeOf ctrl Is DataGridView
                    estilizarGridHelius(CType(ctrl, DataGridView))

                Case TypeOf ctrl Is TabControl
                    estilizarTabs(CType(ctrl, TabControl))

                Case TypeOf ctrl Is GroupBox
                    ctrl.BackColor = HeliusSurface
                    ctrl.ForeColor = HeliusTexto

                Case TypeOf ctrl Is CheckBox, TypeOf ctrl Is RadioButton
                    ctrl.BackColor = Color.Transparent
                    ctrl.ForeColor = HeliusTexto

            End Select

            If ctrl.HasChildren Then
                aplicarControlesHelius(ctrl.Controls)
            End If

        Next

    End Sub

#End Region

#Region "Estilizadores individuais"

    Public Shared Sub estilizarBotao(pBtn As Button)

        pBtn.FlatStyle                  = FlatStyle.Flat
        pBtn.BackColor                  = Surface
        pBtn.ForeColor                  = Accent
        pBtn.FlatAppearance.BorderColor = Accent
        pBtn.FlatAppearance.BorderSize  = 1
        pBtn.FlatAppearance.MouseOverBackColor = AccentHover
        pBtn.FlatAppearance.MouseDownBackColor = Accent
        pBtn.Font                       = New Font("Segoe UI", 9, FontStyle.Regular)
        pBtn.Cursor                     = Cursors.Hand
        pBtn.Height                     = Math.Max(pBtn.Height, 28)

    End Sub

    Public Shared Sub estilizarBotaoPrimario(pBtn As Button)

        estilizarBotao(pBtn)
        pBtn.BackColor = Accent
        pBtn.ForeColor = Fundo
        pBtn.Font      = New Font("Segoe UI", 9, FontStyle.Bold)

    End Sub

    Public Shared Sub estilizarBotaoVermelho(pBtn As Button)

        estilizarBotao(pBtn)
        pBtn.ForeColor                    = Vermelho
        pBtn.FlatAppearance.BorderColor   = Vermelho
        pBtn.FlatAppearance.MouseOverBackColor = Color.FromArgb(60, 20, 20)

    End Sub

    Public Shared Sub estilizarTextBox(pTxt As TextBox)

        pTxt.BackColor   = SurfaceClaro
        pTxt.ForeColor   = Texto
        pTxt.BorderStyle = BorderStyle.FixedSingle
        pTxt.Font        = New Font("Segoe UI", 9)

    End Sub

    Public Shared Sub estilizarBotaoHelius(pBtn As Button)

        pBtn.FlatStyle = FlatStyle.Flat
        pBtn.BackColor = HeliusSurface
        pBtn.ForeColor = HeliusTexto
        pBtn.FlatAppearance.BorderColor = HeliusBorda
        pBtn.FlatAppearance.BorderSize = 1
        pBtn.FlatAppearance.MouseOverBackColor = HeliusSurfaceClaro
        pBtn.FlatAppearance.MouseDownBackColor = Color.FromArgb(88, 90, 95)
        pBtn.Font = New Font("Segoe UI", 9, FontStyle.Regular)
        pBtn.Cursor = Cursors.Hand
        pBtn.Height = Math.Max(pBtn.Height, 28)

    End Sub

    Public Shared Sub estilizarTextBoxHelius(pTxt As TextBox)

        pTxt.BackColor = HeliusSurfaceClaro
        pTxt.ForeColor = HeliusTexto
        pTxt.BorderStyle = BorderStyle.FixedSingle
        pTxt.Font = New Font("Segoe UI", 9)

    End Sub

    Public Shared Sub estilizarComboBoxHelius(pCbo As ComboBox)

        pCbo.BackColor = HeliusSurfaceClaro
        pCbo.ForeColor = HeliusTexto
        pCbo.FlatStyle = FlatStyle.Flat
        pCbo.Font = New Font("Segoe UI", 9)

    End Sub

    Public Shared Sub estilizarGridHelius(pGrid As DataGridView)

        With pGrid
            .BackgroundColor = HeliusSurface
            .BorderStyle = BorderStyle.None
            .GridColor = HeliusBorda
            .EnableHeadersVisualStyles = False
            .ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None
            .CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal

            .ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(74, 76, 80)
            .ColumnHeadersDefaultCellStyle.ForeColor = HeliusTexto
            .ColumnHeadersDefaultCellStyle.Font = New Font("Segoe UI", 9, FontStyle.Bold)
            .ColumnHeadersDefaultCellStyle.SelectionBackColor = Color.FromArgb(74, 76, 80)
            .ColumnHeadersHeight = 32

            .DefaultCellStyle.BackColor = HeliusSurface
            .DefaultCellStyle.ForeColor = HeliusTexto
            .DefaultCellStyle.SelectionBackColor = Color.FromArgb(96, 98, 104)
            .DefaultCellStyle.SelectionForeColor = HeliusTexto
            .DefaultCellStyle.Font = New Font("Segoe UI", 9)

            .AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(63, 65, 70)

            .RowTemplate.Height = 26
            .RowHeadersVisible = False
            .AllowUserToAddRows = False
            .AllowUserToDeleteRows = False
            .AllowUserToResizeRows = False
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect
            .MultiSelect = False
            .ReadOnly = True
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        End With

    End Sub

    Public Shared Sub estilizarComboBox(pCbo As ComboBox)

        pCbo.BackColor     = SurfaceClaro
        pCbo.ForeColor     = Texto
        pCbo.FlatStyle     = FlatStyle.Flat
        pCbo.Font          = New Font("Segoe UI", 9)

    End Sub

    Public Shared Sub estilizarGrid(pGrid As DataGridView)

        With pGrid
            .BackgroundColor                             = Surface
            .BorderStyle                                 = BorderStyle.None
            .GridColor                                   = Borda
            .EnableHeadersVisualStyles                   = False
            .ColumnHeadersBorderStyle                    = DataGridViewHeaderBorderStyle.None
            .CellBorderStyle                             = DataGridViewCellBorderStyle.SingleHorizontal

            .ColumnHeadersDefaultCellStyle.BackColor     = Color.FromArgb(24, 32, 42)
            .ColumnHeadersDefaultCellStyle.ForeColor     = Accent
            .ColumnHeadersDefaultCellStyle.Font          = New Font("Segoe UI", 9, FontStyle.Bold)
            .ColumnHeadersDefaultCellStyle.SelectionBackColor = Color.FromArgb(24, 32, 42)
            .ColumnHeadersHeight                         = 32

            .DefaultCellStyle.BackColor                  = Surface
            .DefaultCellStyle.ForeColor                  = Texto
            .DefaultCellStyle.SelectionBackColor         = Accent
            .DefaultCellStyle.SelectionForeColor         = Fundo
            .DefaultCellStyle.Font                       = New Font("Segoe UI", 9)

            .AlternatingRowsDefaultCellStyle.BackColor   = Color.FromArgb(22, 29, 38)

            .RowTemplate.Height                          = 26
            .RowHeadersVisible                           = False
            .AllowUserToAddRows                          = False
            .AllowUserToDeleteRows                       = False
            .AllowUserToResizeRows                       = False
            .SelectionMode                               = DataGridViewSelectionMode.FullRowSelect
            .MultiSelect                                 = False
            .ReadOnly                                    = True
            .AutoSizeColumnsMode                         = DataGridViewAutoSizeColumnsMode.Fill
        End With

    End Sub

    Public Shared Sub estilizarTabs(pTab As TabControl)

        pTab.Appearance = TabAppearance.Normal
        pTab.DrawMode   = TabDrawMode.OwnerDrawFixed

        AddHandler pTab.DrawItem, AddressOf drawTabItem

        For Each page As TabPage In pTab.TabPages
            page.BackColor = Fundo
            page.ForeColor = Texto
            page.Padding   = New Padding(8)
        Next

    End Sub

    Private Shared Sub drawTabItem(sender As Object, e As DrawItemEventArgs)

        Dim tab As TabControl = CType(sender, TabControl)
        Dim g   As Graphics   = e.Graphics
        Dim r   As Rectangle  = tab.GetTabRect(e.Index)
        Dim ativa As Boolean  = (tab.SelectedIndex = e.Index)

        Dim corFundoTab As Color = If(ativa, Surface,  Fundo)
        Dim corTexto    As Color = If(ativa, Accent,   TextoMutado)

        Using b As New SolidBrush(corFundoTab)
            g.FillRectangle(b, r)
        End Using

        If ativa Then
            Using p As New Pen(Accent, 2)
                g.DrawLine(p, r.Left, r.Bottom - 1, r.Right, r.Bottom - 1)
            End Using
        End If

        Using br As New SolidBrush(corTexto)
            Dim sf As New StringFormat() With {
                .Alignment = StringAlignment.Center,
                .LineAlignment = StringAlignment.Center
            }
            g.DrawString(tab.TabPages(e.Index).Text, New Font("Segoe UI", 9), br, r, sf)
        End Using

    End Sub

#End Region

End Class
