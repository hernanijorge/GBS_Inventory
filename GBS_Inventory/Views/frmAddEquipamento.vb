Imports System.Windows.Forms
Imports System.Drawing
Imports System.Data
Imports System.Configuration
Imports Oracle.ManagedDataAccess.Client
Imports GBS_Inventory.OracleHelper

Public Class frmAddEquipamento

#Region "Properties"

    Public Property SavedUID As String = ""

    ''' <summary>Diagnosticos ligados a este equipamento ao salvar (PROC_RELINK).</summary>
    Public Property RelinkedTests As Integer = 0

#End Region

#Region "Pre-preenchimento a partir de um diagnostico"

    ''' <summary>
    ''' Preenche a tela com os dados de um teste de maquina ainda nao cadastrada
    ''' (linha do PACK_DIAGNOSTICO.PROC_LIST). O tecnico confere e completa o resto.
    ''' </summary>
    Public Sub PreencherDoDiagnostico(pDiag As DataRow)

        Dim fabricante As String = NormalizarFabricante(Texto(pDiag, "MANUFACTURER"))

        txtUID.Text = Texto(pDiag, "INTERNAL_UID")
        Dim tag As String = Texto(pDiag, "SERVICE_TAG")
        txtSerial.Text = If(tag.Equals("UNKNOWN", StringComparison.OrdinalIgnoreCase), "", tag)

        ' Marca: usa o item existente do combo quando houver (mesma grafia do cadastro)
        Dim existente As String = cboManufacturer.Items.Cast(Of Object)().Select(Function(o) o.ToString()).
            FirstOrDefault(Function(s) s.Equals(fabricante, StringComparison.OrdinalIgnoreCase))
        cboManufacturer.Text = If(existente, fabricante)

        txtModel.Text = NormalizarModelo(Texto(pDiag, "MODEL"), fabricante)
        txtProcessor.Text = NormalizarProcessador(Texto(pDiag, "CPU"))

        If Not IsDBNull(pDiag("RAM_GIB")) Then
            txtRam.Text = CInt(Math.Ceiling(Convert.ToDouble(pDiag("RAM_GIB")))).ToString()
        End If
        Dim sto As Integer = CapacidadeNominalGB(Texto(pDiag, "STORAGE"))
        If sto > 0 Then txtStorage.Text = sto.ToString()

        Dim bat As String = If(IsDBNull(pDiag("BATTERY_HEALTH")), "-", Convert.ToDecimal(pDiag("BATTERY_HEALTH")).ToString("0.#", Globalization.CultureInfo.InvariantCulture) & "%")
        Dim ssd As String = If(IsDBNull(pDiag("SSD_LIFE")), "-", Convert.ToDecimal(pDiag("SSD_LIFE")).ToString("0.#", Globalization.CultureInfo.InvariantCulture) & "%")
        Dim dataTeste As String = If(IsDBNull(pDiag("DATA_TESTE")), "", CDate(pDiag("DATA_TESTE")).ToString("MM/dd/yyyy", Globalization.CultureInfo.InvariantCulture))
        txtNotes.Text = $"Battery {bat} / SSD life {ssd} - tested {dataTeste}"

        Me.Text &= " (from diagnostic)"

    End Sub

    ' "Dell Inc." -> DELL, "Hewlett-Packard" -> HP; o resto em maiusculas sem sufixo societario
    Public Shared Function NormalizarFabricante(pFab As String) As String
        Dim f As String = If(pFab, "").Trim().ToUpperInvariant()
        If f.StartsWith("DELL") Then Return "DELL"
        If f = "HP" OrElse f.StartsWith("HEWLETT") OrElse f.StartsWith("HP ") Then Return "HP"
        If f.StartsWith("LENOVO") Then Return "LENOVO"
        For Each suf As String In {", INC.", " INC.", " INC", " CORPORATION", " CORP.", " LTD.", " CO., LTD."}
            If f.EndsWith(suf) Then f = f.Substring(0, f.Length - suf.Length).Trim()
        Next
        Return f
    End Function

    ' "HP EliteBook 840 G7 Notebook PC" -> "ELITEBOOK 840 G7" (padrao do cadastro: maiusculas)
    Public Shared Function NormalizarModelo(pModelo As String, pFabricante As String) As String
        Dim m As String = If(pModelo, "").Trim().ToUpperInvariant()
        If pFabricante <> "" AndAlso m.StartsWith(pFabricante & " ") Then m = m.Substring(pFabricante.Length + 1)
        For Each suf As String In {" NOTEBOOK PC", " LAPTOP PC", " LAPTOP", " NOTEBOOK"}
            If m.EndsWith(suf) Then m = m.Substring(0, m.Length - suf.Length)
        Next
        Return If(m.Length > 100, m.Substring(0, 100), m).Trim()
    End Function

    ' "Intel(R) Core(TM) i5-10310U CPU @ 1.70GHz" -> "i5-10310U"; "AMD Ryzen 5 PRO 4650U ..." -> "RYZEN 5 PRO 4650U"
    Public Shared Function NormalizarProcessador(pCpu As String) As String
        Dim c As String = If(pCpu, "").Trim()
        Dim m = System.Text.RegularExpressions.Regex.Match(c, "\b(i[3579]-\w+)\b", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
        If m.Success Then Return m.Groups(1).Value
        m = System.Text.RegularExpressions.Regex.Match(c, "Ryzen\s+\d\s+(PRO\s+)?\w+", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
        If m.Success Then Return m.Value.ToUpperInvariant()
        Return If(c.Length > 120, c.Substring(0, 120), c)
    End Function

    ''' <summary>
    ''' "SAMSUNG MZVLB256HBHQ-000H7 238GB" -> 256: primeiro "N GB/GiB/TB" convertido para a
    ''' capacidade comercial mais proxima acima (o SO reporta 238 GiB num disco de 256 GB). 0 = nao achou.
    ''' </summary>
    Public Shared Function CapacidadeNominalGB(pStorage As String) As Integer
        Dim m = System.Text.RegularExpressions.Regex.Match(If(pStorage, ""), "(\d+(?:[.,]\d+)?)\s*(TB|TIB|GB|GIB)\b",
                                                          System.Text.RegularExpressions.RegexOptions.IgnoreCase)
        If Not m.Success Then Return 0
        Dim v As Double = Double.Parse(m.Groups(1).Value.Replace(",", "."), Globalization.CultureInfo.InvariantCulture)
        If m.Groups(2).Value.ToUpperInvariant().StartsWith("T") Then v *= 1000
        For Each nominal As Integer In {32, 64, 128, 256, 512, 1000, 2000, 4000, 8000}
            If v <= nominal * 1.02 Then Return nominal
        Next
        Return CInt(Math.Round(v))
    End Function

    Private Shared Function Texto(r As DataRow, pCol As String) As String
        If r Is Nothing OrElse Not r.Table.Columns.Contains(pCol) OrElse IsDBNull(r(pCol)) Then Return ""
        Return r(pCol).ToString().Trim()
    End Function

#End Region

#Region "Inicialização"

    Public Sub New()
        InitializeComponent()
        TemaEscuro.aplicarHelius(Me)
        ConfigurarEstilos()
        CarregarManufacturers()
        ConfigurarAutoComplete(txtModel, carregarSugestoesModel())
        ConfigurarAutoComplete(txtProcessor, carregarSugestoesProcessor())
        ConfigurarAutoComplete(txtScreenSize, carregarSugestoesScreenSize())
    End Sub

    Private Sub frmAddEquipamento_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        cboCondition.SelectedIndex = 1  ' GOOD
        cboStatus.SelectedIndex    = 0  ' IN_STOCK
        txtUID.Focus()
    End Sub

    Private Sub ConfigurarEstilos()
        btnSave.BackColor = TemaEscuro.Accent
        btnSave.ForeColor = TemaEscuro.Fundo
        pnlFooter.BackColor = TemaEscuro.Surface
        lblScreenHint.ForeColor = TemaEscuro.TextoMutado
    End Sub

#End Region

#Region "Carregamento"

    Private Sub CarregarManufacturers()
        Try
            Dim cs As String = ConfigurationManager.ConnectionStrings("OracleDB").ConnectionString
            Dim ds As DataSet = OracleHelper.ExecuteDataset(cs, CommandType.Text,
                "SELECT DISTINCT MARCA FROM TBL_EQUIPAMENTO WHERE MARCA IS NOT NULL ORDER BY MARCA")
            cboManufacturer.Items.Clear()
            If ds IsNot Nothing AndAlso ds.Tables.Count > 0 Then
                For Each row As DataRow In ds.Tables(0).Rows
                    Dim val As String = row(0).ToString().Trim()
                    If Not String.IsNullOrEmpty(val) Then cboManufacturer.Items.Add(val)
                Next
            End If
        Catch
            ' silent — combo stays empty; user can still type a new brand
        End Try
    End Sub

    ' Suggestions are loaded once per dialog open (not per keystroke). UPPER(TRIM()) folds
    ' case-only duplicates ("Latitude 5420" / "LATITUDE 5420") into a single suggestion.
    Private Function carregarSugestoes(pColuna As String) As AutoCompleteStringCollection
        Dim col As New AutoCompleteStringCollection()
        Try
            Dim cs As String = ConfigurationManager.ConnectionStrings("OracleDB").ConnectionString
            Dim ds As DataSet = OracleHelper.ExecuteDataset(cs, CommandType.Text,
                "SELECT DISTINCT UPPER(TRIM(" & pColuna & ")) FROM TBL_EQUIPAMENTO" &
                " WHERE " & pColuna & " IS NOT NULL ORDER BY 1")
            If ds IsNot Nothing AndAlso ds.Tables.Count > 0 Then
                For Each row As DataRow In ds.Tables(0).Rows
                    Dim val As String = row(0).ToString()
                    If Not String.IsNullOrEmpty(val) Then col.Add(val)
                Next
            End If
        Catch
            ' silent — no suggestions; the user can still type freely
        End Try
        Return col
    End Function

    Private Function carregarSugestoesModel() As AutoCompleteStringCollection
        Return carregarSugestoes("MODEL")
    End Function

    Private Function carregarSugestoesProcessor() As AutoCompleteStringCollection
        Return carregarSugestoes("PROCESSADOR")
    End Function

    Private Function carregarSugestoesScreenSize() As AutoCompleteStringCollection
        Return carregarSugestoes("SCREEN_SIZE")
    End Function

    Private Sub ConfigurarAutoComplete(pTxt As TextBox, pSugestoes As AutoCompleteStringCollection)
        pTxt.AutoCompleteMode         = AutoCompleteMode.SuggestAppend
        pTxt.AutoCompleteSource       = AutoCompleteSource.CustomSource
        pTxt.AutoCompleteCustomSource = pSugestoes
    End Sub

#End Region

#Region "Ações"

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click

        ' Collect values
        Dim uid   As String = txtUID.Text.Trim()
        Dim marca As String = cboManufacturer.Text.Trim()
        Dim model As String = txtModel.Text.Trim()
        Dim cond  As String = If(cboCondition.SelectedIndex >= 0,
                                  cboCondition.SelectedItem.ToString(), "")
        Dim stat  As String = If(cboStatus.SelectedIndex >= 0,
                                  cboStatus.SelectedItem.ToString(), "")

        ' Validate required fields
        Dim erros As New List(Of String)
        If String.IsNullOrEmpty(uid)   Then erros.Add("  · Internal UID")
        If String.IsNullOrEmpty(marca) Then erros.Add("  · Manufacturer")
        If String.IsNullOrEmpty(model) Then erros.Add("  · Model")
        If String.IsNullOrEmpty(cond)  Then erros.Add("  · Condition Status")
        If String.IsNullOrEmpty(stat)  Then erros.Add("  · Status")

        If erros.Count > 0 Then
            MessageBox.Show("Required fields missing:" & vbCrLf & String.Join(vbCrLf, erros),
                            "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' Parse numeric fields (empty → 0)
        Dim ramGb As Integer = 0
        Dim stoGb As Integer = 0
        Integer.TryParse(txtRam.Text.Trim(), ramGb)
        Integer.TryParse(txtStorage.Text.Trim(), stoGb)

        Dim serial   As String = txtSerial.Text.Trim()
        Dim proc     As String = txtProcessor.Text.Trim()
        Dim screen   As String = txtScreenSize.Text.Trim()
        Dim batch    As String = txtBatch.Text.Trim()
        Dim notes    As String = txtNotes.Text.Trim()

        btnSave.Enabled = False
        Cursor = Cursors.WaitCursor

        Try
            Dim cs As String = ConfigurationManager.ConnectionStrings("OracleDB").ConnectionString

            Using con As New OracleConnection(cs)
                con.Open()
                Using trans As OracleTransaction = con.BeginTransaction()
                    Try
                        ' ── PROC_INSERT ───────────────────────────────
                        Using cmd As New OracleCommand("PACK_EQUIPAMENTO.PROC_INSERT", con)
                            cmd.Transaction  = trans
                            cmd.CommandType  = CommandType.StoredProcedure
                            cmd.BindByName   = True
                            cmd.Parameters.Add("P_INTERNAL_UID",       OracleDbType.Varchar2).Value = uid
                            cmd.Parameters.Add("P_SERIAL_NUMBER",      OracleDbType.Varchar2).Value = If(String.IsNullOrEmpty(serial), DBNull.Value, CObj(serial))
                            cmd.Parameters.Add("P_MODELO",             OracleDbType.Varchar2).Value = model
                            cmd.Parameters.Add("P_MARCA",              OracleDbType.Varchar2).Value = marca
                            cmd.Parameters.Add("P_PROCESSADOR",        OracleDbType.Varchar2).Value = If(String.IsNullOrEmpty(proc), DBNull.Value, CObj(proc))
                            cmd.Parameters.Add("P_RAM_GB",             OracleDbType.Decimal).Value  = ramGb
                            cmd.Parameters.Add("P_STORAGE_GB",         OracleDbType.Decimal).Value  = stoGb
                            cmd.Parameters.Add("P_CONDITION_STATUS",   OracleDbType.Varchar2).Value = cond
                            cmd.Parameters.Add("P_STATUS_EQUIPAMENTO", OracleDbType.Varchar2).Value = stat
                            cmd.Parameters.Add("P_OBSERVACAO",         OracleDbType.Varchar2).Value = If(String.IsNullOrEmpty(notes), DBNull.Value, CObj(notes))
                            cmd.Parameters.Add("P_BATTERY_CHECK",      OracleDbType.Varchar2).Value = DBNull.Value
                            cmd.Parameters.Add("P_SOURCE_BATCH",       OracleDbType.Varchar2).Value = If(String.IsNullOrEmpty(batch), DBNull.Value, CObj(batch))
                            cmd.Parameters.Add("P_SCREEN_SIZE",        OracleDbType.Varchar2).Value = If(String.IsNullOrEmpty(screen), DBNull.Value, CObj(screen))
                            cmd.ExecuteNonQuery()
                        End Using

                        trans.Commit()

                    Catch
                        Try : trans.Rollback() : Catch : End Try
                        Throw
                    End Try
                End Using
            End Using

            SavedUID     = uid

            ' Testes feitos antes do cadastro passam a apontar para este equipamento
            ' (match unico por UID / service tag). Falha aqui nao desfaz o cadastro.
            Try
                Dim r = New DiagnosticoController().relink()
                RelinkedTests = r.Linked
            Catch
                ' a aba Diagnostics refaz o relink no proximo Refresh
            End Try

            DialogResult = DialogResult.OK
            Close()

        Catch ex As Exception
            MessageBox.Show("Error saving equipment:" & vbCrLf & ex.Message,
                            "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            btnSave.Enabled = True
            Cursor = Cursors.Default
        End Try

    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        DialogResult = DialogResult.Cancel
        Close()
    End Sub

#End Region

End Class
