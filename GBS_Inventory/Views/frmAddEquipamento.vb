Imports System.Windows.Forms
Imports System.Drawing
Imports System.Data
Imports System.Configuration
Imports Oracle.ManagedDataAccess.Client
Imports GBS_Inventory.OracleHelper

Public Class frmAddEquipamento

#Region "Properties"

    Public Property SavedUID As String = ""

#End Region

#Region "Inicialização"

    Public Sub New()
        InitializeComponent()
        TemaEscuro.aplicarHelius(Me)
        ConfigurarEstilos()
        CarregarManufacturers()
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

        Dim serial As String = txtSerial.Text.Trim()
        Dim proc   As String = txtProcessor.Text.Trim()
        Dim batch  As String = txtBatch.Text.Trim()
        Dim notes  As String = txtNotes.Text.Trim()

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
                            cmd.Transaction   = trans
                            cmd.CommandType   = CommandType.StoredProcedure
                            cmd.Parameters.Add("P_INTERNAL_UID",       OracleDbType.Varchar2).Value = uid
                            cmd.Parameters.Add("P_SERIAL_NUMBER",      OracleDbType.Varchar2).Value = If(String.IsNullOrEmpty(serial), DBNull.Value, CObj(serial))
                            cmd.Parameters.Add("P_MODELO",             OracleDbType.Varchar2).Value = model
                            cmd.Parameters.Add("P_MARCA",              OracleDbType.Varchar2).Value = marca
                            cmd.Parameters.Add("P_PROCESSADOR",        OracleDbType.Varchar2).Value = If(String.IsNullOrEmpty(proc), DBNull.Value, CObj(proc))
                            cmd.Parameters.Add("P_RAM_GB",             OracleDbType.Int32).Value    = ramGb
                            cmd.Parameters.Add("P_STORAGE_GB",         OracleDbType.Int32).Value    = stoGb
                            cmd.Parameters.Add("P_STATUS_EQUIPAMENTO", OracleDbType.Varchar2).Value = stat
                            cmd.Parameters.Add("P_OBSERVACAO",         OracleDbType.Varchar2).Value = If(String.IsNullOrEmpty(notes), DBNull.Value, CObj(notes))
                            cmd.ExecuteNonQuery()
                        End Using

                        ' ── UPDATE CONDITION_STATUS + SOURCE_BATCH ─────
                        ' These columns may not exist in all environments — ORA-00904 is silenced.
                        Try
                            Using cmd2 As New OracleCommand(
                                "UPDATE TBL_EQUIPAMENTO" &
                                "   SET CONDITION_STATUS = :P_COND," &
                                "       SOURCE_BATCH     = :P_BATCH" &
                                " WHERE INTERNAL_UID = :P_UID", con)
                                cmd2.Transaction = trans
                                cmd2.Parameters.Add("P_COND",  OracleDbType.Varchar2).Value = cond
                                cmd2.Parameters.Add("P_BATCH", OracleDbType.Varchar2).Value = If(String.IsNullOrEmpty(batch), DBNull.Value, CObj(batch))
                                cmd2.Parameters.Add("P_UID",   OracleDbType.Varchar2).Value = uid
                                cmd2.ExecuteNonQuery()
                            End Using
                        Catch exUpd As Exception
                            If Not exUpd.Message.ToUpperInvariant().Contains("ORA-00904") Then Throw
                        End Try

                        trans.Commit()

                    Catch
                        Try : trans.Rollback() : Catch : End Try
                        Throw
                    End Try
                End Using
            End Using

            SavedUID     = uid
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
