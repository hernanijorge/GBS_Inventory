Imports System.Configuration
Imports Oracle.ManagedDataAccess.Client
Imports OracleTypes = Oracle.ManagedDataAccess.Types

''' <summary>
''' Acesso ao Oracle (PACK_DIAGNOSTICO).
''' XE 10.2 + ODP.NET 21.x: LOB em SELECT/cursor da "TTC Error" no cliente;
''' LOB como parametro (IN ou OUT) funciona. Por isso toda leitura de LOB e por OUT.
''' </summary>
Public Class DiagnosticoRepository

    Public Class Registro
        Public InternalUid As String
        Public ServiceTag As String
        Public Manufacturer As String
        Public Model As String
        Public Cpu As String
        Public RamGib As Decimal?
        Public Storage As String
        Public Verdict As String
        Public BatteryHealth As Decimal?
        Public BatteryCycles As Decimal?
        Public SsdLife As Decimal?
        Public FailItems As String
        Public WarnItems As String
        Public NotTestedItems As String
        Public Technician As String
        Public Origin As String
        Public ToolVersion As String
        Public StampUtc As String
        Public ResultJson As String
        Public PdfFilename As String
        Public PdfBytes As Byte()
    End Class

    Private ReadOnly ConnectionString As String

    Public Sub New()
        ConnectionString = ConfigurationManager.ConnectionStrings("OracleDB").ConnectionString
    End Sub

    ''' <summary>Grava o diagnostico. Retorna (ID_DIAGNOSTICO, ID_EQUIPAMENTO ou Nothing).</summary>
    Public Function Inserir(pReg As Registro) As (Id As Long, IdEquipamento As Long?)

        Using con As New OracleConnection(ConnectionString)
            con.Open()
            Using cmd As New OracleCommand("PACK_DIAGNOSTICO.PROC_INSERT", con)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.BindByName = True

                AddVarchar(cmd, "P_INTERNAL_UID", pReg.InternalUid)
                AddVarchar(cmd, "P_SERVICE_TAG", pReg.ServiceTag)
                AddVarchar(cmd, "P_MANUFACTURER", pReg.Manufacturer)
                AddVarchar(cmd, "P_MODEL", pReg.Model)
                AddVarchar(cmd, "P_CPU", pReg.Cpu)
                AddNumber(cmd, "P_RAM_GIB", pReg.RamGib)
                AddVarchar(cmd, "P_STORAGE", pReg.Storage)
                AddVarchar(cmd, "P_VERDICT", pReg.Verdict)
                AddNumber(cmd, "P_BATTERY_HEALTH", pReg.BatteryHealth)
                AddNumber(cmd, "P_BATTERY_CYCLES", pReg.BatteryCycles)
                AddNumber(cmd, "P_SSD_LIFE", pReg.SsdLife)
                AddVarchar(cmd, "P_FAIL_ITEMS", pReg.FailItems)
                AddVarchar(cmd, "P_WARN_ITEMS", pReg.WarnItems)
                AddVarchar(cmd, "P_NOT_TESTED_ITEMS", pReg.NotTestedItems)
                AddVarchar(cmd, "P_TECHNICIAN", pReg.Technician)
                AddVarchar(cmd, "P_ORIGIN", pReg.Origin)
                AddVarchar(cmd, "P_TOOL_VERSION", pReg.ToolVersion)
                AddVarchar(cmd, "P_STAMP_UTC", pReg.StampUtc)

                Dim pJson As OracleParameter = cmd.Parameters.Add("P_RESULT_JSON", OracleDbType.Clob)
                pJson.Value = If(String.IsNullOrEmpty(pReg.ResultJson), CObj(DBNull.Value), pReg.ResultJson)

                AddVarchar(cmd, "P_PDF_FILENAME", pReg.PdfFilename)

                Dim pPdf As OracleParameter = cmd.Parameters.Add("P_PDF_BLOB", OracleDbType.Blob)
                pPdf.Value = If(pReg.PdfBytes Is Nothing OrElse pReg.PdfBytes.Length = 0, CObj(DBNull.Value), pReg.PdfBytes)

                Dim pId As OracleParameter = cmd.Parameters.Add("P_ID_OUT", OracleDbType.Decimal)
                pId.Direction = ParameterDirection.Output
                Dim pIdEq As OracleParameter = cmd.Parameters.Add("P_ID_EQUIPAMENTO_OUT", OracleDbType.Decimal)
                pIdEq.Direction = ParameterDirection.Output

                cmd.ExecuteNonQuery()

                Dim id As Long = CLng(pId.Value.ToString())
                Dim idEq As Long? = Nothing
                Dim vEq As OracleTypes.OracleDecimal = CType(pIdEq.Value, OracleTypes.OracleDecimal)
                If Not vEq.IsNull Then idEq = CLng(vEq.Value)

                Return (id, idEq)
            End Using
        End Using

    End Function

    ''' <summary>Le o PDF (BLOB) por parametro OUT.</summary>
    Public Function ObterPdf(pId As Long) As (Nome As String, Bytes As Byte())
        Using con As New OracleConnection(ConnectionString)
            con.Open()
            Using cmd As New OracleCommand("PACK_DIAGNOSTICO.PROC_GET_PDF", con)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.BindByName = True
                cmd.Parameters.Add("P_ID", OracleDbType.Decimal).Value = pId
                Dim pNome As OracleParameter = cmd.Parameters.Add("P_PDF_FILENAME", OracleDbType.Varchar2, 200)
                pNome.Direction = ParameterDirection.Output
                Dim pBlob As OracleParameter = cmd.Parameters.Add("P_PDF_BLOB", OracleDbType.Blob)
                pBlob.Direction = ParameterDirection.Output
                cmd.ExecuteNonQuery()

                Dim blob As OracleTypes.OracleBlob = CType(pBlob.Value, OracleTypes.OracleBlob)
                Dim nome As OracleTypes.OracleString = CType(pNome.Value, OracleTypes.OracleString)
                Return (If(nome.IsNull, Nothing, nome.Value), If(blob.IsNull, Nothing, blob.Value))
            End Using
        End Using
    End Function

    ''' <summary>Le o RESULT_JSON (CLOB) por parametro OUT.</summary>
    Public Function ObterJson(pId As Long) As String
        Using con As New OracleConnection(ConnectionString)
            con.Open()
            Using cmd As New OracleCommand("PACK_DIAGNOSTICO.PROC_GET_JSON", con)
                cmd.CommandType = CommandType.StoredProcedure
                cmd.BindByName = True
                cmd.Parameters.Add("P_ID", OracleDbType.Decimal).Value = pId
                Dim pClob As OracleParameter = cmd.Parameters.Add("P_RESULT_JSON", OracleDbType.Clob)
                pClob.Direction = ParameterDirection.Output
                cmd.ExecuteNonQuery()

                Dim clob As OracleTypes.OracleClob = CType(pClob.Value, OracleTypes.OracleClob)
                Return If(clob.IsNull, Nothing, clob.Value)
            End Using
        End Using
    End Function

    ''' <summary>Usado apenas pelo --selftest para limpar o registro de teste.</summary>
    Public Sub Excluir(pId As Long)
        Using con As New OracleConnection(ConnectionString)
            con.Open()
            Using cmd As New OracleCommand("DELETE FROM TBL_DIAGNOSTICO WHERE ID_DIAGNOSTICO = :P_ID", con)
                cmd.Parameters.Add("P_ID", OracleDbType.Decimal).Value = pId
                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Private Shared Sub AddVarchar(pCmd As OracleCommand, pNome As String, pValor As String)
        Dim p As OracleParameter = pCmd.Parameters.Add(pNome, OracleDbType.Varchar2)
        p.Value = If(String.IsNullOrWhiteSpace(pValor), CObj(DBNull.Value), pValor)
    End Sub

    Private Shared Sub AddNumber(pCmd As OracleCommand, pNome As String, pValor As Decimal?)
        Dim p As OracleParameter = pCmd.Parameters.Add(pNome, OracleDbType.Decimal)
        p.Value = If(pValor.HasValue, CObj(pValor.Value), DBNull.Value)
    End Sub

End Class
