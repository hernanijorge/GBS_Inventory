Imports System.Configuration
Imports System.Data
Imports Oracle.ManagedDataAccess.Client
Imports Oracle.ManagedDataAccess.Types

''' <summary>
''' Leitura de TBL_DIAGNOSTICO (PACK_DIAGNOSTICO).
''' XE 10.2 + ODP.NET 21.x: LOB em SELECT/cursor da "TTC Error"; por isso PROC_LIST
''' nao traz LOB e o PDF e lido por parametro OUT, dentro da conexao aberta.
''' </summary>
Public Class clsReadDiagnostico

#Region "Attributes"

    Private ConnectionString As String

#End Region

#Region "Constructor"

    Public Sub New()
        Me.ConnectionString = ConfigurationManager.ConnectionStrings("OracleDB").ConnectionString
    End Sub

#End Region

#Region "Queries"

    Public Function selectByFilter(pUid As String, pTag As String, pVerdict As String,
                                   pDataIni As Date?, pDataFim As Date?) As DataSet

        Dim oPar(5) As OracleParameter
        oPar(0) = New OracleParameter("P_INTERNAL_UID", OracleDbType.Varchar2,  ParameterDirection.Input)
        oPar(1) = New OracleParameter("P_SERVICE_TAG",  OracleDbType.Varchar2,  ParameterDirection.Input)
        oPar(2) = New OracleParameter("P_VERDICT",      OracleDbType.Varchar2,  ParameterDirection.Input)
        oPar(3) = New OracleParameter("P_DATA_INI",     OracleDbType.Date,      ParameterDirection.Input)
        oPar(4) = New OracleParameter("P_DATA_FIM",     OracleDbType.Date,      ParameterDirection.Input)
        oPar(5) = New OracleParameter("P_CURSOR",       OracleDbType.RefCursor, ParameterDirection.Output)

        oPar(0).Value = If(String.IsNullOrWhiteSpace(pUid),     DBNull.Value, CObj(pUid.Trim()))
        oPar(1).Value = If(String.IsNullOrWhiteSpace(pTag),     DBNull.Value, CObj(pTag.Trim()))
        oPar(2).Value = If(String.IsNullOrWhiteSpace(pVerdict), DBNull.Value, CObj(pVerdict))
        oPar(3).Value = If(pDataIni.HasValue, CObj(pDataIni.Value.Date), DBNull.Value)
        oPar(4).Value = If(pDataFim.HasValue, CObj(pDataFim.Value.Date), DBNull.Value)

        Try
            Return OracleHelper.ExecuteDataset(Me.ConnectionString, CommandType.StoredProcedure,
                                               "PACK_DIAGNOSTICO.PROC_LIST", oPar)
        Catch ex As Exception
            Throw New Exception(ex.ToString)
        End Try

    End Function

    ''' <summary>Retorna (nome do arquivo, bytes do PDF). Bytes = Nothing quando nao ha PDF.</summary>
    Public Function selectPdf(pId As Long) As (FileName As String, Bytes As Byte())

        Using oCon As New OracleConnection(Me.ConnectionString)
            oCon.Open()
            Using oCmd As New OracleCommand("PACK_DIAGNOSTICO.PROC_GET_PDF", oCon)
                oCmd.CommandType = CommandType.StoredProcedure
                oCmd.BindByName = True

                oCmd.Parameters.Add("P_ID", OracleDbType.Decimal).Value = pId
                Dim pNome As OracleParameter = oCmd.Parameters.Add("P_PDF_FILENAME", OracleDbType.Varchar2, 200)
                pNome.Direction = ParameterDirection.Output
                Dim pBlob As OracleParameter = oCmd.Parameters.Add("P_PDF_BLOB", OracleDbType.Blob)
                pBlob.Direction = ParameterDirection.Output

                oCmd.ExecuteNonQuery()

                ' O OracleBlob precisa ser lido com a conexao ainda aberta
                Dim blob As OracleBlob = CType(pBlob.Value, OracleBlob)
                Dim nome As OracleString = CType(pNome.Value, OracleString)
                Return (If(nome.IsNull, Nothing, nome.Value), If(blob.IsNull, Nothing, blob.Value))
            End Using
        End Using

    End Function

#End Region

End Class
