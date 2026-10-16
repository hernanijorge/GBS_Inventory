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

    ''' <summary>
    ''' PROC_RELINK: liga diagnosticos sem equipamento (match unico por UID / service tag).
    ''' Retorna quantos foram ligados agora e quantos continuam sem equipamento.
    ''' </summary>
    Public Function relink() As (Linked As Integer, Unlinked As Integer)

        Dim oPar(1) As OracleParameter
        oPar(0) = New OracleParameter("P_LINKED",   OracleDbType.Decimal, ParameterDirection.Output)
        oPar(1) = New OracleParameter("P_UNLINKED", OracleDbType.Decimal, ParameterDirection.Output)

        Try
            OracleHelper.ExecuteNonQuery(Me.ConnectionString, CommandType.StoredProcedure,
                                         "PACK_DIAGNOSTICO.PROC_RELINK", oPar)
            Return (CInt(oPar(0).Value.ToString()), CInt(oPar(1).Value.ToString()))
        Catch ex As Exception
            Throw New Exception(ex.ToString)
        End Try

    End Function

    ''' <summary>Uma linha por equipamento da remessa com o diagnostico mais recente (sem LOB).</summary>
    Public Function selectByShipment(pIdRemessa As Long) As DataSet

        Dim oPar(1) As OracleParameter
        oPar(0) = New OracleParameter("P_ID_REMESSA", OracleDbType.Decimal,   ParameterDirection.Input)
        oPar(1) = New OracleParameter("P_CURSOR",     OracleDbType.RefCursor, ParameterDirection.Output)
        oPar(0).Value = pIdRemessa

        Try
            Return OracleHelper.ExecuteDataset(Me.ConnectionString, CommandType.StoredProcedure,
                                               "PACK_DIAGNOSTICO.PROC_LIST_BY_SHIPMENT", oPar)
        Catch ex As Exception
            Throw New Exception(ex.ToString)
        End Try

    End Function

    ''' <summary>
    ''' Cabecalho da remessa para a capa / e-mail. Cliente = TBL_REMESSA.ID_CLIENTE (script 34).
    ''' CLIENTE_NOME = nome do cliente ou, sem cliente, o DESTINATARIO; CLIENTE_EMAIL = e-mail do
    ''' cliente (nulo se nao ha cliente ou ele nao tem e-mail); ID_CLIENTE nulo = sem cliente vinculado.
    ''' </summary>
    Public Function selectShipmentHeader(pIdRemessa As Long) As DataSet

        Dim oPar(0) As OracleParameter
        oPar(0) = New OracleParameter("P_ID_REMESSA", OracleDbType.Decimal, ParameterDirection.Input)
        oPar(0).Value = pIdRemessa

        Try
            Return OracleHelper.ExecuteDataset(Me.ConnectionString, CommandType.Text,
                "SELECT R.ID_REMESSA, R.CODIGO_REMESSA, R.DESTINATARIO, R.DATA_ENVIO, R.DATA_CADASTRO," &
                "       R.ID_CLIENTE," &
                "       NVL(NVL(C.NOME_FANTASIA, C.NOME_RAZAO), R.DESTINATARIO) AS CLIENTE_NOME," &
                "       TRIM(C.EMAIL) AS CLIENTE_EMAIL" &
                "  FROM TBL_REMESSA R" &
                "  LEFT JOIN TBL_CLIENTE C ON C.ID_CLIENTE = R.ID_CLIENTE" &
                " WHERE R.ID_REMESSA = :P_ID_REMESSA",
                oPar)
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
