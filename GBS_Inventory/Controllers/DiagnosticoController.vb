Imports System.Data

Public Class DiagnosticoController

#Region "Attributes"

    Private oReader As clsReadDiagnostico

#End Region

#Region "Constructor"

    Public Sub New()
        oReader = New clsReadDiagnostico()
    End Sub

#End Region

#Region "Queries"

    Public Function fetchByFilter(pUid As String, pTag As String, pVerdict As String,
                                  pDataIni As Date?, pDataFim As Date?) As DataSet
        Return oReader.selectByFilter(pUid, pTag, pVerdict, pDataIni, pDataFim)
    End Function

    ''' <summary>
    ''' Grava o PDF do diagnostico em pPasta e retorna o caminho, ou "" se o registro nao tem PDF.
    ''' </summary>
    Public Function savePdf(pId As Long, pPasta As String) As String

        Dim pdf = oReader.selectPdf(pId)
        If pdf.Bytes Is Nothing OrElse pdf.Bytes.Length = 0 Then Return ""

        Dim nome As String = If(String.IsNullOrWhiteSpace(pdf.FileName), "diagnostic_" & pId & ".pdf", pdf.FileName)
        For Each c As Char In IO.Path.GetInvalidFileNameChars()
            nome = nome.Replace(c, "_"c)
        Next
        If Not nome.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase) Then nome &= ".pdf"

        If Not IO.Directory.Exists(pPasta) Then IO.Directory.CreateDirectory(pPasta)
        Dim caminho As String = IO.Path.Combine(pPasta, nome)
        Try
            IO.File.WriteAllBytes(caminho, pdf.Bytes)
        Catch ex As IO.IOException
            ' Arquivo aberto no leitor de PDF (travado): grava com outro nome
            caminho = IO.Path.Combine(pPasta, IO.Path.GetFileNameWithoutExtension(nome) &
                                      "_" & DateTime.Now.ToString("HHmmss") & ".pdf")
            IO.File.WriteAllBytes(caminho, pdf.Bytes)
        End Try
        Return caminho

    End Function

#End Region

End Class
