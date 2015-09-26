Imports System.IO
Imports System.Configuration

Public Class FileManager

    Private Shared ctx As HttpContext = HttpContext.Current

    Public Shared Sub DownloadFile(ByVal fileName As String)

        'no friendly name, just use the file's real name
        DownloadFile(fileName, fileName)

    End Sub

    Public Shared Sub DownloadFile(ByVal fileName As String, ByVal friendlyFileName As String)
        Dim path As String = ctx.Server.MapPath(ConfigurationManager.AppSettings("ResourcesFolder"))
        Dim fs As FileStream
        fs = File.Open(path & "\" & fileName, FileMode.Open)
        Dim fileBytes(fs.Length) As Byte
        fs.Read(fileBytes, 0, fs.Length)
        fs.Close()

        ctx.Response.AddHeader("Content-disposition", String.Format("attachment; filename={0}", friendlyFileName))
        ctx.Response.ContentType = "application/octet-stream"
        ctx.Response.BinaryWrite(fileBytes)
        ctx.Response.End()
    End Sub

    Public Shared Function GetFiles(ByVal SourceDirectory As String) As IEnumerable(Of String)
        Dim returnCollection As New Queue(Of String)
        For Each f In Directory.GetFiles(SourceDirectory)
            returnCollection.Enqueue(f)
        Next
        Return returnCollection
    End Function

End Class
