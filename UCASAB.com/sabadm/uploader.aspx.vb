Imports System.IO
Imports System.Threading.Tasks

Partial Class sabadm_uploader
    Inherits System.Web.UI.Page

    Private savePath As String = Server.MapPath("~/images/uploaded/")

    Protected Sub Button1_Click(sender As Object, e As System.EventArgs) Handles Button1.Click
        SaveFile(fil1)
    End Sub

    Protected Sub SaveFile(fil As FileUpload)
        'check for file paths in the uploader object
        If fil.HasFile Then
            Dim tempFileName As String = fil.FileName
            Dim pathToCheck As String = Path.Combine(savePath, tempFileName)
            If File.Exists(pathToCheck) Then
                Dim counter As Int16 = 1
                While File.Exists(pathToCheck)
                    counter += 1
                    tempFileName = String.Concat(counter.ToString() + fil.FileName)
                    pathToCheck = Path.Combine(savePath, tempFileName)
                End While
                Label1.Text = "A file with this name already exists. "
            End If
            Label1.Text += String.Format("Your file was saved as ""/images/uploaded/{0}"". Place this in the imageURL field.", tempFileName)
            fil.SaveAs(pathToCheck)
        Else
            Label1.Text = "Select a file to upload."
        End If
    End Sub
End Class
