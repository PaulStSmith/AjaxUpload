Partial Public Class _Default
    Inherits System.Web.UI.Page

    Private Sub AjaxUpload1_RemovedFile(ByVal sender As Object, ByVal e As AjaxUpload.RemovedFileEventArgs) Handles AjaxUpload1.RemovedFile

        Try
            Dim dir As String = Server.MapPath("~/upload")
            dir &= IIf(dir.EndsWith("\"), "", "\")
            Dim newFileName As String = e.ID.ToString() & e.FileName.Substring(e.FileName.LastIndexOf("."))
            Dim savedPath As String = dir & newFileName
            If (System.IO.File.Exists(savedPath)) Then
                System.IO.File.Delete(savedPath)
            End If
        Catch ex As Exception
            e.PreventRemove = True
            e.PreventRemoveReason = ex.Message
        End Try

    End Sub

    Private Sub AjaxUpload1_UploadedFile(ByVal sender As Object, ByVal e As AjaxUpload.UploadedFileEventArgs) Handles AjaxUpload1.UploadedFile

        Try
            Dim dir As String = Server.MapPath("~/upload")
            dir &= IIf(dir.EndsWith("\"), "", "\")
            Dim newFileName As String = e.ID.ToString() & e.File.FileName.Substring(e.File.FileName.LastIndexOf("."))
            Dim savePath As String = dir & newFileName
            e.File.SaveAs(savePath)
            e.VirtualPath = UrlBuilder("~/upload/" & newFileName)
        Catch ex As Exception
            e.UploadError = True
            e.ErrorMessage = ex.Message
        End Try

    End Sub

    ''' <summary>
    ''' Builds a full qualified URL for the specified file.
    ''' </summary>
    ''' <param name="fileName">The file name to create the URL.</param>
    ''' <returns>A full qualified URL to the specified filename.</returns>
    Private Function UrlBuilder(ByVal fileName As String) As Uri
        Return New Uri(Request.Url.Scheme & Uri.SchemeDelimiter & Request.Url.Host & IIf(Request.Url.Port <> 80, ":" & Request.Url.Port, "") & Me.ResolveUrl(fileName))
    End Function

End Class