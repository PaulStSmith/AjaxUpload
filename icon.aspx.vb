'**************************************************
' FILE      : icon.aspx.vb
' AUTHOR    : Paulo.Santos
' CREATION  : 10/26/2007 10:52:47 AM
' COPYRIGHT : Copyright © 2007
'             PJ on Development
'             All Rights Reserved.
'
' Description:
'       Generates the icon image for the specified file.
'
' Change log:
' 0.1   10/26/2007 10:52:47 AM
'       Paulo.Santos
'       Created.
'***************************************************

Public Partial Class icon
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        '*
        '* Get the Icon
        '*
        Dim ico As Drawing.Icon = IconFunctions.GetAssociatedIcon("blablabla." & Request("extension"), IconSize.SystemSmall)
        Dim bmp As New Drawing.Bitmap(ico.Width, ico.Height)
        Dim g As Drawing.Graphics = Drawing.Graphics.FromImage(bmp)
        g.DrawIcon(ico, 0, 0)
        g.Dispose()
        ico.Dispose()

        '*
        '* Save the image
        '*
        Dim stream As New IO.MemoryStream
        bmp.Save(stream, Drawing.Imaging.ImageFormat.Png)
        bmp.Dispose()

        '*
        '* Send the image
        '*
        Response.Clear()
        Response.ContentType = "image/png"
        Response.OutputStream.Write(stream.ToArray(), 0, stream.Length)
        Response.End()

    End Sub

End Class