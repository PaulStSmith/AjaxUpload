'**************************************************
' FILE      : IconFunctions.vb
' AUTHOR    : Paulo.Santos
' CREATION  : 2007.OCT.10
' COPYRIGHT : Copyright © 2007
'             PJ on Development
'             All Rights Reserved.
'
' Description:
'       Function to handle icons in files
'
' Change log:
' 0.1   2007.OCT.10
'       Paulo.Santos
'       Created.
'***************************************************

Imports System.Text
Imports System.Drawing
Imports System.ComponentModel
Imports System.Runtime.InteropServices

Public Module IconFunctions

#Region " Enumerators "

    ''' <summary>
    ''' Enumerates the possible icon sizes.
    ''' </summary>
    Public Enum IconSize
        ''' <summary>
        ''' Denotes a small icon (16x16 pixels).
        ''' </summary>
        SystemSmall

        ''' <summary>
        ''' Denotes a large icon (32x32 pixels)).
        ''' </summary>
        SystemLarge

        ''' <summary>
        ''' Denotes a small icon for the shell (16x16 pixels).
        ''' </summary>
        ShellSmall

        ''' <summary>
        ''' Denotes a large icon for the shell 
        ''' (usually 24x24 pixels or as specified by 
        ''' HKEY_CURRENT_USER\Control Panel\desktop\WindowMetrics\Shell Icon Size)
        ''' </summary>
        ShellLarge

    End Enum

#End Region

#Region " GetNumIcon Method "

    ''' <summary>
    ''' Gets the number of icons of the specified file name.
    ''' </summary>
    ''' <param name="fileName">The name of the file to get the number of icons.</param>
    ''' <returns>The number of icons inside the file.</returns>
    ''' <overloads>Get the number of icons of a file.</overloads>
    Public Function GetNumIcons(ByVal fileName As String) As Integer
        Return ExtractIconEx(fileName, -1, Nothing, Nothing, 0)
    End Function

    ''' <summary>
    ''' Gets the number of icons of the file specified by the <see cref="IO.FileInfo"/> object.
    ''' </summary>
    ''' <param name="file">The <see cref="IO.FileInfo"/> object that holds the information about the file.</param>
    ''' <returns>The number of icons inside the file.</returns>
    Public Function GetNumIcons(ByVal file As IO.FileInfo) As Integer
        Return GetNumIcons(file.FullName)
    End Function

#End Region

#Region " GetAllIcons Methods "

    ''' <summary>
    ''' Extracts all the icons from the file.
    ''' </summary>
    ''' <param name="fileName">The name of the file to extract the icons from.</param>
    ''' <param name="iconSize">The desired size of the icon.</param>
    ''' <returns>An array of icons that contains all the icons within the file.</returns>
    ''' <overloads>Extracts all the icons of a file.</overloads>
    Public Function GetAllIcons(ByVal fileName As String, ByVal iconSize As IconSize) As System.Drawing.Icon()
        If (String.IsNullOrEmpty(fileName)) Then
            Throw New ArgumentNullException("fileName")
        End If
        Dim icons As New List(Of System.Drawing.Icon)
        For i As Integer = 0 To (GetNumIcons(fileName) - 1)
            icons.Add(GetIcon(fileName, iconSize, i))
        Next
        Return icons.ToArray()
    End Function

    ''' <summary>
    ''' Extracts all the icons from the file.
    ''' </summary>
    ''' <param name="file">The <see cref="IO.FileInfo"/> object that holds the information about the file.</param>
    ''' <param name="iconSize">The desired size of the icon.</param>
    ''' <returns>An array of icons that contains all the icons within the file.</returns>
    Public Function GetAllIcons(ByVal file As IO.FileInfo, ByVal iconSize As IconSize) As System.Drawing.Icon()
        If (file Is Nothing) Then
            Throw New ArgumentNullException("file")
        End If
        Return GetAllIcons(file.FullName, iconSize)
    End Function

#End Region

#Region " GetIcon Methods "

    ''' <summary>
    ''' Extracts an icon from the file.
    ''' </summary>
    ''' <param name="fileName">The name of the file to extract the icon from.</param>
    ''' <param name="iconSize">The desired size of the icon.</param>
    ''' <param name="iconIndex">The zero-based index of the icon to extract.</param>
    ''' <returns>The extracted icon.</returns>
    Public Function GetIcon(ByVal fileName As String, ByVal iconSize As IconSize, ByVal iconIndex As Integer) As System.Drawing.Icon
        If (String.IsNullOrEmpty(fileName)) Then
            Throw New ArgumentNullException("fileName")
        End If
        Dim intNumIcons As Integer = GetNumIcons(fileName)
        If (iconIndex >= intNumIcons) Then
            Throw New ArgumentOutOfRangeException("iconIndex", iconIndex, String.Format("iconIndex must be between 0 and {0} for the file '{1}'", intNumIcons, fileName))
        End If

        '*
        '* Get the icon from the file name
        '*
        Dim ptrLargeIcon As IntPtr() = {IntPtr.Zero}
        Dim ptrSmallIcon As IntPtr() = {IntPtr.Zero}

        If (ExtractIconEx(fileName, iconIndex, ptrLargeIcon, ptrSmallIcon, 1) <> 0) Then
            Throw New Exception(String.Format("Error while retrieving icon from '{0}'", fileName))
        End If

        '*
        '* Get the icon at the specified size
        '*
        If (iconSize = IconFunctions.IconSize.SystemLarge) Then
            Return GetManagedIconFromHandle(ptrLargeIcon(0))
        ElseIf (iconSize = IconFunctions.IconSize.SystemSmall) Then
            Return GetManagedIconFromHandle(ptrSmallIcon(0))
        End If

        Return GetManagedIconFromHandle(ptrLargeIcon(0))

    End Function

    ''' <summary>
    ''' Extracts an icon from the file.
    ''' </summary>
    ''' <param name="file">The <see cref="IO.FileInfo"/> object that holds the information about the file.</param>
    ''' <param name="iconSize">The desired size of the icon.</param>
    ''' <param name="iconIndex">The zero-based index of the icon to extract.</param>
    ''' <returns>The extracted icon.</returns>
    Public Function GetIcon(ByVal file As IO.FileInfo, ByVal iconSize As IconSize, ByVal iconIndex As Integer) As System.Drawing.Icon
        Return GetIcon(file.FullName, iconSize, iconIndex)
    End Function

#End Region

    ''' <summary>
    ''' Returns the name and the index of the defualt icon for the file specified by fileName.
    ''' </summary>
    ''' <param name="fileName">The name of the file to return the icon location.</param>
    ''' <returns>The name and the index of the defualt icon.</returns>
    ''' <remarks>
    ''' Not working AT ALL.
    ''' </remarks>
    <Obsolete("Not working property.")> _
    Public Function GetIconLocation(ByVal fileName As String) As String

        Dim fi As New ShFileInfo
        fi.szDisplayName = New String(" ", 260)
        fi.szTypeName = New String(" ", 80)
        Dim i As IntPtr = SHGetFileInfo(fileName, 0, fi, Marshal.SizeOf(fi), ShgfiFlags.IconLocation Or ShgfiFlags.Icon)
        Return (fi.szDisplayName & "," & fi.iIcon.ToString())

    End Function

#Region " GetAssociatedIcon Methods "

    ''' <summary>
    ''' Gets the icon associated with the file name.
    ''' </summary>
    ''' <param name="fileName">The name of the file to extract the icon from.</param>
    ''' <param name="iconSize">The size of the icon to extract.</param>
    ''' <returns>The icon associated with the file name in the specified size.</returns>
    ''' <exception cref="ArgumentNullException">The argument <paramref name="fileName" /> is <langword name="null" />.</exception>
    Public Function GetAssociatedIcon(ByVal fileName As String, ByVal iconSize As IconSize) As System.Drawing.Icon
        If (String.IsNullOrEmpty(fileName)) Then
            Throw New ArgumentNullException("fileName")
        End If

        '*
        '* Use the SHGetFileInfo to get the appropriated icon size
        '*
        Dim hIcon As IntPtr = IntPtr.Zero
        Dim shfi As New ShFileInfo
        Dim ret As IntPtr
        Select Case iconSize
            Case IconFunctions.IconSize.ShellLarge
                ret = SHGetFileInfo(fileName, ShfiFileAttributes.Normal, shfi, Marshal.SizeOf(shfi), ShgfiFlags.UseFileAttributes Or ShgfiFlags.Icon Or ShgfiFlags.ShellIconSize)
            Case IconFunctions.IconSize.ShellSmall
                ret = SHGetFileInfo(fileName, ShfiFileAttributes.Normal, shfi, Marshal.SizeOf(shfi), ShgfiFlags.UseFileAttributes Or ShgfiFlags.Icon Or ShgfiFlags.ShellIconSize Or ShgfiFlags.SmallIcon)
            Case IconFunctions.IconSize.SystemLarge
                ret = SHGetFileInfo(fileName, ShfiFileAttributes.Normal, shfi, Marshal.SizeOf(shfi), ShgfiFlags.UseFileAttributes Or ShgfiFlags.Icon Or ShgfiFlags.LargeIcon)
            Case IconFunctions.IconSize.SystemSmall
                ret = SHGetFileInfo(fileName, ShfiFileAttributes.Normal, shfi, Marshal.SizeOf(shfi), ShgfiFlags.UseFileAttributes Or ShgfiFlags.Icon Or ShgfiFlags.SmallIcon)
        End Select

        '*
        '* Return the icon provided by the SHGetFileInfo
        '*
        If (ret = 0) Then
            '*
            '* Just in case the SHGetFileInfo fails, 
            '* defaults to the ExtractAssociatedIcon function.
            '*
            '* But this adds the problem of only retrieving the 32x32 icon
            '*
            Return System.Drawing.Icon.ExtractAssociatedIcon(fileName)
        End If

        Return GetManagedIconFromHandle(shfi.hIcon)

    End Function

    ''' <summary>
    ''' Gets the icon associated with the file and with the specified width and height.
    ''' </summary>
    ''' <param name="fileName">The name of the file to extract the icon from.</param>
    ''' <param name="width">The width, in pixels, of the desired icon.</param>
    ''' <param name="height">The height, in pixels, of the desired icon.</param>
    ''' <returns>The icon associated with the file.</returns>
    Public Function GetAssociatedIcon(ByVal fileName As String, ByVal width As Integer, ByVal height As Integer) As System.Drawing.Icon
        Return GetAssociatedIcon(fileName, New Drawing.Size(width, height))
    End Function

    ''' <summary>
    ''' Gets the icon associated with the file and with the specified size.
    ''' </summary>
    ''' <param name="fileName">The name of the file to extract the icon from.</param>
    ''' <param name="size">The size of the desired icon.</param>
    ''' <returns>The icon associated with the file.</returns>
    ''' <overloads>Gets the icon associated with a file.</overloads>
    ''' <exception cref="ArgumentNullException">The argument <paramref name="fileName" /> is <langword name="null" />.<para>-or-</para>The argument <paramref name="size" /> is <langword name="null" />.</exception>
    Public Function GetAssociatedIcon(ByVal fileName As String, ByVal size As Drawing.Size) As System.Drawing.Icon
        If (String.IsNullOrEmpty(fileName)) Then
            Throw New ArgumentNullException("fileName")
        End If
        If (size = Nothing OrElse size.IsEmpty) Then
            Throw New ArgumentNullException("size")
        End If

        '*
        '* Try to get the proper icon
        '*
        If (size.Width = size.Height) Then
            If (size.Width = 16) Then
                Return GetAssociatedIcon(fileName, IconSize.SystemSmall)
            ElseIf (size.Width = 32) Then
                Return GetAssociatedIcon(fileName, IconSize.SystemLarge)
            End If
        End If

        '*
        '* Unify the GetAssociatedIcon function into a single one.
        '*
        Return New System.Drawing.Icon(GetAssociatedIcon(fileName, IconSize.SystemLarge), size)
    End Function

    ''' <summary>
    ''' Gets the icon associated with the file specified by the <see cref="IO.FileInfo" /> object.
    ''' </summary>
    ''' <param name="file">The <see cref="IO.FileInfo" /> object that holds the information about the file.</param>
    ''' <param name="iconSize">The size of the icon to extract.</param>
    ''' <returns>The icon associated with the file in the specified size.</returns>
    ''' <exception cref="ArgumentNullException">The argument <paramref name="file" /> is <langword name="null" />.</exception>
    Public Function GetAssociatedIcon(ByVal file As IO.FileInfo, ByVal iconSize As IconSize) As System.Drawing.Icon
        If (file Is Nothing) Then
            Throw New ArgumentNullException("file")
        End If
        Return GetAssociatedIcon(file.FullName, iconSize)
    End Function

    ''' <summary>
    ''' Gets the icon associated with the file specified by the <see cref="IO.FileInfo" /> object with the specified size.
    ''' </summary>
    ''' <param name="file">The <see cref="IO.FileInfo" /> object that holds the information about the file.</param>
    ''' <param name="size">The size of the desired icon.</param>
    ''' <returns>The icon associated with the file with the specified size.</returns>
    ''' <exception cref="ArgumentNullException">The argument <paramref name="file" /> is <langword name="null" />.</exception>
    Public Function GetAssociatedIcon(ByVal file As IO.FileInfo, ByVal size As Drawing.Size) As system.Drawing.Icon
        If (file Is Nothing) Then
            Throw New ArgumentNullException("file")
        End If
        Return GetAssociatedIcon(file.FullName, size)
    End Function

    ''' <summary>
    ''' Gets the icon associated with the file specified by the <see cref="IO.FileInfo" /> object with the specified width and height.
    ''' </summary>
    ''' <param name="file">The <see cref="IO.FileInfo" /> object that holds the information about the file.</param>
    ''' <param name="width">The width, in pixels, of the desired icon.</param>
    ''' <param name="height">The height, in pixels, of the desired icon.</param>
    ''' <returns>The icon associated with the file with the specified width and height.</returns>
    ''' <exception cref="ArgumentNullException">The argument <paramref name="file" /> is <langword name="null" />.</exception>
    Public Function GetAssociatedIcon(ByVal file As IO.FileInfo, ByVal width As Integer, ByVal height As Integer) As System.Drawing.Icon
        If (file Is Nothing) Then
            Throw New ArgumentNullException("file")
        End If
        Return GetAssociatedIcon(file.FullName, width, height)
    End Function

#End Region

#Region " Private Methods "

    ''' <summary>
    ''' Returns a managed version of the icon specified by the unmanaged handle.
    ''' </summary>
    ''' <param name="hIcon">The handle to the unmanaged icon.</param>
    ''' <returns>A managed version of the icon.</returns>
    ''' <remarks>
    ''' This would be a lot simples if the constructor Icon.ctor(IntPtr, Boolean) 
    ''' were available.
    ''' <p/>
    ''' As it is, this constructor is private to the Icon class.
    ''' </remarks>
    Private Function GetManagedIconFromHandle(ByVal hIcon As IntPtr) As System.Drawing.Icon

        Dim ico As System.Drawing.Icon = System.Drawing.Icon.FromHandle(hIcon).Clone()
        DestroyIcon(hIcon)
        Return ico

    End Function

#End Region

End Module
