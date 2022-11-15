'**************************************************
' FILE      : WinAPI.vb
' AUTHOR    : Paulo.Santos
' CREATION  : 2007.10.15
' COPYRIGHT : Copyright © 2007
'             PJ on Development
'             All Rights Reserved.
'
' Description:
'       Public module for Win32API declaration and functions.
'
' Change log:
' 0.1   2007.10.15
'       Paulo.Santos
'       Created.
'***************************************************

Imports System.Text
Imports System.Runtime.InteropServices

''' <summary>
''' Defines several Win32 API functions.
''' </summary>
Friend Module WinAPI

#Region " Icon Functions "

    ''' <summary>
    ''' Enumerates the file attributes for the ShGetFileInfo method.
    ''' </summary>
    Public Enum ShfiFileAttributes As Integer
        ''' <summary>
        ''' The file is read only.
        ''' </summary>
        [ReadOnly] = &H1

        ''' <summary>
        ''' The file is hidden.
        ''' </summary>
        Hidden = &H2

        ''' <summary>
        ''' The file is a system file.
        ''' </summary>
        System = &H4

        ''' <summary>
        ''' The file is a directory.
        ''' </summary>
        Directory = &H10

        ''' <summary>
        ''' The file is an archive.
        ''' </summary>
        Archive = &H20

        ''' <summary>
        ''' The file is a device.
        ''' </summary>
        Device = &H40

        ''' <summary>
        ''' The file has no specia attributes.
        ''' </summary>
        Normal = &H80

        ''' <summary>
        ''' The file is temporary.
        ''' </summary>
        Temporary = &H100

        ''' <summary>
        ''' The file is a sparse file.
        ''' </summary>
        Sparse_File = &H200

        Reparse_Point = &H400

        ''' <summary>
        ''' The file is compressed.
        ''' </summary>
        Compressed = &H800

        ''' <summary>
        ''' The file is offline.
        ''' </summary>
        Offline = &H1000

        Not_Content_Indexed = &H2000

        ''' <summary>
        ''' The file is encrypted.
        ''' </summary>
        Encrypted = &H4000
    End Enum

    ''' <summary>
    ''' Enumerates all the flags used by the ShFileInfo structure.
    ''' </summary>
    Public Enum ShgfiFlags As Integer
        ''' <summary>
        ''' Get large icon.
        ''' </summary>
        LargeIcon = &H0

        ''' <summary>
        ''' Get small icon.
        ''' </summary>
        SmallIcon = &H1
        ''' <summary>
        ''' Get open icon.
        ''' </summary>
        OpenIcon = &H2

        ''' <summary>
        ''' Get shell size icon.
        ''' </summary>
        ''' <remarks></remarks>
        ShellIconSize = &H4

        ''' <summary>
        ''' pszPath is a PIDL.
        ''' </summary>
        PIDL = &H8

        ''' <summary>
        ''' Use passed dwFileAttribute.
        ''' </summary>
        UseFileAttributes = &H10

        ''' <summary>
        ''' Apply the appropriate overlays.
        ''' </summary>
        AddOverlays = &H20

        ''' <summary>
        ''' Get the index of the overlay in the upper 8 bits of the iIcon.
        ''' </summary>
        OverlayIndex = &H40

        ''' <summary>
        ''' Get icon.
        ''' </summary>
        Icon = &H100

        ''' <summary>
        ''' Get display name.
        ''' </summary>
        DisplayName = &H200

        ''' <summary>
        ''' Get type name.
        ''' </summary>
        TypeName = &H400

        ''' <summary>
        ''' Get attributes.
        ''' </summary>
        Attributes = &H800

        ''' <summary>
        ''' Get icon location.
        ''' </summary>
        IconLocation = &H1000

        ''' <summary>
        ''' Return exe type.
        ''' </summary>
        ExeType = &H2000

        ''' <summary>
        ''' Get system icon index.
        ''' </summary>
        SysIconIndex = &H4000

        ''' <summary>
        ''' Put a link overlay on icon.
        ''' </summary>
        LinkOverlay = &H8000

        ''' <summary>
        ''' Show icon in selected state.
        ''' </summary>
        Selected = &H10000

        ''' <summary>
        ''' Get only specified attributes.
        ''' </summary>
        Attr_Specified = &H20000
    End Enum

    ''' <summary>
    ''' Contains information about a file object.
    ''' </summary>
    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto)> _
    Public Structure ShFileInfo
        ''' <summary>
        ''' Handle to the icon that represents the file.
        ''' </summary>
        Public hIcon As IntPtr

        ''' <summary>
        ''' Index of the icon image within the system image list.
        ''' </summary>
        Public iIcon As Integer

        ''' <summary>
        ''' Array of values that indicates the attributes of the file object.
        ''' </summary>
        Public dwAttributes As Integer

        ''' <summary> 
        ''' String that contains the name of the file as it appears in the 
        ''' Microsoft Windows Shell, or the path and file name of the file 
        ''' that contains the icon representing the file.
        ''' </summary>
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=260)> _
        Public szDisplayName As String

        ''' <summary>
        ''' String that describes the type of file.
        ''' </summary>
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=80)> _
        Public szTypeName As String

    End Structure

    ''' <summary>
    ''' Retrieves information about an object in the file system, 
    ''' such as a file, a folder, a directory, or a drive root.
    ''' </summary>
    ''' <param name="pszPath">
    ''' Pointer to a null-terminated string of maximum length 
    ''' MAX_PATH that contains the path and file name. 
    ''' Both absolute and relative paths are valid.
    ''' </param>
    ''' <param name="dwFileAttributes"></param>
    ''' <param name="psfi"></param>
    ''' <param name="cbFileInfo"></param>
    ''' <param name="uFlags"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    <DllImport("shell32.dll", CharSet:=CharSet.Auto)> _
    Public Function SHGetFileInfo(ByVal pszPath As String, ByVal dwFileAttributes As ShfiFileAttributes, ByRef psfi As ShFileInfo, ByVal cbFileInfo As Integer, ByVal uFlags As ShgfiFlags) As IntPtr
    End Function

    ''' <summary>
    ''' Extract an icon from a file.
    ''' </summary>
    ''' <param name="szFileName">The name of the file to extract the icon from.</param>
    ''' <param name="nIconIndex">The index of the icon inside the file.</param>
    ''' <param name="phiconLarge">
    ''' Pointer to an array of icon handles that receives handles to 
    ''' the large icons extracted from the file. If this parameter is NULL, 
    ''' no large icons are extracted from the file.</param>
    ''' <param name="phiconSmall">
    ''' Pointer to an array of icon handles that receives handles to 
    ''' the small icons extracted from the file. If this parameter is NULL, 
    ''' no small icons are extracted from the file.</param>
    ''' <param name="nIcons">Specifies the number of icons to extract from the file.</param>
    ''' <returns>If the <i>nIconIndex</i> parameter is -1, the <i>phiconLarge</i> 
    ''' parameter is NULL, and the <i>phiconSmall</i> parameter is NULL, 
    ''' then the return value is the number of icons contained in the specified file. <br/>
    ''' Otherwise, the return value is the number of icons successfully extracted from the file.</returns>
    <DllImport("shell32.dll", CharSet:=CharSet.Auto)> _
    Public Function ExtractIconEx(ByVal szFileName As String, ByVal nIconIndex As Integer, ByVal phiconLarge() As IntPtr, ByVal phiconSmall() As IntPtr, ByVal nIcons As UInteger) As UInteger
    End Function

    ''' <summary>
    ''' Destroys an icon and frees any memory the icon occupied.
    ''' </summary>
    ''' <param name="hIcon">Handle to the icon to be destroyed. The icon must not be in use.</param>
    ''' <returns>If the function succeeds, returns True.
    ''' <p/>
    ''' If the function fails, returns False. To get extended error information, call GetLastError.</returns>
    <DllImport("user32.dll", SetLastError:=True)> _
    Public Function DestroyIcon(ByVal hIcon As IntPtr) As Boolean
    End Function

#End Region

#Region " SendMessage Variants "

    '*
    '* All the SendMessage variants use 
    '* System.Int32 in order to prevent 
    '* possible errors using this code 
    '* on 64 bits machines.
    '*

#Region " Messages Enumerator "

    ''' <summary>
    ''' Enumerates possible messages.
    ''' </summary>
    ''' <remarks>
    ''' This list will be increased as the need arises.
    ''' </remarks>
    Public Enum SendMessageList As Int32
        ''' <summary>
        ''' Gets the first visible line of multi-line a TextBox or 
        ''' the first visible character of a single-line TextBox.
        ''' </summary>
        EM_GetFirstVisibleLine = &HCE
        ''' <summary>
        ''' Forces a TextBox to scroll an specified number of lines.
        ''' </summary>
        ''' <remarks>
        ''' The number of lines to scroll is specified on the lParam
        ''' and the SendMessage function returns the number of 
        ''' lines actually scrolled.
        ''' </remarks>
        EM_LineScroll = &HB6
    End Enum

#End Region

    ''' <summary>
    ''' Sends the specified message to a window or windows.
    ''' </summary>
    ''' <param name="hWnd">Handle to the window whose window procedure will receive the message.</param>
    ''' <param name="Msg">Specifies the message to be sent.</param>
    ''' <param name="wParam">Specifies additional message-specific information.</param>
    ''' <param name="lParam">Specifies additional message-specific information.</param>
    ''' <returns>The return value specifies the result of the message processing; it depends on the message sent.</returns>
    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)> _
    Public Function SendMessage( _
        ByVal hWnd As HandleRef, _
        ByVal Msg As SendMessageList, _
        ByVal wParam As IntPtr, _
        ByVal lParam As IntPtr) As IntPtr
    End Function

    ''' <summary>
    ''' Sends the specified message to a window or windows.
    ''' </summary>
    ''' <param name="hWnd">Handle to the window whose window procedure will receive the message.</param>
    ''' <param name="Msg">Specifies the message to be sent.</param>
    ''' <param name="wParam">Specifies additional message-specific information.</param>
    ''' <param name="lParam">Specifies additional message-specific information.</param>
    ''' <returns>The return value specifies the result of the message processing; it depends on the message sent.</returns>
    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)> _
    Public Function SendMessage( _
        ByVal hWnd As HandleRef, _
        ByVal Msg As SendMessageList, _
        ByVal wParam As Int32, _
        ByVal lParam As Int32) As Int32
    End Function

    ''' <summary>
    ''' Sends the specified message to a window or windows.
    ''' </summary>
    ''' <param name="hWnd">Handle to the window whose window procedure will receive the message.</param>
    ''' <param name="Msg">Specifies the message to be sent.</param>
    ''' <param name="wParam">Specifies additional message-specific information.</param>
    ''' <param name="lParam">Specifies additional message-specific information.</param>
    ''' <returns>The return value specifies the result of the message processing; it depends on the message sent.</returns>
    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)> _
    Public Function SendMessage( _
        ByVal hWnd As IntPtr, _
        ByVal Msg As SendMessageList, _
        ByVal wParam As IntPtr, _
        ByVal lParam As IntPtr) As IntPtr
    End Function

    ''' <summary>
    ''' Sends the specified message to a window or windows.
    ''' </summary>
    ''' <param name="hWnd">Handle to the window whose window procedure will receive the message.</param>
    ''' <param name="Msg">Specifies the message to be sent.</param>
    ''' <param name="wParam">Specifies additional message-specific information.</param>
    ''' <param name="lParam">Specifies additional message-specific information.</param>
    ''' <returns>The return value specifies the result of the message processing; it depends on the message sent.</returns>
    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)> _
    Public Function SendMessage( _
        ByVal hWnd As IntPtr, _
        ByVal Msg As SendMessageList, _
        ByVal wParam As Int32, _
        ByVal lParam As Int32) As Int32
    End Function

    ''' <summary>
    ''' Sends the specified message to a window or windows.
    ''' </summary>
    ''' <param name="hWnd">Handle to the window whose window procedure will receive the message.</param>
    ''' <param name="Msg">Specifies the message to be sent.</param>
    ''' <param name="wParam">Specifies additional message-specific information.</param>
    ''' <param name="lParam">Specifies additional message-specific information.</param>
    ''' <returns>The return value specifies the result of the message processing; it depends on the message sent.</returns>
    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)> _
    Public Function SendMessage( _
        ByVal hWnd As HandleRef, _
        ByVal Msg As SendMessageList, _
        ByVal wParam As IntPtr, _
        ByRef lParam As StringBuilder) As IntPtr
    End Function

    ''' <summary>
    ''' Sends the specified message to a window or windows.
    ''' </summary>
    ''' <param name="hWnd">Handle to the window whose window procedure will receive the message.</param>
    ''' <param name="Msg">Specifies the message to be sent.</param>
    ''' <param name="wParam">Specifies additional message-specific information.</param>
    ''' <param name="lParam">Specifies additional message-specific information.</param>
    ''' <returns>The return value specifies the result of the message processing; it depends on the message sent.</returns>
    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)> _
    Public Function SendMessage( _
        ByVal hWnd As HandleRef, _
        ByVal Msg As SendMessageList, _
        ByVal wParam As Int32, _
        ByRef lParam As StringBuilder) As Int32
    End Function

    ''' <summary>
    ''' Sends the specified message to a window or windows.
    ''' </summary>
    ''' <param name="hWnd">Handle to the window whose window procedure will receive the message.</param>
    ''' <param name="Msg">Specifies the message to be sent.</param>
    ''' <param name="wParam">Specifies additional message-specific information.</param>
    ''' <param name="lParam">Specifies additional message-specific information.</param>
    ''' <returns>The return value specifies the result of the message processing; it depends on the message sent.</returns>
    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)> _
    Public Function SendMessage( _
        ByVal hWnd As IntPtr, _
        ByVal Msg As SendMessageList, _
        ByVal wParam As IntPtr, _
        ByRef lParam As StringBuilder) As IntPtr
    End Function

    ''' <summary>
    ''' Sends the specified message to a window or windows.
    ''' </summary>
    ''' <param name="hWnd">Handle to the window whose window procedure will receive the message.</param>
    ''' <param name="Msg">Specifies the message to be sent.</param>
    ''' <param name="wParam">Specifies additional message-specific information.</param>
    ''' <param name="lParam">Specifies additional message-specific information.</param>
    ''' <returns>The return value specifies the result of the message processing; it depends on the message sent.</returns>
    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)> _
    Public Function SendMessage( _
        ByVal hWnd As IntPtr, _
        ByVal Msg As SendMessageList, _
        ByVal wParam As Int32, _
        ByRef lParam As StringBuilder) As Int32
    End Function

#End Region

End Module
