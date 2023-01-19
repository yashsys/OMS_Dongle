Option Strict Off
Option Explicit On
Module modFTP
    '********************************** START FTP API DECLARATION
    Const FTP_TRANSFER_TYPE_UNKNOWN As Integer = &H0
    Const FTP_TRANSFER_TYPE_ASCII As Integer = &H1
    Const FTP_TRANSFER_TYPE_BINARY As Integer = &H2
    Const INTERNET_DEFAULT_FTP_PORT As Short = 21 ' default for FTP servers
    Const INTERNET_SERVICE_FTP As Short = 1
    Const INTERNET_FLAG_PASSIVE As Integer = &H8000000 ' used for FTP connections
    Const INTERNET_OPEN_TYPE_PRECONFIG As Short = 0 ' use registry configuration
    Const INTERNET_OPEN_TYPE_DIRECT As Short = 1 ' direct to net
    Const INTERNET_OPEN_TYPE_PROXY As Short = 3 ' via named proxy
    Const INTERNET_OPEN_TYPE_PRECONFIG_WITH_NO_AUTOPROXY As Short = 4 ' prevent using java/script/INS
    Const MAX_PATH As Short = 260

    Private Structure WIN32_FIND_DATA
        Dim dwFileAttributes As Integer
        Dim ftCreationTime As FILETIME
        Dim ftLastAccessTime As FILETIME
        Dim ftLastWriteTime As FILETIME
        Dim nFileSizeHigh As Integer
        Dim nFileSizeLow As Integer
        Dim dwReserved0 As Integer
        Dim dwReserved1 As Integer
        'UPGRADE_WARNING: Fixed-length string size must fit in the buffer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
        <VBFixedString(MAX_PATH), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=MAX_PATH)> Public cFileName() As Char
        'UPGRADE_WARNING: Fixed-length string size must fit in the buffer. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="3C1E4426-0B80-443E-B943-0627CD55D48B"'
        <VBFixedString(14), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValArray, SizeConst:=14)> Public cAlternate() As Char
    End Structure

    Private Structure FILETIME
        Dim dwLowDateTime As Integer
        Dim dwHighDateTime As Integer
    End Structure

    Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Integer) As Short
    Private Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" (ByVal hInternetSession As Integer, ByVal sServerName As String, ByVal nServerPort As Short, ByVal sUserName As String, ByVal sPassword As String, ByVal lService As Integer, ByVal lFlags As Integer, ByVal lContext As Integer) As Integer
    Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Integer, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Integer) As Integer
    Private Declare Function FtpSetCurrentDirectory Lib "wininet.dll" Alias "FtpSetCurrentDirectoryA" (ByVal hFtpSession As Integer, ByVal lpszDirectory As String) As Boolean
    Private Declare Function FtpGetCurrentDirectory Lib "wininet.dll" Alias "FtpGetCurrentDirectoryA" (ByVal hFtpSession As Integer, ByVal lpszCurrentDirectory As String, ByRef lpdwCurrentDirectory As Integer) As Integer
    Private Declare Function FtpCreateDirectory Lib "wininet.dll" Alias "FtpCreateDirectoryA" (ByVal hFtpSession As Integer, ByVal lpszDirectory As String) As Boolean
    Private Declare Function FtpRemoveDirectory Lib "wininet.dll" Alias "FtpRemoveDirectoryA" (ByVal hFtpSession As Integer, ByVal lpszDirectory As String) As Boolean
    Private Declare Function FtpDeleteFile Lib "wininet.dll" Alias "FtpDeleteFileA" (ByVal hFtpSession As Integer, ByVal lpszFileName As String) As Boolean
    'UPGRADE_WARNING: Structure WIN32_FIND_DATA may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
    Private Declare Function FtpFindFirstFile Lib "wininet.dll" Alias "FtpFindFirstFileA" (ByVal hFtpSession As Integer, ByVal lpszSearchFile As String, ByRef lpFindFileData As WIN32_FIND_DATA, ByVal dwFlags As Integer, ByVal dwContent As Integer) As Integer
    'UPGRADE_WARNING: Structure WIN32_FIND_DATA may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
    Private Declare Function InternetFindNextFile Lib "wininet.dll" Alias "InternetFindNextFileA" (ByVal hFind As Integer, ByRef lpvFindData As WIN32_FIND_DATA) As Integer
    Private Declare Function FtpRenameFile Lib "wininet.dll" Alias "FtpRenameFileA" (ByVal hFtpSession As Integer, ByVal lpszExisting As String, ByVal lpszNew As String) As Boolean
    Private Declare Function FtpGetFile Lib "wininet.dll" Alias "FtpGetFileA" (ByVal hConnect As Integer, ByVal lpszRemoteFile As String, ByVal lpszNewFile As String, ByVal fFailIfExists As Integer, ByVal dwFlagsAndAttributes As Integer, ByVal dwFlags As Integer, ByRef dwContext As Integer) As Boolean
    Private Declare Function FtpPutFile Lib "wininet.dll" Alias "FtpPutFileA" (ByVal hConnect As Integer, ByVal lpszLocalFile As String, ByVal lpszNewRemoteFile As String, ByVal dwFlags As Integer, ByVal dwContext As Integer) As Boolean
    Private Declare Function InternetGetLastResponseInfo Lib "wininet.dll" Alias "InternetGetLastResponseInfoA" (ByRef lpdwError As Integer, ByVal lpszBuffer As String, ByRef lpdwBufferLength As Integer) As Boolean
    'Const PassiveConnection As Boolean = True

    Private Const INTERNET_FLAG_RELOAD As Integer = &H80000000
    Private Const INTERNET_FLAG_NO_CACHE_WRITE As Integer = &H4000000
    Private Const FILE_ATTRIBUTE_ARCHIVE As Long = &H20
    '********************************** END FTP API DECLARATION
    Private Function GetServerResponse() As String
        Dim lError As Long
        Dim strBuffer As String = ""
        Dim lBufferSize As Long
        Dim retVal As Long
        retVal = InternetGetLastResponseInfo(lError, strBuffer, lBufferSize)
        strBuffer = New String("", lBufferSize + 1)
        retVal = InternetGetLastResponseInfo(lError, strBuffer, lBufferSize)
        GetServerResponse = strBuffer
    End Function

    Public Function Get_FTP_File(ByVal strFTP_Address As String, ByVal strFTP_User_Name As String, ByVal strFTP_User_Password As String, ByVal PassiveConnection As Boolean, ByVal strFTP_File As String, ByVal strLocal_File As String) As Boolean
        Dim hConnection, hOpen As Integer
        'open an internet connection
        hOpen = InternetOpen("API-Guide sample program", INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
        'connect to the FTP server
        hConnection = InternetConnect(hOpen, strFTP_Address, INTERNET_DEFAULT_FTP_PORT, strFTP_User_Name, strFTP_User_Password, INTERNET_SERVICE_FTP, IIf(PassiveConnection, INTERNET_FLAG_PASSIVE, 0), 0)
        If hConnection = 0 Then
            Get_FTP_File = False
            Exit Function
        End If
        If FtpGetFile(hConnection, strFTP_File, strLocal_File, 0, 0, &H1, 0) = True Then
            Get_FTP_File = True
        Else
            Get_FTP_File = False
        End If
        'close the FTP connection
        InternetCloseHandle(hConnection)
        'close the internet connection
        InternetCloseHandle(hOpen)
    End Function

    Public Function Drop_FTP_File(ByVal strFTP_Address As String, ByVal strFTP_User_Name As String, ByVal strFTP_User_Password As String, ByVal PassiveConnection As Boolean, ByVal strFTP_File As String) As Boolean
        Dim hConnection, hOpen As Integer
        'open an internet connection
        hOpen = InternetOpen("API-Guide sample program", INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
        'connect to the FTP server
        hConnection = InternetConnect(hOpen, strFTP_Address, INTERNET_DEFAULT_FTP_PORT, strFTP_User_Name, strFTP_User_Password, INTERNET_SERVICE_FTP, IIf(PassiveConnection, INTERNET_FLAG_PASSIVE, 0), 0)
        If hConnection = 0 Then
            Drop_FTP_File = False
            Exit Function
        End If
        If FtpDeleteFile(hConnection, strFTP_File) = True Then
            Drop_FTP_File = True
        Else
            Drop_FTP_File = False
        End If
        'close the FTP connection
        InternetCloseHandle(hConnection)
        'close the internet connection
        InternetCloseHandle(hOpen)
    End Function

    Public Function Put_FTP_Files(ByVal strFTP_Address As String, ByVal strFTP_User_Name As String, ByVal strFTP_User_Password As String, ByVal PassiveConnection As Boolean, ByVal strFTP_Path As String, ByVal strFile_Path As String, Optional ByVal blnRename As Boolean = False) As Boolean
        Dim hConnection, hOpen As Integer
        Dim sOrgPath As String
        Dim strFile_Name As String
        Dim lngCounter As Long

        'open an internet connection
        hOpen = InternetOpen("API-Guide sample program", INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
        'connect to the FTP server
        hConnection = InternetConnect(hOpen, strFTP_Address, INTERNET_DEFAULT_FTP_PORT, strFTP_User_Name, strFTP_User_Password, INTERNET_SERVICE_FTP, IIf(PassiveConnection, INTERNET_FLAG_PASSIVE, 0), 0)
        If hConnection = 0 Then
            Put_FTP_Files = False
            Exit Function
        End If
        'create a buffer to store the original directory
        sOrgPath = New String(Chr(0), MAX_PATH)
        'get the directory
        FtpGetCurrentDirectory(hConnection, sOrgPath, Len(sOrgPath))
        Dim str() As String

        str = Split(strFTP_Path, "/", -1, CompareMethod.Text)
        For lngCounter = 0 To UBound(str)
            If str(lngCounter) <> "" Then
                FtpCreateDirectory(hConnection, str(lngCounter))
                FtpSetCurrentDirectory(hConnection, str(lngCounter))
            End If
        Next
        'upload the file 'test.htm'
        strFile_Name = Retrive_File_Name(strFile_Path)
        If FtpPutFile(hConnection, strFile_Path, strFile_Name, FTP_TRANSFER_TYPE_UNKNOWN, 0) = False Then
            Put_FTP_Files = False
            Exit Function
        End If
        If blnRename = True Then
            Drop_FTP_File(strFTP_Address, strFTP_User_Name, strFTP_User_Password, PassiveConnection, strFTP_Path & "/" & Replace(strFile_Name, ".tmp", ".BAK"))
            FtpRenameFile(hConnection, strFile_Name, Replace(strFile_Name, ".tmp", ".BAK"))
            'Rename_FTP_File("ftp://" & strFTP_Address & strDT_FTP_Folder_Path & "/IMPORT/" & glngFinancial_Year & "/" & strDT_Company_Id & "/" & strCurrentFile, Replace(strCurrentFile, ".tmp", ".BAK"))
        End If
        'close the FTP connection
        InternetCloseHandle(hConnection)
        'close the internet connection
        InternetCloseHandle(hOpen)

        Put_FTP_Files = True
    End Function

    Private Function StripNull(item As String)
        'Return a string without the chr$(0) terminator.
        Dim pos As Integer
        pos = InStr(item, Chr(0))
        If pos Then
            StripNull = Left(item, pos - 1)
        Else
            StripNull = item
        End If
    End Function

    Public Function Remove_FTP_Folder(ByVal strFTP_Address As String, ByVal strFTP_User_Name As String, ByVal strFTP_User_Password As String, ByVal PassiveConnection As Boolean, ByVal strFTP_Folder As String) As Boolean
        'Remove_Files_From_Folder(strFTP_Address, strFTP_User_Name, strFTP_User_Password, strFTP_Folder)
        Dim hConnection, hOpen As Integer
        Dim sOrgPath As String

        'open an internet connection
        hOpen = InternetOpen("API-Guide sample program", INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
        'connect to the FTP server
        hConnection = InternetConnect(hOpen, strFTP_Address, INTERNET_DEFAULT_FTP_PORT, strFTP_User_Name, strFTP_User_Password, INTERNET_SERVICE_FTP, IIf(PassiveConnection, INTERNET_FLAG_PASSIVE, 0), 0)
        If hConnection = 0 Then
            Remove_FTP_Folder = False
            Exit Function
        End If
        'create a buffer to store the original directory
        sOrgPath = New String(Chr(0), MAX_PATH)
        'get the directory
        FtpGetCurrentDirectory(hConnection, sOrgPath, Len(sOrgPath))
        If FtpRemoveDirectory(hConnection, strFTP_Folder) = False Then
            'MsgBox(GetServerResponse())
        End If
        'close the FTP connection
        InternetCloseHandle(hConnection)
        'close the internet connection
        InternetCloseHandle(hOpen)
        Remove_FTP_Folder = True
    End Function

    Public Function Remove_Files_From_Folder(ByVal strFTP_Address As String, ByVal strFTP_User_Name As String, ByVal strFTP_User_Password As String, ByVal PassiveConnection As Boolean, ByVal strFTP_Folder As String) As Boolean
        Dim WFD As WIN32_FIND_DATA
        Dim tmp As String
        Dim hConnection, hOpen As Integer
        Dim sOrgPath As String
        Dim hFind As Long

        'open an internet connection
        hOpen = InternetOpen("API-Guide sample program", INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
        'connect to the FTP server
        hConnection = InternetConnect(hOpen, strFTP_Address, INTERNET_DEFAULT_FTP_PORT, strFTP_User_Name, strFTP_User_Password, INTERNET_SERVICE_FTP, IIf(PassiveConnection, INTERNET_FLAG_PASSIVE, 0), 0)
        If hConnection = 0 Then
            Remove_Files_From_Folder = False
            Exit Function
        End If
        'create a buffer to store the original directory
        sOrgPath = New String(Chr(0), MAX_PATH)
        'get the directory
        FtpGetCurrentDirectory(hConnection, sOrgPath, Len(sOrgPath))
        'set the current directory to 'root/testing'
        FtpSetCurrentDirectory(hConnection, strFTP_Folder)
        FtpGetCurrentDirectory(hConnection, sOrgPath, Len(sOrgPath))

        hFind = FtpFindFirstFile(hConnection, sOrgPath, WFD, INTERNET_FLAG_RELOAD Or INTERNET_FLAG_NO_CACHE_WRITE, 0&)
        If hFind Then
            Do
                tmp = StripNull(WFD.cFileName)
                If Len(tmp) Then
                    If WFD.dwFileAttributes And vbDirectory Then
                    Else
                        FtpDeleteFile(hConnection, tmp)
                    End If
                End If
                'continue while valid
            Loop While InternetFindNextFile(hFind, WFD)
        End If 'If hFind
        'close the FTP connection
        InternetCloseHandle(hConnection)
        'close the internet connection
        InternetCloseHandle(hOpen)
        Remove_Files_From_Folder = True
    End Function

    Public Function FTP_Folder_Files_List(ByVal strFTP_Address As String, ByVal strFTP_User_Name As String, ByVal strFTP_User_Password As String, ByVal PassiveConnection As Boolean, ByVal strFTP_Folder As String, ByRef Dirlist As List(Of String)) As Boolean
        Dim WFD As WIN32_FIND_DATA
        Dim tmp As String
        Dim hConnection, hOpen As Integer
        Dim sOrgPath As String
        Dim hFind As Long

        'open an internet connection
        hOpen = InternetOpen("API-Guide sample program", INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
        'connect to the FTP server
        hConnection = InternetConnect(hOpen, strFTP_Address, INTERNET_DEFAULT_FTP_PORT, strFTP_User_Name, strFTP_User_Password, INTERNET_SERVICE_FTP, IIf(PassiveConnection, INTERNET_FLAG_PASSIVE, 0), 0)
        If hConnection = 0 Then
            FTP_Folder_Files_List = False
            Exit Function
        End If
        'create a buffer to store the original directory
        sOrgPath = New String(Chr(0), MAX_PATH)
        'get the directory
        FtpGetCurrentDirectory(hConnection, sOrgPath, Len(sOrgPath))
        'set the current directory to 'root/testing'
        FtpSetCurrentDirectory(hConnection, strFTP_Folder)
        FtpGetCurrentDirectory(hConnection, sOrgPath, Len(sOrgPath))

        hFind = FtpFindFirstFile(hConnection, sOrgPath, WFD, INTERNET_FLAG_RELOAD Or INTERNET_FLAG_NO_CACHE_WRITE, 0&)
        If hFind Then
            Do
                tmp = StripNull(WFD.cFileName)
                If Len(tmp) Then
                    If WFD.dwFileAttributes And vbDirectory Then
                    Else
                        Dirlist.Add(tmp)
                    End If
                End If
                'continue while valid
            Loop While InternetFindNextFile(hFind, WFD)
        End If 'If hFind
        'close the FTP connection
        InternetCloseHandle(hConnection)
        'close the internet connection
        InternetCloseHandle(hOpen)
        FTP_Folder_Files_List = True
    End Function

    Public Function FTP_Folder_List(ByVal strFTP_Address As String, ByVal strFTP_User_Name As String, ByVal strFTP_User_Password As String, ByVal PassiveConnection As Boolean, ByVal strFTP_Folder As String, ByRef strCompanies As String) As Boolean
        Dim WFD As WIN32_FIND_DATA
        Dim tmp As String
        Dim hConnection, hOpen As Integer
        Dim sOrgPath As String
        Dim hFind As Long

        'open an internet connection
        hOpen = InternetOpen("API-Guide sample program", INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
        'connect to the FTP server
        hConnection = InternetConnect(hOpen, strFTP_Address, INTERNET_DEFAULT_FTP_PORT, strFTP_User_Name, strFTP_User_Password, INTERNET_SERVICE_FTP, IIf(PassiveConnection, INTERNET_FLAG_PASSIVE, 0), 0)
        If hConnection = 0 Then
            FTP_Folder_List = False
            Exit Function
        End If
        'create a buffer to store the original directory
        sOrgPath = New String(Chr(0), MAX_PATH)
        'get the directory
        FtpGetCurrentDirectory(hConnection, sOrgPath, Len(sOrgPath))
        'set the current directory to 'root/testing'
        FtpSetCurrentDirectory(hConnection, strFTP_Folder)
        FtpGetCurrentDirectory(hConnection, sOrgPath, Len(sOrgPath))
        strCompanies = ""
        hFind = FtpFindFirstFile(hConnection, sOrgPath, WFD, INTERNET_FLAG_RELOAD Or INTERNET_FLAG_NO_CACHE_WRITE, 0&)
        If hFind Then
            Do
                tmp = StripNull(WFD.cFileName)
                If Len(tmp) = 4 Then
                    If WFD.dwFileAttributes And vbDirectory Then
                        If strCompanies = "" Then
                            strCompanies = "'" & tmp & "'"
                        Else
                            strCompanies = strCompanies & ",'" & tmp & "'"
                        End If
                    End If
                End If
                'continue while valid
            Loop While InternetFindNextFile(hFind, WFD)
        End If 'If hFind
        'close the FTP connection
        InternetCloseHandle(hConnection)
        'close the internet connection
        InternetCloseHandle(hOpen)
        FTP_Folder_List = True
    End Function

    Public Function Retrive_File_Name(ByVal strPath As String) As String
        Dim intCounter As Short
        Dim intXCounter As Short

        If InStr(1, strPath, "\\") > 0 Then
            intCounter = InStr(3, strPath, "\") + 1
        Else
            intCounter = 1
        End If
        Do While InStr(intCounter, strPath, "\") > 0
            intCounter = InStr(intCounter, strPath, "\") + 1
            intXCounter = intCounter - 1
        Loop
        Retrive_File_Name = Mid(strPath, intCounter, Len(strPath))
    End Function
End Module