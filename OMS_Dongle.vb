' ================================================ For Sending SMS ================================================
' We use the HttpUtility class from the System.Web namespace
'
' If you see of the error "'HttpUtility' is not declared", you are probably
' using a newer version of Visual Studio. You need to navigate to
' Project | <Project name> Properties | Compile | Advanced Compiler Settings,
' and select e.g. ".NET Framework 4" instead of ".NET Framework 4 Client Profile".
'
' Next, visit Project | Add reference, and select "System.Web" (specifically
' System.Web - not System.Web.<something>).

'Imports System.Web
'Imports System.IO
'Imports System.Net
'Imports System.Text
'Imports System.Collections
' ================================================ For Sending SMS ================================================
'Imports System.Threading

Imports System.ServiceProcess
Imports System.Configuration.Install
Imports System.Text
Imports System.IO
Imports System.Data.SqlClient
Imports System.Net
Imports System.Net.Security
Imports System.Security.Cryptography.X509Certificates
Imports System.Net.Mail
Imports System.Runtime.CompilerServices
Imports Microsoft.Win32

Imports System
Imports System.Data
Imports System.Collections
Imports System.Reflection

Public Class OMS_Dongle
    Inherits System.ServiceProcess.ServiceBase
    Public Const gstrAllowed_Multiple_Companies_LockIds As String = "DISK,1917058163,164941920,666002111,210241960,503194557,1860805334,669405361" ',728258899
    Public gobjLock_Type As Integer
    Public gstrAsyncronous_String As String
    Public gstrCompany_Desc_Block As String
    Public gstrCompanies_Block As String
    Public gstrHasp_LockId As String
    Public gstrCompany_Block1 As String
    Public gstrCompany_Block2 As String
    Public gstrCompany_Block3 As String
    Public gstrLicense_Information As String
    Public gstrSMS_Message As String
    Private Declare Auto Function GetPrivateProfileString Lib "kernel32" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As StringBuilder, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
    Private strHO_Company_Code As String, strHO_Desc As String, strBranch_Desc As String, strOpening_Date As String, strExe_Version As String

    Private gintCompany_Type As Integer
    Private gintSAP_Hourly As Integer
    Private gintProcess_SAP_Inbound_After_Dayend As Integer

    Private gstrHourly_Inbound_File As String
    Private gstrShift_Counter_Department_User As String
    Private gstrServer_IP_Address As String
    Private gintServer_Type As Integer
    Private gstrServer_Name As String
    Private gstrSQL_Server_Instance_Name As String
    Private gstrSQL_Server_Port As String
    Private gstrOffline_System_SQL_Server_Instance_Name As String
    Private gstrShared_Folder As String
    Private gstrSQL_Instance_User_Name As String
    Private gstrPublication_Database As String
    Friend WithEvents Timer1 As System.Timers.Timer
    Private blnBOD As Boolean
    Private strBOD_Date As String
    Private Dirlist As List(Of String)
    Private blnRecursive As Boolean
    Private blnFirst_Col As Boolean
    Private strSQL_String As String, strBackup_File As String, strField_Name As String
    Private strLinked_Server As String
    Private strTable As String
    Private intFTP_Disable_PM As Integer
    Private adoSAP As SqlConnection

    Private adapter As SqlDataAdapter
    Private adoCon_Client As SqlConnection
    Private adoCon_Company As SqlConnection

    Private blnFirst_File As Boolean
    Private command As SqlCommand
    Private gintAttachment As Integer
    Private adoRS As DataSet
    Private adoRs_Table As DataSet
    Private trans_Server As SqlTransaction
    Private trans_Client As SqlTransaction
    Private connetionString As String
    Private intcounter As Integer
    Private gstrGSB_Department As String
    Private gstrSMTP_Report As String, gstrMobile_Message As String
    Private strFTP_Files() As String
    Private strFTP_Address As String, strFTP_User_Name As String, strFTP_User_Password As String, strSecondary_FTP_Path As String
    Private strSMTP_Host As String, strSMTP_Port As String, strSMTP_User_Name As String, strSMTP_User_Password As String, strSMTP_Subject As String, strSMTP_To_Addresses As String, strFile_Name As String
    Private strSMTP_CC_Addresses As String, strSMTP_BCC_Address As String, strSMSDeveloperAPI As String, strSMS_User_Name As String, strSMS_User_Password As String, strSMSMobileNo1 As String, strSMSMobileNo2 As String, strSMSMobileNo3 As String, strSMSMobileNo4 As String, strSMSMobileNo5 As String, strSMSMobileNo6 As String, strSMSMobileNo7 As String, strSMSMobileNo8 As String, strSMSMobileNo9 As String, strSMSMobileNo10 As String, strUpload_After_Every As String
    Private strLocal_Stock_File_Path As String
    Private intEvent_Full_Backup As Integer, intEvent_Diff_Backup As Integer, intEvent_Log_Backup As Integer, intBackup_Info_Through As Integer, intEvent_DayEnd As Integer, intDayEnd_Info_Through As Integer, intEvent_SAP_Inbound As Integer, intEvent_SAP_Outbound As Integer, intSend_Birthday_Wishes As Integer, intBirthday_Info_Through As Integer, intSend_Invoice_Info As Integer, intInvoice_Info_Through As Integer, intReIndex_Database As Integer, intAEL As Integer, intUpload_Negative_Stock As Integer, intUpload_Branch As Integer, intSend_ALL_Emails_within As Integer

    Private strHO_IP_Address As String, strHO_Server_Name As String, strHO_SQL_Server_Instance_Name As String, strDT_Company_Id As String, strDT_HO_Location_Id As String, strDT_HO_Database As String, strDTB_FTP_User_Password As String, strDT_Start_Time As String, strDT_Master As String, strDT_Transaction As String
    Private dtDT_Schedule_Date As Date, dtDT_Schedule_Start_Date As Date, dtDT_Time_LED As Date
    Private intDT_Schedule_Type As Integer, intHO_SQL_Port As Integer, intDT_After_Dayend As Integer
    Private strSAP_Indicator As String, strTax_Header As String, dtGST_DATE As String, strDT_FTP_Folder_Path As String
    Private strFile_Header_Name_Of_Table_Structure As String, strFile_Header_Direction As String, strFile_Header_Name_Of_Basic_Type As String, strFile_Header_Message_Type As String, strFile_Header_Sender_Port As String, strFile_Header_Partner_Type_Of_Sender As String, strFile_Header_Logical_Address_Of_Sender As String, strFile_Header_Receiver_Port As String, strFile_Header_Partner_Type_Of_Receiver As String, strBill_Header_Name_Of_Table_Structure As String, strBill_Header_Currency_Code As String, strItem_Record_Name_Of_Item_Row_Structure As String, strItem_Record_Transaction_Type_In_POS_System As String, strItem_Record_Qualifier_For_Following_Fields As String, strItem_Record_Name_Of_MRP_Row_Structure As String, strItem_Record_Type_Of_Condition_Discount As String, strItem_Record_Name_Of_Rate_Difference_Row_Structure As String, strItem_Record_Discount_At_POS_Level As String, strName_Of_VAT_Row_Structure As String, strGS_TABNAM As String, strName_Of_Addon_Row_Structure As String, strPayment_Row_Credit_Sale As String, strPayment_Row_Name_Of_Table_Structure As String, strAddon_Row_Type_Of_Discout_For_Roff As String, strAddon_Row_Type_Of_Disc_For_Other_Addon As String
    Private gstrCompany_Id As String, gstrCompany_Name As String, gstrCompany_Address As String, glngFinancial_Year As Long, strSite_Id As String, strClient_Id As String, strLogical_Address_Of_Sender As String, strReceiver_Port As String, strExport_Item_With_Prefix As String, strLocation_ID As String, strHead_Office_Id As String, strFrom_Time As String, strTo_Time As String, strInbound_File_Path As String, strInBound_Format As String, strInBound_Directory As String, intSchedule_Type As Integer, dtInBound_Schedule_Date As Date, dtInBound_Schedule_Start_Date As Date, strBillwise_Time As String, strSummary_Time As String, strCSV_Format_Time As String, strCSV_Summary_Time As String, blnSchedule_Time As Boolean, intInbound_Output_Format As Integer, strFile_Header_Partner_Number_Of_Receiver As String, strFTP_Out_Bound_Folder As String, dtBillwise_Time_LED As Date, dtSummary_Time_LED As Date, dtCSV_Format_Time_LED As Date, dtCSV_Summary_Time_LED As Date, strSender_Port As String, strSAP_Ind_CAPA As String, strSAP_Ind_CAPB As String, strSale_Not_Found_Email As String
    Private gstrLock_Company_Id As String
    Private dtVoucher_Date As Date
    Private intNo_Of_Invoices_In_Each_File As Integer, intNo_Of_Items_In_Each_File As Integer, intInvoicing_As_Per_SAP_With_Prefix As Integer, intAPLY_GST As Integer, intRestore_Type_DBK As Integer
    Private objfile_Stream As FileStream
    Private objStreamWriter As StreamWriter
    Private blnUpdate_Last_Char As Boolean
    Private intTotal_Records As Integer, intRecord_Counter As Integer
    Private objStreamReader As StreamReader
    'Private MyLog As New EventLog() ' create a new event log 
    Private outputZip As String '= "output zip file path"
    Private inputZip As String '= "input zip file path"
    Private inputFolder As String '= "input folder path"
    Private outputFolder As String '= "output folder path"
    Private strMonth As String
    Private strFTP_Backup_Folder As String, strDelete_FTP_Backup_Folder As String
    Private intBackup_Sequence_Control As Integer
    Private strFTP_ADDR(10) As String, strFTP_USER(10) As String, strFTP_PASS(10) As String
    Private intFTP_Type(10) As Integer
    Private strError As String

    Private colFiles As Collection
    '*********************FTP START
    Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer
    Private Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" (ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, ByVal sUserName As String, ByVal sPassword As String, ByVal lService As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
    Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal ByVallFlags As Long) As Long
    Private Declare Function FtpSetCurrentDirectory Lib "wininet.dll" Alias "FtpSetCurrentDirectoryA" (ByVal ByValhFtpSession As Long, ByVal lpszDirectory As String) As Boolean
    Private Declare Function FtpGetCurrentDirectory Lib "wininet.dll" Alias "FtpGetCurrentDirectoryA" (ByVal ByValhFtpSession As Long, ByVal lpszCurrentDirectory As String, ByVal lpdwCurrentDirectory As Long) As Long
    Private Declare Function FtpCreateDirectory Lib "wininet.dll" Alias "FtpCreateDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean
    Private Declare Function FtpRemoveDirectory Lib "wininet.dll" Alias "FtpRemoveDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean
    Private Declare Function FtpDeleteFile Lib "wininet.dll" Alias "FtpDeleteFileA" (ByVal hFtpSession As Long, ByVal lpszFileName As String) As Boolean
    Private Declare Function FtpRenameFile Lib "wininet.dll" Alias "FtpRenameFileA" (ByVal hFtpSession As Long, ByVal lpszExisting As String, ByVal lpszNew As String) As Boolean
    Private Declare Function FtpGetFile Lib "wininet.dll" Alias "FtpGetFileA" (ByVal hConnect As Long, ByVal ByVallpszRemoteFile As String, ByVal lpszNewFile As String, ByVal fFailIfExists As Long, ByVal dwFlagsAndAttributes As Long, ByVal dwFlags As Long, ByRef dwContext As Long) As Boolean
    Private Declare Function FtpPutFile Lib "wininet.dll" Alias "FtpPutFileA" (ByVal hConnect As Long, ByVal lpszLocalFile As String, ByVal lpszNewRemoteFile As String, ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean
    Private Declare Function InternetGetLastResponseInfo Lib "wininet.dll" Alias "InternetGetLastResponseInfoA" (ByVal lpdwError As Long, ByVal lpszBuffer As String, ByVal lpdwBufferLength As Long) As Boolean
    'Private Declare Function FtpFindFirstFile Lib "wininet.dll" Alias "FtpFindFirstFileA" (ByVal hFtpSession As Long, ByVal lpszSearchFile As String, ByVal lpFindFileData As WIN32_FIND_DATA, ByVal dwFlags As Long, ByVal dwContent As Long) As Long
    'Private Declare Function InternetFindNextFile Lib "wininet.dll" Alias "InternetFindNextFileA" (ByVal hFind As Long, ByVal lpvFindData As WIN32_FIND_DATA) As Long

    Const FTP_TRANSFER_TYPE_UNKNOWN = &H0
    Const FTP_TRANSFER_TYPE_ASCII = &H1
    Const FTP_TRANSFER_TYPE_BINARY = &H2
    Const INTERNET_DEFAULT_FTP_PORT = 21 ' default for FTP servers
    Const INTERNET_SERVICE_FTP = 1
    Const INTERNET_FLAG_PASSIVE = &H8000000 ' used for FTP connections
    Const INTERNET_OPEN_TYPE_PRECONFIG = 0 ' use registry configuration
    Const INTERNET_OPEN_TYPE_DIRECT = 1 ' direct to net
    Const INTERNET_OPEN_TYPE_PROXY = 3 ' via named proxy
    Const INTERNET_OPEN_TYPE_PRECONFIG_WITH_NO_AUTOPROXY = 4 ' prevent using java/script/INS
    Const MAX_PATH = 260

    '*********************FTP END
    'Declare the shell object
    Private shObj As Object = Activator.CreateInstance(Type.GetTypeFromProgID("Shell.Application"))


    Private Shared Function customCertValidation(ByVal sender As Object,
                                                ByVal cert As X509Certificate,
                                                ByVal chain As X509Chain,
                                                ByVal errors As SslPolicyErrors) As Boolean

        Return True

    End Function

    Public Sub New()
        MyBase.New()
        blnBOD = True

        'Check if the the Event Log Exists 
        'If Not MyLog.SourceExists("OMS_Dongle") Then
        'MyLog.CreateEventSource("OMS_Dongle", "Client Log")
        ' Create Log 
        'End If
        'MyLog.Source = "OMS_Dongle"
        'MyLog.WriteEntry("It is running", "Client Log", EventLogEntryType.Information)
        InitializeComponents()
        ' TODO: Add any further initialization code
    End Sub

    Private Sub InitializeComponents()
        Me.ServiceName = "OMS_Dongle"
        Me.AutoLog = True
        Me.CanStop = True
        Me.Timer1 = New System.Timers.Timer()
        Me.Timer1.Interval = 20000

        Me.Timer1.Enabled = True
    End Sub

    ' This method starts the service.
    <MTAThread()> Shared Sub Main()
        ' To run more than one service you have to add them to the array
        System.ServiceProcess.ServiceBase.Run(New System.ServiceProcess.ServiceBase() {New OMS_Dongle})
    End Sub
    ' Clean up any resources being used.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        MyBase.Dispose(disposing)
        ' TODO: Add cleanup code here (if required)
    End Sub

    Protected Overrides Sub OnStart(ByVal args() As String)
        ' TODO: Add start code here (if required)
        ' to start your service.
        Me.Timer1.Enabled = True
    End Sub
    'Protected Overrides Sub OnPause()
    '    ' TODO: Add tear-down code here (if required) 
    '    ' to stop your service.
    '    Me.Timer1.Enabled = False
    '    'Thread.Sleep(40000)
    'End Sub

    Protected Overrides Sub OnStop()
        ' TODO: Add tear-down code here (if required) 
        ' to stop your service.

        Me.Timer1.Enabled = False
        'Dim fileLoc As String = My.Application.Info.DirectoryPath & "\Client_Stop_On_" & Format(Now(), "dd/MM/yyyy hh:mm:ss") & ".txt"
        'Dim fs As FileStream = Nothing
        'fs = File.Create(fileLoc)
        'fs.Close()
        'If File.Exists(fileLoc) Then
        '    Using sw As StreamWriter = New StreamWriter(fileLoc, True)
        '        sw.Write("Client Service Stop")
        '    End Using
        'End If
        'Thread.Sleep(40000)
    End Sub

    Private Sub InitializeComponent()
        Me.Timer1 = New System.Timers.Timer
        CType(Me.Timer1, System.ComponentModel.ISupportInitialize).BeginInit()
        '
        'Timer1
        '
        Me.Timer1.Enabled = True
        Me.Timer1.Interval = 20000
        '
        'Client
        '
        Me.CanPauseAndContinue = True
        Me.CanShutdown = True
        CType(Me.Timer1, System.ComponentModel.ISupportInitialize).EndInit()

    End Sub

    Private Sub Timer1_Elapsed(ByVal sender As System.Object, ByVal e As System.Timers.ElapsedEventArgs) Handles Timer1.Elapsed

        Timer1.Enabled = False
        Refresh_Server_Data()
        Timer1.Enabled = True
    End Sub

    Private Function ReadINI() As Boolean

        Try
            Dim dl As Long
            Dim strRetValue As StringBuilder
            Dim strFile As String
            strFile = Environ$("AllUsersProfile") & "\OMS_Soft\OMS_Soft.ini"

            strRetValue = New StringBuilder(255)
            dl = GetPrivateProfileString("AppData", "S1", "", strRetValue, 255, strFile)
            If dl <> 0 Then
                gstrServer_IP_Address = Decrypt_User_Password(strRetValue.ToString) '.Substring(1, dl)
            Else
                ReadINI = False
                Exit Function
            End If
            dl = GetPrivateProfileString("AppData", "S2", "", strRetValue, 255, strFile)
            If dl <> 0 Then
                gstrServer_Name = Left(Decrypt_User_Password(strRetValue.ToString), 15) '.Substring(1, dl)
            Else
                ReadINI = False
                Exit Function
            End If
            dl = GetPrivateProfileString("AppData", "S3", "", strRetValue, 255, strFile)
            If dl <> 0 Then
                gstrSQL_Server_Instance_Name = Decrypt_User_Password(strRetValue.ToString) '.Substring(1, dl)
            Else
                ReadINI = False
                Exit Function
            End If
            dl = GetPrivateProfileString("AppData", "S4", "", strRetValue, 255, strFile)
            If dl <> 0 Then
                gstrSQL_Server_Port = "," & strRetValue.ToString '.Substring(1, dl)
            Else
                gstrSQL_Server_Port = ""
            End If
            gstrSQL_Instance_User_Name = "BO"
            If Left(gstrSQL_Server_Instance_Name, 15) = gstrServer_Name Then
                gstrSQL_Server_Instance_Name = gstrServer_Name
            Else
                gstrSQL_Server_Instance_Name = gstrServer_Name & "\" & gstrSQL_Server_Instance_Name
            End If

            ReadINI = True
        Catch ex1 As Exception
            Print_Error_Only("ReadINI", ex1)
            ReadINI = False
        End Try
    End Function

    Private Function Create_Folder_Or_Delete_Old_Files(ByVal strConsume_Folder As String) As Boolean
        Try
            Dim di As DirectoryInfo
            Dim fiArr As FileInfo()
            Dim objFile As FileInfo
            If Directory.Exists(strConsume_Folder) Then
                di = New DirectoryInfo(strConsume_Folder)
                fiArr = di.GetFiles()
                For Each objFile In fiArr
                    If Math.Abs(DateDiff(DateInterval.Day, Now(), objFile.CreationTime)) > 7 Then
                        objFile.Delete()
                    End If
                Next
            Else
                Directory.CreateDirectory(strConsume_Folder)
            End If
        Catch ex1 As Exception
            Print_Error_Only("Create Folder Or Delete Old Files", ex1)
        End Try
    End Function

    Private Function Import_CSV_File(ByRef adoCon_Stock As SqlConnection, ByVal strSource_File As String, ByVal strTABLE As String) As Boolean
        Dim strLine As String
        Dim blnFirst_Line As Boolean
        Dim lngCounter As Long
        Dim strFile_Name As String
        Dim strheader() As String
        Dim intCounter As Integer
        Dim intCol_Count As Integer
        strFile_Name = Retrive_File_Name(strSource_File)
        Try
            blnFirst_Line = False
            lngCounter = 1
            intCol_Count = 0
            objStreamReader = New StreamReader(strSource_File)
            Do
                strLine = objStreamReader.ReadLine
                If InStr(1, strLine, ",", vbTextCompare) > 0 Then
                    If blnFirst_Line = False Then
                        blnFirst_Line = True
                        strheader = Split(strLine, ",")
                        strSQL_String = "BEGIN"
                        strSQL_String = strSQL_String & vbCrLf & "If EXISTS(SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[" & strTABLE & "]') AND type in (N'U'))"
                        strSQL_String = strSQL_String & vbCrLf & "DROP TABLE [dbo].[" & strTABLE & "]"
                        strSQL_String = strSQL_String & vbCrLf & "CREATE TABLE [dbo].[" & strTABLE & "](Production_Date DateTime NULL"
                        For intCounter = 0 To UBound(strheader)
                            If Trim(UCase(strheader(intCounter))) = "" Then
                                strSQL_String = strSQL_String & vbCrLf & ", Blank_Col varchar(1000) null"
                                ''ElseIf Trim(UCase(strheader(intCounter))) = "LAST_LOGIN_DATE" And UCase(strTABLE) = UCase("User_Last_Login") Then
                                ''strSQL_String = strSQL_String & vbCrLf & ", [" & Trim(UCase(strheader(intCounter))) & "] DATETIME null"
                            Else
                                strSQL_String = strSQL_String & vbCrLf & ", [" & Trim(UCase(strheader(intCounter))) & "] varchar(1000) null"
                            End If
                        Next
                        strSQL_String = strSQL_String & vbCrLf & ")"
                        strSQL_String = strSQL_String & vbCrLf & "END"

                        command = New SqlCommand(strSQL_String, adoCon_Stock)
                        command.CommandTimeout = 0
                        command.ExecuteNonQuery()

                    ElseIf lngCounter = 2 Or lngCounter = 500 Then
                        If lngCounter = 500 Then
                            command = New SqlCommand(strSQL_String, adoCon_Stock)
                            command.CommandTimeout = 0
                            command.ExecuteNonQuery()

                            lngCounter = 2
                        End If

                        strSQL_String = "INSERT INTO [dbo].[" & strTABLE & "]"
                        strSQL_String = strSQL_String & vbCrLf & "VALUES(GETDATE(),RTRIM('" & Replace(Replace(strLine, "'", "''"), ",", "'),RTRIM('") & "'))"
                    Else
                        strSQL_String = strSQL_String & vbCrLf & ",(GETDATE(),RTRIM('" & Replace(Replace(strLine, "'", "''"), ",", "'),RTRIM('") & "'))"
                    End If
                    lngCounter = lngCounter + 1
                End If
            Loop Until strLine Is Nothing
            If strSQL_String <> "" Then
                command = New SqlCommand(strSQL_String, adoCon_Stock)
                command.CommandTimeout = 0
                command.ExecuteNonQuery()
            End If
            objStreamReader.Close()
            Import_CSV_File = True
        Catch ex4 As Exception
            Print_Error_Only("Import CSV File ", ex4)
            Import_CSV_File = False
            Exit Function
        End Try
    End Function

    'Private Function Upload_RSInfo(ByVal strSetup_Path As String, ByVal strCompany_Id As String) As Boolean
    '    Dim strConsume_Folder As String
    '    Upload_RSInfo = False
    '    Try
    '        Dim adapter As New SqlDataAdapter
    '        Dim adoRs_Stock As DataSet
    '        Dim strFile As String
    '        Dim strHeading As String

    '        command = New SqlCommand("SELECT TOP 1 '" & strCompany_Id & "' AS Company_Id, CONVERT(VARCHAR,Exe_Date,103) + ' ' + CONVERT(VARCHAR,Exe_Date,108) AS Exe_Date, REPLACE(Exe_Size,',','') AS Exe_Size, '" & strExe_Version & "' Exe_Version, IP_Address, Server_Name, '" & Replace(Replace(gstrLicense_Information, "'", ""), ",", "") & "' AS Info,'" & strCompany_Id & " - " & RTrim(Replace(Replace(gstrCompany_Name, "'", ""), ",", "")) & "','" & RTrim(strHO_Desc) & "','" & RTrim(strBranch_Desc) & "', REPLACE(Program_Name,'RETAIL_SOFT_','Server Info(') + ') On Date : ' + CONVERT(VARCHAR,GETDATE(),103) + ' ' + CONVERT(VARCHAR,GETDATE(),108),'" & strOpening_Date & "' FROM [RetailSoft_Company].DBO.tblServer", adoCon_Company)
    '        command.CommandTimeout = 0

    '        adapter = New SqlDataAdapter
    '        adapter.SelectCommand = command
    '        adoRs_Stock = New DataSet
    '        adapter.Fill(adoRs_Stock)
    '        adapter.Dispose()
    '        adapter = Nothing
    '        command.Dispose()

    '        If adoRs_Stock.Tables(0).Rows.Count > 0 Then
    '            strConsume_Folder = My.Application.Info.DirectoryPath & "\Log_Files"
    '            Create_Folder_Or_Delete_Old_Files(strConsume_Folder)

    '            strFile = strConsume_Folder & "\Restore_RSInfo_" & Format(Now(), "ddMMyyyy-HHmmss") & ".csv"
    '            strHeading = "Company_Id,Exe_Date,Exe_Size,Exe_Version,IP_Address,Server_Name,Info,Company,HO,Branch,Server_info,Go_Live"

    '            Adodb_To_CSV(adoRs_Stock, strFile, strHeading)

    '            If Put_FTP_Files("59.90.32.112", "FTP_User1", "Timken123#", True, strSetup_Path, strFile, False) = True Then
    '                Upload_RSInfo = True
    '            End If
    '            File.Delete(strFile)
    '        End If
    '    Catch ex1 As Exception
    '        Print_Error_Only("Send Counter File To CS", ex1)
    '    End Try
    'End Function

    Private Function Update_RSLog(ByVal strCompany_Id As String) As Boolean
        Dim strConsume_Folder As String
        Try
            Dim adapter As New SqlDataAdapter
            Dim adoRs_Stock As DataSet
            Dim strFile As String
            Dim strHeading As String

            command = New SqlCommand("SELECT Monthly_Upload_Date FROM [RetailSoft_Company].[dbo].[tblRSInfo_Log] WHERE Company_Id = '" & strCompany_Id & "' AND MONTH(Monthly_Upload_Date) = MONTH(GETDATE()) AND Download_Upload = 3", adoCon_Company)
            command.CommandTimeout = 0

            adapter = New SqlDataAdapter
            adapter.SelectCommand = command
            adoRs_Stock = New DataSet
            adapter.Fill(adoRs_Stock)
            adapter.Dispose()
            adapter = Nothing
            command.Dispose()

            If adoRs_Stock.Tables(0).Rows.Count = 0 Then
                strSQL_String = "BEGIN"
                strSQL_String = strSQL_String & vbCrLf & "  IF NOT EXISTS(SELECT * FROM [RetailSoft_Company].DBO.tblRSInfo_Log WHERE Company_Id = '" & gstrCompany_Id & "' AND Download_Upload = 3)"
                strSQL_String = strSQL_String & vbCrLf & "      INSERT INTO [RetailSoft_Company].DBO.tblRSInfo_Log(Company_Id, Entry_Date, Update_Values,Download_Upload,Monthly_Upload_Date) VALUES('" & gstrCompany_Id & "', GETDATE(),'',3, GETDATE())"
                strSQL_String = strSQL_String & vbCrLf & "  ELSE"
                strSQL_String = strSQL_String & vbCrLf & "      UPDATE [RetailSoft_Company].DBO.tblRSInfo_Log SET Company_Id = '" & gstrCompany_Id & "', Entry_Date = GETDATE(), Monthly_Upload_Date = GETDATE() WHERE Company_Id = '" & gstrCompany_Id & "' AND Download_Upload = 3"
                strSQL_String = strSQL_String & vbCrLf & "END"

                command = New SqlCommand(strSQL_String, adoCon_Company)
                command.CommandTimeout = 0
                command.ExecuteNonQuery()
                Update_RSLog = True
            Else
                Update_RSLog = False
            End If

        Catch ex1 As Exception
            Print_Error_Only("Update_RSLog", ex1)
            Update_RSLog = False
        End Try
    End Function

    Private Function Upload_RSInfo(ByVal strSetup_Path As String, ByVal strCompany_Id As String) As Boolean
        Dim strConsume_Folder As String
        Try
            Dim adapter As New SqlDataAdapter
            Dim adoRs_Stock As DataSet
            Dim strFile As String
            Dim strHeading As String

            command = New SqlCommand("SELECT Company_Id,CONVERT(VARCHAR,Exe_Date,103) + ' ' + CONVERT(VARCHAR,Exe_Date,108) AS Exe_Date,Exe_Size,Exe_Version,IP_Address,Server_Name,Info,Company,HO,Branch,Server_info,Go_Live FROM [RetailSoft_Company].[dbo].[tblRsInfo_Up] WHERE Company_Id = '" & strCompany_Id & "'", adoCon_Company)
            command.CommandTimeout = 0

            adapter = New SqlDataAdapter
            adapter.SelectCommand = command
            adoRs_Stock = New DataSet
            adapter.Fill(adoRs_Stock)
            adapter.Dispose()
            adapter = Nothing
            command.Dispose()

            If adoRs_Stock.Tables(0).Rows.Count > 0 Then
                strConsume_Folder = My.Application.Info.DirectoryPath & "\Log_Files"
                Create_Folder_Or_Delete_Old_Files(strConsume_Folder)

                strFile = strConsume_Folder & "\Restore_RSInfo_" & Format(Now(), "ddMMyyyy-HHmmss") & ".csv"
                strHeading = "Company_Id,Exe_Date,Exe_Size,Exe_Version,IP_Address,Server_Name,Info,Company,HO,Branch,Server_info,Go_Live"

                Adodb_To_CSV(adoRs_Stock, strFile, strHeading)

                If Put_FTP_Files("59.90.32.112", "FTP_User1", "Timken123#", True, strSetup_Path, strFile, False) = True Then
                    command = New SqlCommand("DELETE FROM [RetailSoft_Company].DBO.tblRsInfo_Up WHERE Company_Id = '" & strCompany_Id & "'", adoCon_Company)
                    command.CommandTimeout = 0
                    command.ExecuteNonQuery()
                End If
                File.Delete(strFile)
            End If
        Catch ex1 As Exception
            Print_Error_Only("Upload_RSInfo", ex1)
        End Try
    End Function

    Public Function Read_NetRed_Hasp_Lock(ByRef strString_Bytes As String)
        Service = NET_READ_MEMO_BLOCK
        MemoHaspBuffer.txt = ""      ' Clear the buffer before reading
        p1 = 104
        p2 = 24
        p3 = 0
        p4& = p2
        Call hasp(Service, SeedCode, ProgramNum, Passw1, Passw2, p1, p2, p3, p4&)
        If p3 <> 0 Then
            strString_Bytes = ""
            Exit Function
        End If
        Call ReadHaspBlock(Service, MemoHaspBuffer, p2)
        gstrCompany_Block1 = Left(MemoHaspBuffer.txt, 48)

        MemoHaspBuffer.txt = ""      ' Clear the buffer before reading
        p1 = 128
        p2 = 24
        p3 = 0
        p4& = p2
        Call hasp(Service, SeedCode, ProgramNum, Passw1, Passw2, p1, p2, p3, p4&)
        If p3 <> 0 Then
            strString_Bytes = ""
            Exit Function
        End If
        Call ReadHaspBlock(Service, MemoHaspBuffer, p2)
        gstrCompany_Block2 = Left(MemoHaspBuffer.txt, 48)

        MemoHaspBuffer.txt = ""      ' Clear the buffer before reading

        p1 = 152
        p2 = 24
        p3 = 0
        p4& = p2
        Call hasp(Service, SeedCode, ProgramNum, Passw1, Passw2, p1, p2, p3, p4&)
        If p3 <> 0 Then
            strString_Bytes = ""
            Exit Function
        End If
        Call ReadHaspBlock(Service, MemoHaspBuffer, p2)
        gstrCompany_Block3 = Left(MemoHaspBuffer.txt, 48)

        MemoHaspBuffer.txt = ""      ' Clear the buffer before reading

        p1 = 176
        p2 = 24
        p3 = 0
        p4& = p2
        Call hasp(Service, SeedCode, ProgramNum, Passw1, Passw2, p1, p2, p3, p4&)
        If p3 <> 0 Then
            strString_Bytes = ""
            Exit Function
        End If
        Call ReadHaspBlock(Service, MemoHaspBuffer, p2)
        gstrCompany_Desc_Block = Left(MemoHaspBuffer.txt, 48)

        p1 = 200
        p2 = 24
        p3 = 0
        p4& = p2
        Call hasp(Service, SeedCode, ProgramNum, Passw1, Passw2, p1, p2, p3, p4&)
        If p3 <> 0 Then
            strString_Bytes = ""
            Exit Function
        End If
        Call ReadHaspBlock(Service, MemoHaspBuffer, p2)
        strString_Bytes = Left(MemoHaspBuffer.txt, 48)
        MemoHaspBuffer.txt = "" ' Clear the buffer before reading

        p1 = 200 + 24
        p2 = 24
        p3 = 0
        p4& = p2
        Call hasp(Service, SeedCode, ProgramNum, Passw1, Passw2, p1, p2, p3, p4&)
        Call ReadHaspBlock(Service, MemoHaspBuffer, p2)
        'MsgBox Left(MemoHaspBuffer.txt, 48)
        strString_Bytes = strString_Bytes & Left(MemoHaspBuffer.txt, 48)

    End Function

    Public Function Search_Lock() As Boolean
        Dim strString_Bytes As String

        strString_Bytes = ""
        Service = IS_HASP
        SeedCode = 100
        Passw1 = 17367&
        Passw2 = 20767&
        Search_Lock = False
        Call hasp(Service, SeedCode, 0&, Passw1, Passw2, p1, p2, p3, p4&)
        If p1 = 1 Then
            gobjLock_Type = Is_MemoHASP()

            If gobjLock_Type = enumLock_Type.LOCK_LOCAL_RED_HASP Or gobjLock_Type = enumLock_Type.LOCK_LOCAL_USB_RED_HASP Then   'NET HASP LOCK FOUND ON LOCAL MACHINE
                MemoHaspBuffer.txt = ""      'Clear the buffer before reading
                Service = 50 'READ_MEMO_BLOCK
                p1 = 104
                p2 = 144
                p3 = 0
                p4& = 0
                Call hasp(Service, SeedCode, LptNum, Passw1, Passw2, p1, p2, p3, p4&)
                If (p3 = 0) Then
                    Call ReadHaspBlock(Service, MemoHaspBuffer, p2)

                    strString_Bytes = CStr(MemoHaspBuffer.txt)

                    gstrCompany_Block1 = Left(strString_Bytes, 48)
                    gstrCompany_Block2 = Mid(strString_Bytes, 49, 48)
                    gstrCompany_Block3 = Mid(strString_Bytes, 97, 48)

                    gstrCompany_Desc_Block = Mid(strString_Bytes, 145, 48)
                    strString_Bytes = Mid(strString_Bytes, 193, 96)

                    gstrLicense_Information = strString_Bytes & gstrCompany_Block1 & gstrCompany_Block2 & gstrCompany_Block3 & gstrCompany_Desc_Block
                    gstrLock_Company_Id = Trim(Mid(gstrLicense_Information, 92, 4))
                    'hasp id
                    Call hasp(6, SeedCode, LptNum, Passw1, Passw2, p1, p2, p3, p4&)

                    If p3 = 0 Then  'Success
                        gstrHasp_LockId = CStr(p1 + (65536 * p2))

                    End If
                    Search_Lock = True
                    Exit Function
                End If
            End If
        End If
        Service = NET_LOGIN
        Call hasp(Service, SeedCode, ProgramNum, Passw1, Passw2, p1, p2, p3, p4&)
        Call hasp(NET_LAST_STATUS, SeedCode, ProgramNum, Passw1, Passw2, RetStatus&, dummy&, dummy&, dummy&)
        If (RetStatus = 0) Then 'NET HASP LOCK FOUND ON REMOTE MACHINE
            Read_NetRed_Hasp_Lock(strString_Bytes)
            gobjLock_Type = enumLock_Type.LOCK_REMOTE_RED_HASP

            gstrLicense_Information = strString_Bytes & gstrCompany_Block1 & gstrCompany_Block2 & gstrCompany_Block3 & gstrCompany_Desc_Block
            gstrLock_Company_Id = Trim(Mid(gstrLicense_Information, 92, 4))
            'hasp id
            Call hasp(46, SeedCode, ProgramNum, Passw1, Passw2, p1, p2, p3, p4&)
            If p3 = 0 Then  'Success
                gstrHasp_LockId = CStr(p1 + (65536 * p2))
            End If
            Search_Lock = True
        End If
    End Function

    Private Function Download_RSInfo(ByVal strSetup_Path As String, ByVal strCompany_Id As String) As Boolean
        '''TS-1838 | Retail-Soft service should check for the specific key available on FTP and consume it once available for the store.
        Try
            Dim strConsume_Folder As String
            Dim strCurrentFile As String
            Dim strLine As String
            Dim intFile As Integer
            Dim blnSuccess As Boolean

            intFile = 0
            Dirlist = New List(Of String) 'I prefer List() instead of an array
            FTP_Folder_Files_List("59.90.32.112", "FTP_User1", "Timken123#", True, strSetup_Path, Dirlist)
            If Dirlist.Count > 0 Then
                If intFile = 0 Then
                    Search_Lock()

                    intFile = 1
                End If

                strConsume_Folder = My.Application.Info.DirectoryPath & "\Log_Files"
                Create_Folder_Or_Delete_Old_Files(strConsume_Folder)

                For intx = 0 To (Dirlist.Count - 1)
                    strCurrentFile = Dirlist.Item(intx)
                    strCurrentFile = Mid(strCurrentFile, InStr(strCurrentFile, "/", CompareMethod.Text) + 1, Len(strCurrentFile))
                    If InStr(UCase(strCurrentFile), "RSINFO_" & Format(Now(), "ddMMyyyy"), CompareMethod.Text) > 0 Then
                        If File.Exists(strConsume_Folder & "/" & strCurrentFile) Then
                            File.Delete(strConsume_Folder & "/" & strCurrentFile)
                        End If
                        If Get_FTP_File("59.90.32.112", "FTP_User1", "Timken123#", True, strSetup_Path & "/" & strCurrentFile, strConsume_Folder & "/" & strCurrentFile) = True Then
                            objStreamReader = New StreamReader(strConsume_Folder & "/" & strCurrentFile)
                            Do
                                strLine = objStreamReader.ReadLine
                                Exit Do
                            Loop Until strLine Is Nothing
                            objStreamReader.Close()
                            If File.Exists(strConsume_Folder & "/" & strCurrentFile) Then
                                File.Delete(strConsume_Folder & "/" & strCurrentFile)
                            End If
                            gstrSMS_Message = ""
                            If Update_Lock_From_FTP(strLine) = True Then
                                If Drop_FTP_File("59.90.32.112", "FTP_User1", "Timken123#", True, strSetup_Path & "/" & strCurrentFile) = True Then
                                    blnSuccess = True
                                    If gstrSMS_Message <> "" And gstrHasp_LockId <> "1917058163" Then
                                        Send_SMS_SP("9922964296", gstrSMS_Message)
                                        'Send_SMS_SP("9326264245", gstrSMS_Message)
                                    End If
                                End If
                            End If
                        End If
                    Else
                        Drop_FTP_File("59.90.32.112", "FTP_User1", "Timken123#", True, strSetup_Path & "/" & strCurrentFile)
                    End If
                Next
            End If
            Download_RSInfo = blnSuccess
        Catch ex1 As Exception
            Print_Error_Only("Download_RSInfo", ex1)
            Download_RSInfo = False
        End Try
    End Function

    Public Function ConvertStr(ByVal strNum As String) As String
        Dim lngCounter As Long
        Dim strReturn As String

        For lngCounter = 1 To Len(strNum)
            If Asc(Mid(strNum, lngCounter, 1)) = 82 Then
                strReturn = strReturn & Chr(48)
            ElseIf Asc(Mid(strNum, lngCounter, 1)) = 86 Then
                strReturn = strReturn & Chr(49)
            ElseIf Asc(Mid(strNum, lngCounter, 1)) = 73 Then
                strReturn = strReturn & Chr(50)
            ElseIf Asc(Mid(strNum, lngCounter, 1)) = 74 Then
                strReturn = strReturn & Chr(51)
            ElseIf Asc(Mid(strNum, lngCounter, 1)) = 65 Then
                strReturn = strReturn & Chr(52)
            ElseIf Asc(Mid(strNum, lngCounter, 1)) = 89 Then
                strReturn = strReturn & Chr(53)
            ElseIf Asc(Mid(strNum, lngCounter, 1)) = 83 Then
                strReturn = strReturn & Chr(54)
            ElseIf Asc(Mid(strNum, lngCounter, 1)) = 85 Then
                strReturn = strReturn & Chr(55)
            ElseIf Asc(Mid(strNum, lngCounter, 1)) = 80 Then
                strReturn = strReturn & Chr(56)
            ElseIf Asc(Mid(strNum, lngCounter, 1)) = 69 Then
                strReturn = strReturn & Chr(57)
            End If
        Next
        ConvertStr = strReturn
    End Function

    Public Function ConvertYear(ByVal strYear As String) As String
        If strYear = "M" Then
            ConvertYear = 2012
        ElseIf strYear = "N" Then
            ConvertYear = 2013
        ElseIf strYear = "O" Then
            ConvertYear = 2014
        ElseIf strYear = "P" Then
            ConvertYear = 2015
        ElseIf strYear = "Q" Then
            ConvertYear = 2016
        ElseIf strYear = "R" Then
            ConvertYear = 2017
        ElseIf strYear = "S" Then
            ConvertYear = 2018
        ElseIf strYear = "T" Then
            ConvertYear = 2019
        ElseIf strYear = "U" Then
            ConvertYear = 2020
        ElseIf strYear = "V" Then
            ConvertYear = 2021
        ElseIf strYear = "W" Then
            ConvertYear = 2022
        ElseIf strYear = "X" Then
            ConvertYear = 2023
        ElseIf strYear = "Y" Then
            ConvertYear = 2024
        ElseIf strYear = "Z" Then
            ConvertYear = 2025
        ElseIf strYear = "A" Then
            ConvertYear = 2026
        ElseIf strYear = "B" Then
            ConvertYear = 2027
        ElseIf strYear = "C" Then
            ConvertYear = 2028
        ElseIf strYear = "D" Then
            ConvertYear = 2029
        ElseIf strYear = "E" Then
            ConvertYear = 2030
        ElseIf strYear = "F" Then
            ConvertYear = 2031
        ElseIf strYear = "G" Then
            ConvertYear = 2032
        ElseIf strYear = "H" Then
            ConvertYear = 2033
        ElseIf strYear = "I" Then
            ConvertYear = 2034
        ElseIf strYear = "J" Then
            ConvertYear = 2035
        ElseIf strYear = "K" Then
            ConvertYear = 2036
        ElseIf strYear = "L" Then
            ConvertYear = 2037
        ElseIf strYear = "0" Then
            ConvertYear = 2038
        ElseIf strYear = "1" Then
            ConvertYear = 2039
        ElseIf strYear = "2" Then
            ConvertYear = 2040
        ElseIf strYear = "3" Then
            ConvertYear = 2041
        ElseIf strYear = "4" Then
            ConvertYear = 2042
        ElseIf strYear = "5" Then
            ConvertYear = 2043
        ElseIf strYear = "6" Then
            ConvertYear = 2044
        ElseIf strYear = "7" Then
            ConvertYear = 2045
        ElseIf strYear = "8" Then
            ConvertYear = 2046
        ElseIf strYear = "9" Then
            ConvertYear = 2047
        End If
    End Function

    Public Function ChangeLockInfo(ByVal intCompany_No As Integer, ByVal gobjTempLockFlag As enumLock_Type, ByVal strAction As String, ByVal strNew_Value As String) As Boolean
        Dim strLock_Information As String
        Dim strVersion_String As String
        Dim strSerial_Key As String
        Dim strReturnedString As String
        Dim strCompany_Desc As String
        Dim lngFree_File As Long
        Dim lngErr_File As Long
        Dim strFile_Name As String

        ChangeLockInfo = True
        If gobjTempLockFlag = enumLock_Type.LOCK_LOCAL_RED_HASP Or gobjTempLockFlag = enumLock_Type.LOCK_LOCAL_USB_RED_HASP Then
            If strAction = "A" Then 'HERE "A" DENOTE FOR COMPANY NAME
                MemoHaspBuffer.txt = ""      'Clear the buffer before reading
                Service = READ_MEMO_BLOCK
                p1 = 176
                p2 = 24
                p3 = 0
                p4& = 0
                Call hasp(Service, SeedCode, LptNum, Passw1, Passw2, p1, p2, p3, p4&)
                If (p3 = 0) Then
                    Call ReadHaspBlock(Service, MemoHaspBuffer, p2)
                    strLock_Information = Trim(MemoHaspBuffer.txt)

                    Service = WRITE_MEMO_BLOCK
                    MemoHaspBuffer.txt = Left(strNew_Value, 48)
                    p1 = 176       ' Block size in words!
                    p2 = 24       ' Block size in words!
                    Call WriteHaspBlock(Service, MemoHaspBuffer, p2)
                    Call hasp(Service, SeedCode, LptNum, Passw1, Passw2, p1, p2, p3, p4&)
                    If p3 <> 0 Then
                        ChangeLockInfo = False
                    End If
                End If
            Else
                If intCompany_No = 0 Then
                    MemoHaspBuffer.txt = ""      ' Clear the buffer before reading
                    Service = READ_MEMO_BLOCK
                    p1 = 200
                    p2 = 48
                    p3 = 0
                    p4& = 0
                    Call hasp(Service, SeedCode, LptNum, Passw1, Passw2, p1, p2, p3, p4&)
                    If (p3 = 0) Then
                        Call ReadHaspBlock(Service, MemoHaspBuffer, p2)
                        strLock_Information = MemoHaspBuffer.txt
                    End If
                End If
                Update_Lock_String(strLock_Information, strAction, strNew_Value, intCompany_No)

                Service = WRITE_MEMO_BLOCK
                MemoHaspBuffer.txt = strLock_Information
                If intCompany_No = 0 Then
                    p1 = 200       ' Block size in words!
                    p2 = 48       ' Block size in words!
                ElseIf intCompany_No = 1 Then
                    p1 = 104       ' Block size in words!
                    p2 = 32       ' Block size in words!
                ElseIf intCompany_No = 2 Then
                    p1 = 136       ' Block size in words!
                    p2 = 32       ' Block size in words!
                ElseIf intCompany_No = 3 Then
                    p1 = 168       ' Block size in words!
                    p2 = 32       ' Block size in words!
                End If
                If strAction = "$" Then
                    gstrLicense_Information = strLock_Information
                End If

                Call WriteHaspBlock(Service, MemoHaspBuffer, p2)
                Call hasp(Service, SeedCode, LptNum, Passw1, Passw2, p1, p2, p3, p4&)
                If strAction = "$" Then
                    gstrLicense_Information = strLock_Information
                End If

                If p3 <> 0 Then
                    ChangeLockInfo = False
                    'MsgBox "Write Information failed. Error code = " & CStr(p3), vbInformation, "Inst"
                End If
            End If
        ElseIf gobjTempLockFlag = enumLock_Type.LOCK_REMOTE_RED_HASP Then
            Service = NET_LOGIN
            Call hasp(Service, SeedCode, ProgramNum, Passw1, Passw2, p1, p2, p3, p4&)
            Call hasp(NET_LAST_STATUS, SeedCode, ProgramNum, Passw1, Passw2, RetStatus&, dummy&, dummy&, dummy&)
            If (RetStatus = 0) Then 'NET HASP LOCK FOUND ON REMOTE MACHINE
                If strAction = "A" Then 'HERE "A" DENOTE FOR COMPANY NAME
                    MemoHaspBuffer.txt = ""      ' Clear the buffer before reading
                    Service = NET_READ_MEMO_BLOCK
                    p1 = 176
                    p2 = 24
                    p3 = 0
                    p4& = p2
                    Call hasp(Service, SeedCode, ProgramNum, Passw1, Passw2, p1, p2, p3, p4&)

                    Service = NET_WRITE_MEMO_BLOCK
                    MemoHaspBuffer.txt = Left(strNew_Value, 48)
                    p1 = 176       ' Block size in words!
                    p2 = 24       ' Block size in words!
                    p3 = 0
                    p4& = p2
                    Call WriteHaspBlock(Service, MemoHaspBuffer, p2)
                    Call hasp(Service, SeedCode, ProgramNum, Passw1, Passw2, p1, p2, p3, p4&)
                    Call hasp(NET_LAST_STATUS, SeedCode, ProgramNum, Passw1, Passw2, RetStatus&, dummy&, dummy&, dummy&)
                Else
                    If intCompany_No = 0 Then
                        Read_NetRed_Hasp_Lock(strLock_Information)
                    End If
                    Update_Lock_String(strLock_Information, strAction, strNew_Value, intCompany_No)

                    Service = NET_WRITE_MEMO_BLOCK

                    MemoHaspBuffer.txt = Left(strLock_Information, 48)
                    If intCompany_No = 0 Then
                        p1 = 200       ' Block size in words!
                    ElseIf intCompany_No = 1 Then
                        p1 = 104       ' Block size in words!
                    ElseIf intCompany_No = 2 Then
                        p1 = 136       ' Block size in words!
                    ElseIf intCompany_No = 3 Then
                        p1 = 168       ' Block size in words!
                    End If
                    p2 = 24       ' Block size in words!
                    p3 = 0
                    p4& = p2
                    Call WriteHaspBlock(Service, MemoHaspBuffer, p2)
                    Call hasp(Service, SeedCode, ProgramNum, Passw1, Passw2, p1, p2, p3, p4&)
                    Call hasp(NET_LAST_STATUS, SeedCode, ProgramNum, Passw1, Passw2, RetStatus&, dummy&, dummy&, dummy&)

                    If intCompany_No = 0 Then
                        MemoHaspBuffer.txt = Mid(strLock_Information, 49, 48)
                        p1 = 200 + 24     ' Block size in words!
                        p2 = 24           ' Block size in words!
                    ElseIf intCompany_No = 1 Then
                        MemoHaspBuffer.txt = Mid(strLock_Information, 49, 16)
                        p1 = 104 + 24     ' Block size in words!
                        p2 = 8           ' Block size in words!
                    ElseIf intCompany_No = 2 Then
                        MemoHaspBuffer.txt = Mid(strLock_Information, 49, 16)
                        p1 = 136 + 24     ' Block size in words!
                        p2 = 8           ' Block size in words!
                    ElseIf intCompany_No = 3 Then
                        MemoHaspBuffer.txt = Mid(strLock_Information, 49, 16)
                        p1 = 168 + 24     ' Block size in words!
                        p2 = 8           ' Block size in words!
                    End If

                    p3 = 0
                    p4& = p2
                    Call WriteHaspBlock(Service, MemoHaspBuffer, p2)
                    Call hasp(Service, SeedCode, ProgramNum, Passw1, Passw2, p1, p2, p3, p4&)
                    Call hasp(NET_LAST_STATUS, SeedCode, ProgramNum, Passw1, Passw2, RetStatus&, dummy&, dummy&, dummy&)

                    If strAction = "$" Then
                        gstrLicense_Information = strLock_Information
                    End If
                End If
                If p3 <> 0 Then
                    ChangeLockInfo = False
                End If
            End If
        End If
        If ChangeLockInfo = True Then
            Lock_Log_File(intCompany_No, gobjTempLockFlag, strAction, strNew_Value)
        End If
    End Function

    Private Sub Lock_Log_File(ByVal intCompany_No As Integer, ByVal gobjTempLockFlag As enumLock_Type, ByVal strAction As String, ByVal strNew_Value As String)
        Try
            Dim adocommand As New SqlCommand("INSERT INTO [RetailSoft_Company].dbo.tblLock_Log(Entry_Date,Action,Company_Position,Lock_Type,New_Values,Updated_By_IP,Updated_By_Machine)VALUES(GETDATE(),'" & strAction & "'," & intCompany_No & "," & gobjTempLockFlag & ",'" & Replace(strNew_Value, "'", "''") & "','" & gstrServer_IP_Address & "','" & gstrServer_Name & "')", adoCon_Company)
            adocommand.CommandTimeout = 0
            adocommand.ExecuteNonQuery()
        Catch ex1 As Exception
            Print_Error_Only("Lock_Log_File", ex1)
        End Try
    End Sub

    Private Function Update_Lock_String(ByRef strLock_Information As String, ByVal strAction As String, ByVal strNew_Value As String, ByVal intCompany_No As Integer)
        If strAction = "$" Then 'HERE "$" SPECIALLY USE FOR ONCE TO UPDATE COMPANY ID & FLAG
            strLock_Information = Left(strLock_Information, 91) & strNew_Value
        ElseIf strAction = "B" Then 'HERE "B" DENOTE FOR DEMO
            strLock_Information = Left(strLock_Information, 16) & Trim(strNew_Value) & Mid(strLock_Information, 18, Len(strLock_Information))
        ElseIf strAction = "C" Then 'HERE "C" DENOTE FOR PRINT FOOTER
            strLock_Information = Left(strLock_Information, 20) & Trim(strNew_Value) & Mid(strLock_Information, 22, Len(strLock_Information))
        ElseIf strAction = "D" Then 'HERE "D" DENOTE FOR FROM DATE TO DATE
            strLock_Information = Trim(strNew_Value) & Mid(strLock_Information, 17, Len(strLock_Information))
        ElseIf strAction = "E" Or strAction = "#" Then 'HERE "E" DENOTE FOR VERSION
            If intCompany_No = 0 Then
                strLock_Information = Left(strLock_Information, 37) & Left(Trim(strNew_Value) & Repl_String(36, " "), 36) & Mid(strLock_Information, 74, Len(strLock_Information))
            Else
                strLock_Information = Left(gstrAsyncronous_String, 13) & Left(Trim(strNew_Value) & Repl_String(36, " "), 36) & Mid(gstrAsyncronous_String, 50, 14)
            End If
        ElseIf strAction = "F" Then 'HERE "F" DENOTE FOR NO OF COMPANY
            strLock_Information = Left(strLock_Information, 17) & Trim(strNew_Value) & Mid(strLock_Information, 21, Len(strLock_Information))
        ElseIf strAction = "G" Then 'HERE "G" DENOTE FOR HEAD OFFICE + LOCATION + BRANCH TYPE +
            strLock_Information = Left(strLock_Information, 21) & Trim(strNew_Value) & Mid(strLock_Information, 29, Len(strLock_Information))
        ElseIf strAction = "H" Then 'HERE "H" DENOTE FOR TOTAL USER
            If intCompany_No = 0 Then
                strLock_Information = Left(strLock_Information, 28) & Left(Trim(strNew_Value), 3) & Mid(strLock_Information, 32, Len(strLock_Information))
            Else
                strLock_Information = Left(gstrAsyncronous_String, 4) & Left(Trim(strNew_Value), 3) & Mid(gstrAsyncronous_String, 8, 57)
            End If
        ElseIf strAction = "I" Then 'HERE "I" DENOTE FOR BO USER
            If intCompany_No = 0 Then
                strLock_Information = Left(strLock_Information, 31) & Left(Trim(strNew_Value), 3) & Mid(strLock_Information, 35, Len(strLock_Information))
            Else
                strLock_Information = Left(gstrAsyncronous_String, 7) & Left(Trim(strNew_Value), 3) & Mid(gstrAsyncronous_String, 11, 54)
            End If
        ElseIf strAction = "J" Then 'HERE "J" DENOTE FOR PHY USER
            If intCompany_No = 0 Then
                strLock_Information = Left(strLock_Information, 34) & Left(Trim(strNew_Value), 3) & Mid(strLock_Information, 38, Len(strLock_Information))
            Else
                strLock_Information = Left(gstrAsyncronous_String, 10) & Left(Trim(strNew_Value), 3) & Mid(gstrAsyncronous_String, 14, 51)
            End If
        ElseIf strAction = "L" Then 'HERE "L" DENOTE FOR MOBILE POS USER
            If intCompany_No = 0 Then
                strLock_Information = Left(strLock_Information, 88) & Left(Trim(strNew_Value), 3) & Mid(strLock_Information, 92, Len(strLock_Information))
            Else
                strLock_Information = Left(gstrAsyncronous_String, 52) & Left(Trim(strNew_Value), 3) & Mid(gstrAsyncronous_String, 56, 9)
            End If
        ElseIf strAction = "M" Then 'HERE "M" DENOTE FOR OFFLINE POS USER
            If intCompany_No = 0 Then
                strLock_Information = Left(strLock_Information, 85) & Left(Trim(strNew_Value), 3) & Mid(strLock_Information, 89, Len(strLock_Information))
            Else
                strLock_Information = Left(gstrAsyncronous_String, 49) & Left(Trim(strNew_Value), 3) & Mid(gstrAsyncronous_String, 53, 12)
            End If
        ElseIf strAction = "N" Then 'HERE "N" DENOTE FOR SAP Setup Type
            If intCompany_No = 0 Then
                strLock_Information = Left(strLock_Information, 84) & Left(strNew_Value, 1) & Mid(strLock_Information, 86, Len(strLock_Information))
            Else
                strLock_Information = Left(gstrAsyncronous_String, 57) & Left(strNew_Value, 1) & Mid(gstrAsyncronous_String, 59, 6)
            End If
        ElseIf strAction = "O" Then 'HERE "O" DENOTE FOR OFFLINE STOCK TAKE USER
            If intCompany_No = 0 Then
                strLock_Information = Left(strLock_Information, 82) & Left(Trim(strNew_Value), 2) & Mid(strLock_Information, 85, Len(strLock_Information))
            Else
                strLock_Information = Left(gstrAsyncronous_String, 55) & Left(Trim(strNew_Value), 2) & Mid(gstrAsyncronous_String, 58, 7)
            End If
        ElseIf strAction = "R" Then 'HERE "R" DENOTE FOR ANDROID STOCK TAKE
            If intCompany_No = 0 Then
                strLock_Information = Left(strLock_Information, 81) & Left(Trim(strNew_Value), 1) & Mid(strLock_Information, 83, Len(strLock_Information))
            Else
                strLock_Information = Left(gstrAsyncronous_String, 58) & Left(Trim(strNew_Value), 1) & Mid(gstrAsyncronous_String, 60, 5)
            End If
        ElseIf strAction = "S" Then 'HERE "S" DENOTE FOR HISTORICAL DATA
            strLock_Information = Left(strLock_Information, 80) & Left(Trim(strNew_Value), 1) & Mid(strLock_Information, 82, Len(strLock_Information))
        ElseIf strAction = "T" Then 'HERE "T" DENOTE FOR PHY USER
            If intCompany_No = 0 Then
                strLock_Information = Left(strLock_Information, 78) & Left(Trim(strNew_Value), 2) & Mid(strLock_Information, 81, Len(strLock_Information))
            Else
                strLock_Information = Left(gstrAsyncronous_String, 59) & Left(Trim(strNew_Value), 2) & Mid(gstrAsyncronous_String, 62, 3)
            End If
        End If
    End Function

    Public Function ConvertVersion(ByVal strVersion As String) As String
        Dim lngCounter As Long
        Dim strVersionKey As String

        For lngCounter = 1 To Len(strVersion)
            If Mid(strVersion, lngCounter, 1) = "O" Then
                strVersionKey = strVersionKey & "A"
            ElseIf Mid(strVersion, lngCounter, 1) = "M" Then
                strVersionKey = strVersionKey & "B"
            ElseIf Mid(strVersion, lngCounter, 1) = "K" Then
                strVersionKey = strVersionKey & "C"
            ElseIf Mid(strVersion, lngCounter, 1) = "I" Then
                strVersionKey = strVersionKey & "D"
            ElseIf Mid(strVersion, lngCounter, 1) = "G" Then
                strVersionKey = strVersionKey & "E"
            ElseIf Mid(strVersion, lngCounter, 1) = "E" Then
                strVersionKey = strVersionKey & "F"
            ElseIf Mid(strVersion, lngCounter, 1) = "C" Then
                strVersionKey = strVersionKey & "G"
            ElseIf Mid(strVersion, lngCounter, 1) = "A" Then
                strVersionKey = strVersionKey & "H"
            ElseIf Mid(strVersion, lngCounter, 1) = "9" Then
                strVersionKey = strVersionKey & "I"
            ElseIf Mid(strVersion, lngCounter, 1) = "7" Then
                strVersionKey = strVersionKey & "J"
            ElseIf Mid(strVersion, lngCounter, 1) = "5" Then
                strVersionKey = strVersionKey & "K"
            ElseIf Mid(strVersion, lngCounter, 1) = "3" Then
                strVersionKey = strVersionKey & "L"
            ElseIf Mid(strVersion, lngCounter, 1) = "1" Then
                strVersionKey = strVersionKey & "M"
            ElseIf Mid(strVersion, lngCounter, 1) = "2" Then
                strVersionKey = strVersionKey & "N"
            ElseIf Mid(strVersion, lngCounter, 1) = "4" Then
                strVersionKey = strVersionKey & "O"
            ElseIf Mid(strVersion, lngCounter, 1) = "6" Then
                strVersionKey = strVersionKey & "P"
            ElseIf Mid(strVersion, lngCounter, 1) = "8" Then
                strVersionKey = strVersionKey & "Q"
            ElseIf Mid(strVersion, lngCounter, 1) = "0" Then
                strVersionKey = strVersionKey & "R"
            ElseIf Mid(strVersion, lngCounter, 1) = "B" Then
                strVersionKey = strVersionKey & "S"
            ElseIf Mid(strVersion, lngCounter, 1) = "D" Then
                strVersionKey = strVersionKey & "T"
            ElseIf Mid(strVersion, lngCounter, 1) = "F" Then
                strVersionKey = strVersionKey & "U"
            ElseIf Mid(strVersion, lngCounter, 1) = "H" Then
                strVersionKey = strVersionKey & "V"
            ElseIf Mid(strVersion, lngCounter, 1) = "J" Then
                strVersionKey = strVersionKey & "W"
            ElseIf Mid(strVersion, lngCounter, 1) = "L" Then
                strVersionKey = strVersionKey & "X"
            ElseIf Mid(strVersion, lngCounter, 1) = "N" Then
                strVersionKey = strVersionKey & "Y"
            ElseIf Mid(strVersion, lngCounter, 1) = "P" Then
                strVersionKey = strVersionKey & "Z"
            ElseIf Mid(strVersion, lngCounter, 1) = "Q" Then
                strVersionKey = strVersionKey & "1"
            ElseIf Mid(strVersion, lngCounter, 1) = "S" Then
                strVersionKey = strVersionKey & "2"
            ElseIf Mid(strVersion, lngCounter, 1) = "U" Then
                strVersionKey = strVersionKey & "3"
            ElseIf Mid(strVersion, lngCounter, 1) = "W" Then
                strVersionKey = strVersionKey & "4"
            ElseIf Mid(strVersion, lngCounter, 1) = "Y" Then
                strVersionKey = strVersionKey & "5"
            ElseIf Mid(strVersion, lngCounter, 1) = "R" Then
                strVersionKey = strVersionKey & "6"
            ElseIf Mid(strVersion, lngCounter, 1) = "T" Then
                strVersionKey = strVersionKey & "7"
            ElseIf Mid(strVersion, lngCounter, 1) = "V" Then
                strVersionKey = strVersionKey & "8"
            ElseIf Mid(strVersion, lngCounter, 1) = "X" Then
                strVersionKey = strVersionKey & "9"
            ElseIf Mid(strVersion, lngCounter, 1) = "Z" Then
                strVersionKey = strVersionKey & "0"
            ElseIf Mid(strVersion, lngCounter, 1) = "`" Then
                strVersionKey = strVersionKey & " "
            ElseIf Mid(strVersion, lngCounter, 1) = "@" Then
                strVersionKey = strVersionKey & "-"
            Else
                strVersionKey = strVersionKey & Mid(strVersion, lngCounter, 1)
            End If
        Next
        ConvertVersion = strVersionKey
    End Function

    Public Function Write_NetHasp_Companies_Info(ByVal intCompanies_Count As Integer, ByVal strCompany_String_Bytes As String) As Boolean
        Dim intStart_Position As Integer
        MemoHaspBuffer.txt = ""
        '    MsgBox Len(strCompany_String_Bytes)
        Service = IS_HASP
        Call hasp(Service, SeedCode, LptNum, Passw1, Passw2, p1, p2, p3, p4&)
        If p1 = 1 Then
            gobjLock_Type = Is_MemoHASP()

            If gobjLock_Type = enumLock_Type.LOCK_LOCAL_RED_HASP Or gobjLock_Type = enumLock_Type.LOCK_LOCAL_USB_RED_HASP Then   'NET HASP LOCK FOUND ON LOCAL MACHINE
                Service = WRITE_MEMO_BLOCK
                MemoHaspBuffer.txt = strCompany_String_Bytes
                If intCompanies_Count = 1 Then
                    p1 = 104       ' Block size in words!
                    p2 = 32       ' Block size in words!
                ElseIf intCompanies_Count = 2 Then
                    p1 = 136       ' Block size in words!
                    p2 = 32       ' Block size in words!
                ElseIf intCompanies_Count = 3 Then
                    p1 = 168       ' Block size in words!
                    p2 = 32       ' Block size in words!
                End If

                Call WriteHaspBlock(Service, MemoHaspBuffer, p2)
                Call hasp(Service, SeedCode, LptNum, Passw1, Passw2, p1, p2, p3, p4&)

                If p3 <> 0 Then
                    Write_NetHasp_Companies_Info = False
                    'MsgBox "Write Information failed. Error code = " & CStr(P3&), vbInformation, "Inst"
                Else
                    Write_NetHasp_Companies_Info = True
                End If
            End If
        Else
            Service = NET_LOGIN
            Call hasp(Service, SeedCode, ProgramNum, Passw1, Passw2, p1, p2, p3, p4&)
            Call hasp(NET_LAST_STATUS, SeedCode, ProgramNum, Passw1, Passw2, RetStatus&, dummy&, dummy&, dummy&)
            If (RetStatus = 0) Then 'NET HASP LOCK FOUND ON REMOTE MACHINE
                Service = NET_WRITE_MEMO_BLOCK
                MemoHaspBuffer.txt = Left(strCompany_String_Bytes, 48)
                If intCompanies_Count = 1 Then
                    intStart_Position = 104
                ElseIf intCompanies_Count = 2 Then
                    intStart_Position = 136
                ElseIf intCompanies_Count = 3 Then
                    intStart_Position = 168
                End If
                p1 = intStart_Position       ' Block size in words!
                p2 = 24                      ' Block size in words!
                p3 = 0
                p4& = p2

                Call WriteHaspBlock(Service, MemoHaspBuffer, p2)
                Call hasp(Service, SeedCode, ProgramNum, Passw1, Passw2, p1, p2, p3, p4&)

                MemoHaspBuffer.txt = Mid(strCompany_String_Bytes, 49, 48)
                p1 = intStart_Position + 24
                p2 = 8       ' Block size in words!

                p3 = 0
                p4& = p2
                Call WriteHaspBlock(Service, MemoHaspBuffer, p2)
                Call hasp(Service, SeedCode, ProgramNum, Passw1, Passw2, p1, p2, p3, p4&)

                If p3 <> 0 Then
                    Write_NetHasp_Companies_Info = False
                    'MsgBox "Write Information failed. Error code = " & CStr(P3&), vbInformation, "Inst"
                Else
                    Write_NetHasp_Companies_Info = True
                End If
            End If
        End If
    End Function

    Private Function Update_Lock_From_FTP(ByVal strUpdate_Values As String) As Boolean
        Dim blnFlag As Boolean
        Dim strSQL_String As String
        Dim intUpdate_Failed As Integer
        Dim strInitial As String
        Try
            intUpdate_Failed = 0
            blnFlag = False

            strInitial = Left(Trim(strUpdate_Values), 1)

            gstrCompany_Id = ConvertVersion(Mid(Trim(strUpdate_Values), 2, 4))

            If strInitial = "Q" Then
                strSQL_String = Mid(Trim(strUpdate_Values), 6, 5)
                If CDate(Format(Now(), "dd/MM/yyyy")) = CDate(ConvertStr(Mid(strSQL_String, 1, 2)) & "/" & ConvertStr(Mid(strSQL_String, 3, 2)) & "/" & ConvertYear(Mid(strSQL_String, 5, 1))) Then
                    ChangeLockInfo(0, gobjLock_Type, "$", gstrCompany_Id & "z")
                    blnFlag = True
                Else
                    intUpdate_Failed = 1
                End If
            ElseIf strInitial = "P" Then
                Dim blnFound As Boolean
                gstrCompanies_Block = gstrCompany_Block1 & gstrCompany_Block2 & gstrCompany_Block3 & gstrCompany_Desc_Block
                If InStr(1, gstrAllowed_Multiple_Companies_LockIds, gstrHasp_LockId) > 0 Then
                    'MsgBox "Secondary Company not allowed to write on whitelisted and disk Lock", vbCritical, "Alert"
                ElseIf gstrLock_Company_Id = gstrCompany_Id Then
                    'MsgBox "Secondary Company " & gstrCompany_Id & " Found in Lock Primary Location" & vbCrLf & vbCrLf & "hence cannot update Secondary company in Lock", vbCritical, "Alert"
                ElseIf Left(gstrCompanies_Block, 4) = gstrCompany_Id Then
                    'MsgBox "Secondary Company " & gstrCompany_Id & " Already Exists in Lock at Location 1" & vbCrLf & vbCrLf & "hence cannot update Secondary company in Lock", vbCritical, "Alert"
                ElseIf Mid(gstrCompanies_Block, 65, 4) = gstrCompany_Id Then
                    'MsgBox "Secondary Company " & gstrCompany_Id & " Already Exists in Lock at Location 2" & vbCrLf & vbCrLf & "hence cannot update Secondary company in Lock", vbCritical, "Alert"
                ElseIf Mid(gstrCompanies_Block, 129, 4) = gstrCompany_Id Then
                    'MsgBox "Secondary Company " & gstrCompany_Id & " Already Exists in Lock at Location 3" & vbCrLf & vbCrLf & "hence cannot update Secondary company in Lock", vbCritical, "Alert"
                Else
                    strSQL_String = Mid(Trim(strUpdate_Values), 6, 5)
                    If CDate(Format(Now(), "dd/MM/yyyy")) = CDate(ConvertStr(Mid(strSQL_String, 1, 2)) & "/" & ConvertStr(Mid(strSQL_String, 3, 2)) & "/" & ConvertYear(Mid(strSQL_String, 5, 1))) Then
                        Dim str() As String
                        str = Split(Trim(strUpdate_Values), "-")
                        If UBound(str) = 3 Then
                            If Len(str(1)) = 9 And Len(str(2)) <= 36 And Len(str(3)) = 12 Then

                                gstrAsyncronous_String = gstrCompany_Id & ConvertVersion(str(1)) & ConvertVersion(str(2) & Repl_String(36 - Len(str(2)), " ")) & ConvertVersion(str(3)) & Repl_String(3, " ")
                                If Len(Trim(Left(gstrCompanies_Block, 4))) = 0 Then
                                    If Write_NetHasp_Companies_Info(1, gstrAsyncronous_String) = True Then
                                        blnFlag = True
                                    Else
                                        intUpdate_Failed = 1
                                    End If
                                ElseIf Len(Trim(Mid(gstrCompanies_Block, 65, 4))) = 0 Then
                                    If Write_NetHasp_Companies_Info(2, gstrAsyncronous_String) = True Then
                                        blnFlag = True
                                    Else
                                        intUpdate_Failed = 1
                                    End If
                                ElseIf Len(Trim(Mid(gstrCompanies_Block, 129, 4))) = 0 Then
                                    If Write_NetHasp_Companies_Info(3, gstrAsyncronous_String) = True Then
                                        blnFlag = True
                                    Else
                                        intUpdate_Failed = 1
                                        'MsgBox "Write Information On Lock Failed", vbCritical, "Alert"
                                    End If
                                Else
                                    blnFlag = False
                                    'MsgBox "Lock Already Full with 4 Companies" & vbCrLf & vbCrLf & "hence cannot update Secondary company in Lock", vbCritical, "Alert"
                                End If
                            Else
                                blnFlag = False
                                'MsgBox "Invalid Secondary Serial Keys Block", vbExclamation, "Alert"
                            End If
                        Else
                            blnFlag = False
                            'MsgBox "Invalid Secondary Serial Keys", vbExclamation, "Alert"
                        End If
                    Else
                        blnFlag = False
                        'MsgBox "Invalid Secondary Serial Key 1", vbExclamation, "Alert"
                    End If
                End If
            ElseIf strInitial = "D" Then
                strSQL_String = Replace(Mid(Trim(strUpdate_Values), 7, Len(Trim(strUpdate_Values))), "-", "")
                If IsDate(ConvertStr(Mid(strSQL_String, 1, 2)) & "/" & ConvertStr(Mid(strSQL_String, 3, 2)) & "/" & ConvertYear(Mid(strSQL_String, 5, 1))) = False Then
                ElseIf IsDate(ConvertStr(Mid(strSQL_String, 6, 2)) & "/" & ConvertStr(Mid(strSQL_String, 8, 2)) & "/" & ConvertYear(Mid(strSQL_String, 10, 1))) = False Then
                Else
                    If ChangeLockInfo(0, gobjLock_Type, "D", ConvertStr(Mid(strSQL_String, 1, 2)) & ConvertStr(Mid(strSQL_String, 3, 2)) & ConvertYear(Mid(strSQL_String, 5, 1)) & ConvertStr(Mid(strSQL_String, 6, 2)) & ConvertStr(Mid(strSQL_String, 8, 2)) & ConvertYear(Mid(strSQL_String, 10, 1))) = True Then
                        blnFlag = True
                    Else
                        intUpdate_Failed = 1
                    End If
                End If
            ElseIf strInitial = "U" Then
                strSQL_String = Mid(Trim(strUpdate_Values), 6, 5)
                If CDate(Format(Now(), "dd/MMM/yyyy")) = CDate(ConvertStr(Mid(strSQL_String, 1, 2)) & "/" & MonthName(ConvertStr(Mid(strSQL_String, 3, 2)), True) & "/" & ConvertYear(Mid(strSQL_String, 5, 1))) Then

                    Dim adocommand As New SqlCommand("UPDATE [RetailSoft_Company].DBO.tblCompany_Detail SET Exe_Date = NULL ", adoCon_Company)
                    adocommand.CommandTimeout = 0
                    adocommand.ExecuteNonQuery()

                    adocommand = New SqlCommand("DELETE FROM [RetailSoft_Company].DBO.tblServer", adoCon_Company)
                    adocommand.CommandTimeout = 0
                    adocommand.ExecuteNonQuery()

                    blnFlag = True
                End If
            Else
                strSQL_String = Mid(Trim(strUpdate_Values), 6, 5)
                If CDate(Format(Now(), "dd/MMM/yyyy")) = CDate(ConvertStr(Mid(strSQL_String, 1, 2)) & "/" & MonthName(ConvertStr(Mid(strSQL_String, 3, 2)), True) & "/" & ConvertYear(Mid(strSQL_String, 5, 1))) Then
                    If ChangeLockInfo(0, gobjLock_Type, strInitial, ConvertVersion(Mid(Trim(strUpdate_Values), 11, Len(Trim(strUpdate_Values))))) = True Then
                        blnFlag = True
                    Else
                        intUpdate_Failed = 1
                    End If
                End If
            End If
            If InStr(1, "210241960,1860805334,669405361", gstrHasp_LockId) > 0 Then
                strInitial = "YSI:" & strInitial
            End If
            If blnFlag = False Then
                If intUpdate_Failed = 1 Then
                    gstrSMS_Message = "Retail-Soft Info Updated for " & gstrCompany_Id & " as Info " & strInitial & " Updation Failed By YSIPOS"
                Else
                    gstrSMS_Message = "Retail-Soft Info Updated for " & gstrCompany_Id & " as Back Date Key " & strInitial & " Failed By YSIPOS"
                End If
            Else
                gstrSMS_Message = "Retail-Soft Info Updated for " & gstrCompany_Id & " as Info " & strInitial & " Updation Successful By YSIPOS"
            End If

            Update_Lock_From_FTP = True
        Catch ex1 As Exception
            Print_Error_Only("Update_Lock_From_FTP", ex1)
            Update_Lock_From_FTP = False
        End Try
    End Function

    Private Function OMS_Event_SMS_Send()
        connetionString = "Data Source=" & gstrSQL_Server_Instance_Name & gstrSQL_Server_Port & ";Initial Catalog=OMSSoft_Company;User ID=" & gstrSQL_Instance_User_Name & ";Password=clsxls@login123;Application Name=Client_YSI"

        Dim adoSMS As New SqlConnection(connetionString)
        Try
            adoSMS.Open()
            connetionString = ""
            Dim adapter As New SqlDataAdapter
            Dim adoRs_SMS As New DataSet
            Dim strMessage As String
            Dim strMobile_No As String
            Dim intRepeat_Event_Reminder As Integer
            Dim strNext_SMS_On As String
            Dim lngEvent_Id As Long
            Dim intSMS_Days_After As Integer

            strSQL_String = "SELECT * FROM tblEvent_Mast WHERE Send_SMS=1 AND GETDATE() > Next_SMS_On AND (Repeat_Option=0 OR Remainder_End_Date >= CAST(GETDATE() AS DATE))"

            adapter.SelectCommand = New SqlCommand(strSQL_String, adoSMS)
            adapter.Fill(adoRs_SMS)
            adapter.Dispose()
            For i = 0 To adoRs_SMS.Tables(0).Rows.Count - 1
                With adoRs_SMS.Tables(0).Rows(i)
                    Dim adocommand As SqlCommand

                    lngEvent_Id = Val(.Item("Event_Id"))

                    strMobile_No = .Item("Mobile_List") & ""
                    strMessage = .Item("SMS_String") & ""
                    intRepeat_Event_Reminder = Val(.Item("Repeat_Event_Reminder") & "")
                    intSMS_Days_After = Val(.Item("SMS_Days_After") & "")
                    strNext_SMS_On = .Item("Next_SMS_On") & ""

                    strMessage = strMessage & " Date : " & Format(.Item("Event_Date"), "dd/MM/yyyy") & ""

                    If CDate(Now()) >= CDate(.Item("SMS_From_Date")) And CDate(Now()) <= CDate(.Item("SMS_Date_Time")) Then
                        strSQL_String = "BEGIN"
                        strSQL_String = strSQL_String & vbCrLf & "EXEC master.dbo.sp_configure 'show advanced options', 1"
                        strSQL_String = strSQL_String & vbCrLf & "EXEC master.dbo.sp_configure 'Ole Automation Procedures', 1"
                        strSQL_String = strSQL_String & vbCrLf & "END"

                        adocommand = New SqlCommand(strSQL_String, adoSMS)
                        adocommand.CommandTimeout = 0
                        adocommand.ExecuteNonQuery()

                        adocommand = New SqlCommand("RECONFIGURE", adoSMS)
                        adocommand.CommandTimeout = 0
                        adocommand.ExecuteNonQuery()

                        adocommand = New SqlCommand("EXEC spOMS_Send_SMS '" & strMobile_No & "','" & Replace(Trim(strMessage), "'", "''") & "', '1707166280617387783', 0", adoSMS)
                        adocommand.CommandTimeout = 0
                        adocommand.ExecuteNonQuery()
                    End If
                    strSQL_String = ""
                    If CDate(strNext_SMS_On) < CDate(.Item("SMS_Date_Time")) Then
                        strSQL_String = "UPDATE [OMSSoft_Company].dbo.tblEvent_Mast SET Next_SMS_On = DATEADD(dd,1,Next_SMS_On), Entry_Date = GETDATE() WHERE Event_Id = " & lngEvent_Id
                    ElseIf intRepeat_Event_Reminder = 1 Then
                        strSQL_String = "UPDATE [OMSSoft_Company].dbo.tblEvent_Mast SET Event_Date = DATEADD(ww,1,Event_Date), SMS_From_Date = DATEADD(ww,1,SMS_From_Date), SMS_Date_Time = DATEADD(ww,1,SMS_Date_Time), Next_SMS_On = DATEADD(ww,1,SMS_From_Date), Entry_Date = GETDATE() WHERE Event_Id = " & lngEvent_Id
                    ElseIf intRepeat_Event_Reminder = 2 Then
                        strSQL_String = "UPDATE [OMSSoft_Company].dbo.tblEvent_Mast SET Event_Date = DATEADD(mm,1,Event_Date), SMS_From_Date = DATEADD(mm,1,SMS_From_Date), SMS_Date_Time = DATEADD(mm,1,SMS_Date_Time), Next_SMS_On = DATEADD(mm,1,SMS_From_Date), Entry_Date = GETDATE() WHERE Event_Id = " & lngEvent_Id
                    ElseIf intRepeat_Event_Reminder = 3 Then
                        strSQL_String = "UPDATE [OMSSoft_Company].dbo.tblEvent_Mast SET Event_Date = DATEADD(qq,1,Event_Date), SMS_From_Date = DATEADD(qq,1,SMS_From_Date), SMS_Date_Time = DATEADD(qq,1,SMS_Date_Time), Next_SMS_On = DATEADD(qq,1,SMS_From_Date), Entry_Date = GETDATE() WHERE Event_Id = " & lngEvent_Id
                    ElseIf intRepeat_Event_Reminder = 4 Then
                        strSQL_String = "UPDATE [OMSSoft_Company].dbo.tblEvent_Mast SET Event_Date = DATEADD(mm,6,Event_Date), SMS_From_Date = DATEADD(mm,6,SMS_From_Date), SMS_Date_Time = DATEADD(mm,6,SMS_Date_Time), Next_SMS_On = DATEADD(mm,6,SMS_From_Date), Entry_Date = GETDATE() WHERE Event_Id = " & lngEvent_Id
                    ElseIf intRepeat_Event_Reminder = 5 Then
                        strSQL_String = "UPDATE [OMSSoft_Company].dbo.tblEvent_Mast SET Event_Date = DATEADD(yyyy,1,Event_Date), SMS_From_Date = DATEADD(yyyy,1,SMS_From_Date), SMS_Date_Time = DATEADD(yyyy,1,SMS_Date_Time), Next_SMS_On = DATEADD(yyyy,1,SMS_From_Date), Entry_Date = GETDATE() WHERE Event_Id = " & lngEvent_Id
                    Else
                        strSQL_String = "UPDATE [OMSSoft_Company].dbo.tblEvent_Mast SET Next_SMS_On = NULL, Entry_Date = GETDATE() WHERE Event_Id = " & lngEvent_Id
                    End If
                    If strSQL_String <> "" Then
                        adocommand = New SqlCommand(strSQL_String, adoSMS)
                        adocommand.CommandTimeout = 0
                        adocommand.ExecuteNonQuery()
                    End If
                End With
            Next
        Catch ex1 As Exception
            If connetionString = "" Then
                Print_Error_Only("OMS Event SMS Send", ex1)
            End If
        End Try
        adoSMS.Close()
        adoSMS.Dispose()
    End Function

    Private Function Send_SMS_SP(ByVal strMobile_No As String, ByVal strMessage As String)
        Try
            strSQL_String = "BEGIN"
            strSQL_String = strSQL_String & vbCrLf & "EXEC master.dbo.sp_configure 'show advanced options', 1"
            strSQL_String = strSQL_String & vbCrLf & "EXEC master.dbo.sp_configure 'Ole Automation Procedures', 1"
            strSQL_String = strSQL_String & vbCrLf & "END"

            Dim adocommand As New SqlCommand(strSQL_String, adoCon_Company)
            adocommand.CommandTimeout = 0
            adocommand.ExecuteNonQuery()

            adocommand = New SqlCommand("RECONFIGURE", adoCon_Company)
            adocommand.CommandTimeout = 0
            adocommand.ExecuteNonQuery()

            adocommand = New SqlCommand("EXEC spSend_SMS '" & strMobile_No & "','" & URLencshort(Replace(Trim(strMessage), "'", "")) & "', '1707162946121274007', 1", adoCon_Company)
            adocommand.CommandTimeout = 0
            adocommand.ExecuteNonQuery()

        Catch ex1 As Exception
            Print_Error_Only("Send_SMS_SP", ex1)
        End Try
    End Function

    Public Function URLencshort(ByRef Text As String) As String
        Dim lngA As Long, strChar As String
        Dim strURLencshort As String
        For lngA = 1 To Len(Text)
            strChar = Mid$(Text, lngA, 1)
            If strChar Like "[A-Za-z0-9]" Then
            ElseIf strChar = " " Then
                strChar = "+"
            Else
                strChar = "%" & Right$("0" & Hex$(Asc(strChar)), 2)
            End If

            strURLencshort = strURLencshort & strChar
        Next lngA
        URLencshort = strURLencshort
    End Function

    Private Sub Backup_Schedule()
        connetionString = "Data Source=" & gstrSQL_Server_Instance_Name & gstrSQL_Server_Port & ";Initial Catalog=OMSSoft_Company;User ID=" & gstrSQL_Instance_User_Name & ";Password=clsxls@login123;Application Name=Client_YSI"
        adoCon_Company = New SqlConnection(connetionString)
        Try
            adoCon_Company.Open()
            Dim adapter As New SqlDataAdapter
            Dim adoRS_Company As New DataSet
            Dim adocommand As SqlCommand
            Dim blnUpload_RSInfo As Boolean

            adapter.SelectCommand = New SqlCommand("SELECT * FROM tblCompany_Detail ORDER BY Company_Id", adoCon_Company)
            adapter.Fill(adoRS_Company)
            adapter.Dispose()
            For i = 0 To adoRS_Company.Tables(0).Rows.Count - 1
                With adoRS_Company.Tables(0).Rows(i)
                    Try
                        gintAttachment = 0
                        gstrPublication_Database = "OMSSoft-" & .Item("Company_Id") & "-" & .Item("Head_Office_Id") & "-" & .Item("Location_Id") & "-" & .Item("Financial_Year")

                        glngFinancial_Year = .Item("Financial_Year")

                        gstrCompany_Id = .Item("Company_Id") & ""

                        strHead_Office_Id = .Item("Head_Office_Id") & ""
                        strLocation_ID = .Item("Location_ID") & ""
                        gstrCompany_Name = .Item("Company_Desc").ToString() & ""

                        strHO_Company_Code = ""
                        Get_HO_Company_Code()
                        strExe_Version = ""
                        Get_Location_Go_Live()
                        blnUpload_RSInfo = False

                        If Download_RSInfo("/RETAIL_SOFT/" & IIf(strHO_Company_Code <> gstrCompany_Id, strHO_Company_Code & "/" & gstrCompany_Id, gstrCompany_Id) & "/Download", gstrCompany_Id) = True Then
                            blnUpload_RSInfo = True
                        End If
                        If Update_RSLog(gstrCompany_Id) = True Or blnUpload_RSInfo = True Then
                            If Search_Lock() = True Then
                                strSQL_String = "BEGIN"
                                strSQL_String = strSQL_String & vbCrLf & "  DELETE FROM [RetailSoft_Company].DBO.tblRsInfo_Up WHERE Company_Id='" & gstrCompany_Id & "'"
                                strSQL_String = strSQL_String & vbCrLf & "  INSERT INTO [RetailSoft_Company].DBO.tblRsInfo_Up(Company_Id, Exe_Date,Exe_Size,Exe_Version,IP_Address,Server_Name,Info,Company,HO,Branch,Server_info,Go_Live) "
                                strSQL_String = strSQL_String & vbCrLf & "  SELECT TOP 1 '" & gstrCompany_Id & "', Exe_Date,REPLACE(Exe_Size,',',''),'" & strExe_Version & "',IP_Address,Server_Name,'" & Replace(Replace(gstrLicense_Information, "'", ""), ",", "") & "','" & gstrCompany_Id & " - " & Replace(Replace(gstrCompany_Name, "'", ""), ",", "") & "','" & RTrim(strHO_Desc) & "','" & RTrim(strBranch_Desc) & "','Server Info (" & gstrHasp_LockId & ") On Date : ' + CONVERT(VARCHAR,GETDATE(),103) + ' ' + CONVERT(VARCHAR,GETDATE(),108),'" & strOpening_Date & "' FROM [RetailSoft_Company].DBO.tblServer"
                                strSQL_String = strSQL_String & vbCrLf & "END"

                                adocommand = New SqlCommand(strSQL_String, adoCon_Company)
                                adocommand.CommandTimeout = 0
                                adocommand.ExecuteNonQuery()
                            End If
                        End If

                        Upload_RSInfo("/RETAIL_SOFT/" & IIf(strHO_Company_Code <> gstrCompany_Id, strHO_Company_Code & "/" & gstrCompany_Id, gstrCompany_Id) & "/Upload", gstrCompany_Id)

                    Catch ex2 As Exception
                        Print_Error_Only("Server Schedule", ex2)
                    End Try
                End With
            Next
            'End If
        Catch ex1 As Exception
            Print_Error_Only("Server_Schedule! ", ex1)
        End Try

        adoCon_Company.Close()
        adoCon_Company.Dispose()
    End Sub

    Private Sub Server_Schedule()
        connetionString = "Data Source=" & gstrSQL_Server_Instance_Name & gstrSQL_Server_Port & ";Initial Catalog=RetailSoft_Company;User ID=" & gstrSQL_Instance_User_Name & ";Password=clsxls@login123;Application Name=Client_YSI"
        adoCon_Company = New SqlConnection(connetionString)
        Try
            adoCon_Company.Open()
            Dim adapter As New SqlDataAdapter
            Dim adoRS_Company As New DataSet
            Dim adocommand As SqlCommand
            Dim blnUpload_RSInfo As Boolean

            adapter.SelectCommand = New SqlCommand("SELECT * FROM tblCompany_Detail WHERE ISNULL(Subscriber,0) = 0 AND CAST(CAST(RIGHT(Financial_Year,4) AS VARCHAR(4))+'0401' AS INT) >= CAST(CONVERT(VARCHAR,GETDATE(),112) AS INT) AND (CAST(CONVERT(VARCHAR,GETDATE(),112) AS INT) BETWEEN CAST(CAST(LEFT(Financial_Year,4) AS VARCHAR(4))+'0401' AS INT) AND CAST(CAST(RIGHT(Financial_Year,4) AS VARCHAR(4))+'0401' AS INT)) ORDER BY Company_Id", adoCon_Company)
            adapter.Fill(adoRS_Company)
            adapter.Dispose()
            For i = 0 To adoRS_Company.Tables(0).Rows.Count - 1
                With adoRS_Company.Tables(0).Rows(i)
                    Try
                        gintAttachment = 0
                        gstrPublication_Database = "RetailSoft-" & .Item("Company_Id") & "-" & .Item("Head_Office_Id") & "-" & .Item("Location_Id") & "-" & .Item("Financial_Year")

                        glngFinancial_Year = .Item("Financial_Year")

                        gstrCompany_Id = .Item("Company_Id") & ""

                        strHead_Office_Id = .Item("Head_Office_Id") & ""
                        strLocation_ID = .Item("Location_ID") & ""
                        gstrCompany_Name = .Item("Company_Desc").ToString() & ""

                        strHO_Company_Code = ""
                        Get_HO_Company_Code()
                        strExe_Version = ""
                        Get_Location_Go_Live()
                        blnUpload_RSInfo = False

                        If Download_RSInfo("/RETAIL_SOFT/" & IIf(strHO_Company_Code <> gstrCompany_Id, strHO_Company_Code & "/" & gstrCompany_Id, gstrCompany_Id) & "/Download", gstrCompany_Id) = True Then
                            blnUpload_RSInfo = True
                        End If
                        If Update_RSLog(gstrCompany_Id) = True Or blnUpload_RSInfo = True Then
                            If Search_Lock() = True Then
                                strSQL_String = "BEGIN"
                                strSQL_String = strSQL_String & vbCrLf & "  DELETE FROM [RetailSoft_Company].DBO.tblRsInfo_Up WHERE Company_Id='" & gstrCompany_Id & "'"
                                strSQL_String = strSQL_String & vbCrLf & "  INSERT INTO [RetailSoft_Company].DBO.tblRsInfo_Up(Company_Id, Exe_Date,Exe_Size,Exe_Version,IP_Address,Server_Name,Info,Company,HO,Branch,Server_info,Go_Live) "
                                strSQL_String = strSQL_String & vbCrLf & "  SELECT TOP 1 '" & gstrCompany_Id & "', Exe_Date,REPLACE(Exe_Size,',',''),'" & strExe_Version & "',IP_Address,Server_Name,'" & Replace(Replace(gstrLicense_Information, "'", ""), ",", "") & "','" & gstrCompany_Id & " - " & Replace(Replace(gstrCompany_Name, "'", ""), ",", "") & "','" & RTrim(strHO_Desc) & "','" & RTrim(strBranch_Desc) & "','Server Info (" & gstrHasp_LockId & ") On Date : ' + CONVERT(VARCHAR,GETDATE(),103) + ' ' + CONVERT(VARCHAR,GETDATE(),108),'" & strOpening_Date & "' FROM [RetailSoft_Company].DBO.tblServer"
                                strSQL_String = strSQL_String & vbCrLf & "END"

                                adocommand = New SqlCommand(strSQL_String, adoCon_Company)
                                adocommand.CommandTimeout = 0
                                adocommand.ExecuteNonQuery()
                            End If
                        End If

                        Upload_RSInfo("/RETAIL_SOFT/" & IIf(strHO_Company_Code <> gstrCompany_Id, strHO_Company_Code & "/" & gstrCompany_Id, gstrCompany_Id) & "/Upload", gstrCompany_Id)

                    Catch ex2 As Exception
                        Print_Error_Only("Server Schedule", ex2)
                    End Try
                End With
            Next
            'End If
        Catch ex1 As Exception
            Print_Error_Only("Server_Schedule! ", ex1)
        End Try

        adoCon_Company.Close()
        adoCon_Company.Dispose()
    End Sub

    Private Function Get_Location_Go_Live() As Boolean
        connetionString = "Data Source=" & gstrSQL_Server_Instance_Name & gstrSQL_Server_Port & ";Initial Catalog=" & gstrPublication_Database & ";User ID=" & gstrSQL_Instance_User_Name & ";Password=clsxls@login123;Application Name=Client_YSI"
        Dim adoCon_Stock As New SqlConnection(connetionString)
        Try
            adoCon_Stock.Open()
            Dim adapter As New SqlDataAdapter
            Dim adoRs_Stock As DataSet

            strSQL_String = "Select Branch_Id, Branch_Desc, Opening_Date  FROM tblBranch_Mast WHERE Branch_Id='" & strLocation_ID & "'"

            command = New SqlCommand(strSQL_String, adoCon_Stock)
            command.CommandTimeout = 0
            adapter = New SqlDataAdapter
            adapter.SelectCommand = command
            adoRs_Stock = New DataSet
            adapter.Fill(adoRs_Stock)
            adapter.Dispose()
            adapter = Nothing
            command.Dispose()
            If adoRs_Stock.Tables(0).Rows.Count > 0 Then
                With adoRs_Stock.Tables(0).Rows(0)
                    strBranch_Desc = CStr(.Item("Branch_Id") & "") & " - " & Replace(Replace(CStr(.Item("Branch_Desc") & ""), "'", ""), ",", "")
                    If IsDate(.Item("Opening_Date") & "") = True Then
                        If CDate(.Item("Opening_Date")) > CDate("01/Apr/1990") Then
                            strOpening_Date = Format(.Item("Opening_Date"), "dd/MMM/yyyy")
                        End If
                    End If

                    Dim strRetail_Soft_File = My.Application.Info.DirectoryPath & "\Retail_Soft.exe"
                    Dim myFileVersionInfo As FileVersionInfo
                    If File.Exists(strRetail_Soft_File) Then
                        myFileVersionInfo = FileVersionInfo.GetVersionInfo(strRetail_Soft_File)
                        strExe_Version = (myFileVersionInfo.FileVersion.Split(" ")(0)).ToString
                    End If
                End With
            End If
        Catch ex1 As Exception
            Print_Error_Only("Get Location Go Live", ex1)
        End Try
        adoCon_Stock.Close()
        adoCon_Stock.Dispose()
    End Function

    Private Function Get_HO_Company_Code() As Boolean
        connetionString = "Data Source=" & gstrSQL_Server_Instance_Name & gstrSQL_Server_Port & ";Initial Catalog=" & gstrPublication_Database & ";User ID=" & gstrSQL_Instance_User_Name & ";Password=clsxls@login123;Application Name=Client_YSI"
        Dim adoCon_Stock As New SqlConnection(connetionString)
        Try
            adoCon_Stock.Open()
            Dim adapter As New SqlDataAdapter
            Dim adoRs_Stock As DataSet

            strSQL_String = "Select Company_Id, Branch_Id, Branch_Desc FROM tblBranch_Mast WHERE Branch_Id='" & strHead_Office_Id & "'"

            command = New SqlCommand(strSQL_String, adoCon_Stock)
            command.CommandTimeout = 0
            adapter = New SqlDataAdapter
            adapter.SelectCommand = command
            adoRs_Stock = New DataSet
            adapter.Fill(adoRs_Stock)
            adapter.Dispose()
            adapter = Nothing
            command.Dispose()
            If adoRs_Stock.Tables(0).Rows.Count > 0 Then
                With adoRs_Stock.Tables(0).Rows(0)
                    strHO_Company_Code = CStr(.Item("Company_Id") & "")
                    strHO_Desc = CStr(.Item("Branch_Id") & "") & " - " & Replace(Replace(CStr(.Item("Branch_Desc") & ""), "'", ""), ",", "")
                End With
            End If
        Catch ex1 As Exception
            Print_Error_Only("Get HO Company Code", ex1)
        End Try
        adoCon_Stock.Close()
        adoCon_Stock.Dispose()
    End Function

    Private Sub Refresh_Server_Data()
        If ReadINI() = True Then
            If UCase(Environment.MachineName) = UCase(gstrServer_Name) Then
                'Search_Lock()
                'If InStr(1, gstrAllowed_Multiple_Companies_LockIds, gstrHasp_LockId) = 0 Then
                '    Server_Schedule()
                'ElseIf gstrHasp_LockId = "1917058163" Then
                OMS_Event_SMS_Send()
                Restore_RSInfo()

                'End If
            End If
        End If
    End Sub

    Private Function Restore_RSInfo() As Boolean
        '''TS-103 | System should Read Ctrl+F4 information available on FTP Upload folder and update the same in OMS data for that client site.
        connetionString = "Data Source=" & gstrSQL_Server_Instance_Name & gstrSQL_Server_Port & ";Initial Catalog=OMSSoft_Central_YSIPL;User ID=" & gstrSQL_Instance_User_Name & ";Password=clsxls@login123;Application Name=Client_YSI"

        Dim adoRestore As New SqlConnection(connetionString)
        Try
            adoRestore.Open()

            connetionString = ""

            Dim adapter As New SqlDataAdapter
            Dim adoRs_Restore As New DataSet
            Dim strCompanies As String

            strCompanies = ""

            FTP_Folder_List("59.90.32.112", "FTP_User1", "Timken123#", True, "\Retail_Soft", strCompanies)

            strSQL_String = "SELECT Site_Id FROM tblClient_Site_Detail WHERE Site_Id IN (" & strCompanies & ") ORDER BY Site_Id"

            adapter.SelectCommand = New SqlCommand(strSQL_String, adoRestore)
            adapter.Fill(adoRs_Restore)
            adapter.Dispose()
            For i = 0 To adoRs_Restore.Tables(0).Rows.Count - 1
                With adoRs_Restore.Tables(0).Rows(i)
                    Dim strConsume_Folder As String
                    Dim strCurrentFile As String
                    Dim strSetup_Path As String
                    Dim strSite_Id As String
                    Dim strHO_Site() As String
                    Dim intUB As Integer

                    strSite_Id = .Item("Site_Id")

                    FTP_Folder_List("59.90.32.112", "FTP_User1", "Timken123#", True, "\Retail_Soft\" & strSite_Id, strCompanies)

                    If strCompanies <> "" Then
                        strHO_Site = Split(Replace(strCompanies, "'", ""), ",")
                        intUB = UBound(strHO_Site)
                    End If


                    strSetup_Path = "\Retail_Soft\" & strSite_Id & "\Upload"
HO_Site_Again:
                    Dirlist = New List(Of String) 'I prefer List() instead of an array
                    FTP_Folder_Files_List("59.90.32.112", "FTP_User1", "Timken123#", True, strSetup_Path, Dirlist)
                    If Dirlist.Count > 0 Then
                        strConsume_Folder = My.Application.Info.DirectoryPath & "\Log_Files"
                        Create_Folder_Or_Delete_Old_Files(strConsume_Folder)

                        For intx = 0 To (Dirlist.Count - 1)
                            strCurrentFile = Dirlist.Item(intx)
                            strCurrentFile = Mid(strCurrentFile, InStr(strCurrentFile, "/", CompareMethod.Text) + 1, Len(strCurrentFile))
                            If InStr(UCase(strCurrentFile), UCase("Restore_RSInfo_"), CompareMethod.Text) > 0 And InStr(UCase(strCurrentFile), UCase(".CSV"), CompareMethod.Text) > 0 Then
                                If File.Exists(strConsume_Folder & "/" & strCurrentFile) Then
                                    File.Delete(strConsume_Folder & "/" & strCurrentFile)
                                End If
                                If Get_FTP_File("59.90.32.112", "FTP_User1", "Timken123#", True, strSetup_Path & "/" & strCurrentFile, strConsume_Folder & "/" & strCurrentFile) = True Then
                                    If Import_CSV_File(adoRestore, strConsume_Folder & "/" & strCurrentFile, "Restore_RSInfo") = True Then
                                        If Execute_Multiple_Query(adoRestore, 1) = False Then
                                            If Execute_Multiple_Query(adoRestore, 2) = False Then
                                            Else
                                                Drop_FTP_File("59.90.32.112", "FTP_User1", "Timken123#", True, strSetup_Path & "/" & strCurrentFile)
                                            End If
                                        Else
                                            Drop_FTP_File("59.90.32.112", "FTP_User1", "Timken123#", True, strSetup_Path & "/" & strCurrentFile)
                                        End If
                                    End If
                                End If
                            Else
                                Drop_FTP_File("59.90.32.112", "FTP_User1", "Timken123#", True, strSetup_Path & "/" & strCurrentFile)
                            End If
                        Next
                    End If
                    If strCompanies <> "" And intUB >= 0 Then
                        strSetup_Path = "\Retail_Soft\" & strSite_Id & "\" & strHO_Site(intUB).ToString() & "\Upload"
                        intUB = intUB - 1
                        GoTo HO_Site_Again
                    End If

                End With
            Next
        Catch ex1 As Exception
            Print_Error_Only("Restore_RSInfo", ex1)
        End Try
    End Function

    Private Function Execute_Multiple_Query(ByRef adoRestore As SqlConnection, ByVal intFlag As Integer) As Boolean
        Try
            strSQL_String = "BEGIN "
            If intFlag = 1 Then
                strSQL_String = strSQL_String & vbCrLf & "SET DATEFORMAT DMY"
            Else
                strSQL_String = strSQL_String & vbCrLf & "SET DATEFORMAT MDY"
            End If
            strSQL_String = strSQL_String & vbCrLf & "DELETE FROM tblRsInfo_Up WHERE Company_Id IN (SELECT Company_Id FROM Restore_RSInfo)"
            strSQL_String = strSQL_String & vbCrLf & "INSERT INTO tblRsInfo_Up(Company_Id,Exe_Date,Exe_Size,Exe_Version,IP_Address,Server_Name,Info,Company,HO,Branch,Server_info,Go_Live)"
            strSQL_String = strSQL_String & vbCrLf & "SELECT Company_Id,Exe_Date,Exe_Size,Exe_Version,IP_Address,Server_Name,Info,Company,HO,Branch,Server_info,Go_Live from Restore_RSInfo"

            strSQL_String = strSQL_String & vbCrLf & ""
            strSQL_String = strSQL_String & vbCrLf & "UPDATE RS1 SET RS1.Entry_Date = RS2.Entry_Date, RS1.Server_info = LEFT(SERVER_INFO,CHARINDEX(':',SERVER_INFO)+1) + CONVERT(VARCHAR,RS2.Entry_Date,103) +' '+ CONVERT(VARCHAR,RS2.Entry_Date,108) FROM tblrsinfo_up RS1"
            strSQL_String = strSQL_String & vbCrLf & "INNER JOIN (select CAST(LTRIM(RTRIM(SUBSTRING(SERVER_INFO,CHARINDEX(':',SERVER_INFO)+1,LEN(SERVER_INFO)))) AS datetime) AS Entry_Date, Company_Id from tblrsinfo_up) RS2 ON RS2.Company_Id = RS1.Company_Id AND DAY(RS2.Entry_Date)<>1 WHERE RS1.Company_Id IN (SELECT Company_Id FROM Restore_RSInfo)"

            strSQL_String = strSQL_String & vbCrLf & "End"

            command = New SqlCommand(strSQL_String, adoRestore)
            command.CommandTimeout = 0
            command.ExecuteNonQuery()

            Update_Installation_Keys(adoRestore)
            Execute_Multiple_Query = True
        Catch ex1 As Exception
            Execute_Multiple_Query = False
        End Try
    End Function

    Private Function Update_Installation_Keys(ByRef adoRestore As SqlConnection)
        Try
            Dim adoRs_Stock As DataSet
            Dim strLicense_Information As String
            Dim strCompany_Id As String
            strSQL_String = "SELECT Company_Id, Info FROM Restore_RSInfo"

            command = New SqlCommand(strSQL_String, adoRestore)
            command.CommandTimeout = 0
            adapter = New SqlDataAdapter
            adapter.SelectCommand = command
            adoRs_Stock = New DataSet
            adapter.Fill(adoRs_Stock)
            adapter.Dispose()
            adapter = Nothing
            command.Dispose()
            If adoRs_Stock.Tables(0).Rows.Count > 0 Then
                With adoRs_Stock.Tables(0).Rows(0)
                    strLicense_Information = CStr(.Item("Info") & "")
                    strCompany_Id = CStr(.Item("Company_Id") & "")

                    strSQL_String = "UPDATE tblRsInfo_Up SET Install_Demo_Version = " & Val(Trim(Mid(strLicense_Information, 17, 1))) & ",Install_Print_Footer = " & Val(Trim(Mid(strLicense_Information, 17, 1))) & ",Install_Version = '" & Trim(Mid(strLicense_Information, 38, 36)) & "',Install_No_of_Company = " & Val(Trim(Mid(strLicense_Information, 18, 3))) & ",Install_POS_Lic = " & Val(Trim(Mid(strLicense_Information, 29, 3))) - Val(Trim(Mid(strLicense_Information, 32, 3))) & ",Install_BO_Lic=" & Val(Trim(Mid(strLicense_Information, 32, 3))) & ",Install_MPOS_Lic=" & Val(Trim(Mid(strLicense_Information, 89, 3))) & ",Install_Phy_Stk_User=" & Val(Trim(Mid(strLicense_Information, 79, 2))) & ",Install_HH_Phy_Stk_User=" & Val(Trim(Mid(strLicense_Information, 35, 3))) & ",Install_Offline_POS_Lic=" & Val(Trim(Mid(strLicense_Information, 86, 3))) & ",Install_Offline_Phy_Stk_User=" & Val(Trim(Mid(strLicense_Information, 83, 2))) & ",SAP_Setup_Type='" & Mid(strLicense_Information, 85, 1) & "',Historical_Data=" & Val(Trim(Mid(strLicense_Information, 81, 1))) & ",Android_Stock_Take=" & Val(Trim(Mid(strLicense_Information, 82, 1))) & " WHERE Company_Id = '" & strCompany_Id & "'"

                    command = New SqlCommand(strSQL_String, adoRestore)
                    command.CommandTimeout = 0
                    command.ExecuteNonQuery()
                End With
            End If

            Update_Installation_Keys = True
        Catch ex1 As Exception
            Update_Installation_Keys = False
        End Try
    End Function

    Private Sub Print_Error_Only(ByVal strSQL_String As String, ByRef objException As Exception)
        Dim fileLoc As String = My.Application.Info.DirectoryPath & "\RS_Inbound_Error_" & WeekdayName(Weekday(Now(), vbMonday), False, vbMonday) & ".txt"
        Dim fs As FileStream = Nothing
        Dim strHead As String
        strHead = strSQL_String
        strSQL_String = "Module        : " & strSQL_String

        If File.Exists(fileLoc) Then
            If FormatDateTime(FileDateTime(fileLoc), DateFormat.ShortDate) <> FormatDateTime(Now(), DateFormat.ShortDate) Then
                File.Delete(fileLoc)
            End If
        End If
        If (Not File.Exists(fileLoc)) Then
            fs = File.Create(fileLoc)
            fs.Close()
        End If

        If File.Exists(fileLoc) Then
            Using sw As StreamWriter = New StreamWriter(fileLoc, True)
                strSQL_String = strSQL_String & vbCrLf & "Date Time     : " & Now() & vbCrLf & "Error Message : " & objException.Message

                Dim Err_Trace As Diagnostics.StackTrace
                Err_Trace = New Diagnostics.StackTrace(objException, True)
                For Each sf As StackFrame In Err_Trace.GetFrames
                    If sf.GetFileLineNumber() > 0 Then
                        strSQL_String = strSQL_String & vbCrLf & "Error Line    : " & sf.GetFileLineNumber() & " Filename: " & IO.Path.GetFileName(sf.GetFileName) & Environment.NewLine
                    End If
                Next
                sw.WriteLine(strSQL_String)
            End Using
        End If
    End Sub


    Private Function Adodb_To_CSV(ByVal adoRs As DataSet, ByVal strFile_Name As String, ByVal strHeading As String)
        Dim varDATA As String
        Dim strValue As String
        If File.Exists(strFile_Name) Then
            File.Delete(strFile_Name)
        End If
        varDATA = ""
        Try
            For i = 0 To adoRs.Tables(0).Rows.Count - 1
                With adoRs.Tables(0).Rows(i)
                    For Each column In adoRs.Tables(0).Columns
                        If IsDBNull(.Item(column.Ordinal)) = True Then
                            strValue = ""
                        Else
                            strValue = Replace(CStr(.Item(column.Ordinal)), ",", "")
                        End If
                        varDATA = varDATA & IIf(column.Ordinal = 0, "", ",") & strValue
                    Next
                    varDATA = varDATA & vbCrLf
                End With
            Next
            varDATA = strHeading & vbCrLf & varDATA
            File.WriteAllText(strFile_Name, varDATA)
        Catch ex As Exception
            Print_Error_Only("Adodb_To_CSV", ex)
        End Try
    End Function

    Public Function Encrypt_User_Password(ByVal strUser_Password As String) As String
        Dim lngCounter As Long
        Dim lngEncrypt_Value As Long
        Dim strNew_User_Password As String

        lngCounter = 1
        strNew_User_Password = ""
        Do While lngCounter <= Len(strUser_Password)
            lngEncrypt_Value = Asc(Mid(strUser_Password, lngCounter, 1)) + 127
            strNew_User_Password = strNew_User_Password + Chr(lngEncrypt_Value)
            lngCounter = lngCounter + 1
        Loop
        Encrypt_User_Password = strNew_User_Password
    End Function

    Public Function Decrypt_User_Password(ByVal strUser_Password As String) As String
        Dim lngCounter As Long
        Dim lngDecrypt_Value As Long
        Dim strNew_User_Password As String

        lngCounter = 1
        strNew_User_Password = ""
        Do While lngCounter <= Len(strUser_Password)
            lngDecrypt_Value = Asc(Mid(strUser_Password, lngCounter, 1)) - 127
            If lngDecrypt_Value = 161 Then
                strNew_User_Password = strNew_User_Password + " "
            ElseIf lngDecrypt_Value < 0 Then
                strNew_User_Password = strNew_User_Password + Mid(strUser_Password, lngCounter, 1)
            Else
                strNew_User_Password = strNew_User_Password + Chr(lngDecrypt_Value)
            End If
            lngCounter = lngCounter + 1
        Loop
        Decrypt_User_Password = Trim(strNew_User_Password)
    End Function
End Class