'****************************************************************************
'**                                                                        **
'**                              API Demo                                  **
'**                    API-Calls for HASP and definitions                  **
'**                                                                        **
'**   This file contains some helpful defines to access a Hasp using	   **
'**   the application programing interface (API) for Hasp.		           **
'**                                                                        **
'**                       Aladdin Knowledge Systems                        **
'**                                                                        **/
'** $Id: AKSHasp.vb,v 1.1 2002/07/18 10:05:49 alex Exp $
'** $Date: 2002/07/18 10:05:49 $
'** $Name:  $
'** $Author: alex $
'**
'******************************************************************************
'* Revision history:
'* $Log: AKSHasp.vb,v $
'* Revision 1.1  2002/07/18 10:05:49  alex
'* initial check-in after rewrite with API 8.0 services
'*
'*
'* 
'****************************************************************************/

Option Strict Off
Option Explicit On
Public Module HASPVBNET
    '**********************************************************************
    '  VB-NET demonstration program for:
    '
    '   HASP4, HASP4 M1, HASP4 M4, HASP4 Time, HASP4 Net
    '   HASP-3, MemoHASP, MemoHASP36, TimeHASP, TimeHASP-4, NetHASP
    '**********************************************************************

    ' A list of HASP and MemoHASP Services.
    Public Const IS_HASP As Short = 1
    Public Const GET_HASP_CODE As Short = 2
    Public Const READ_MEMO As Short = 3
    Public Const WRITE_MEMO As Short = 4
    Public Const GET_HASP_STATUS As Short = 5
    Public Const GET_ID_NUM As Short = 6
    Public Const HASP_GENERATION As Short = 8
    Public Const HASP_NET_STATUS As Short = 9
    Public Const READ_MEMO_BLOCK As Short = 50
    Public Const WRITE_MEMO_BLOCK As Short = 51
    Public Const ENCODE_DATA As Short = 60
    Public Const DECODE_DATA As Short = 61

    ' A list of TimeHASP Services.
    Public Const TIMEHASP_SET_TIME As Short = 70
    Public Const TIMEHASP_GET_TIME As Short = 71
    Public Const TIMEHASP_SET_DATE As Short = 72
    Public Const TIMEHASP_GET_DATE As Short = 73
    Public Const TIMEHASP_WRITE_MEMORY As Short = 74
    Public Const TIMEHASP_READ_MEMORY As Short = 75
    Public Const TIMEHASP_WRITE_BLOCK As Short = 76
    Public Const TIMEHASP_READ_BLOCK As Short = 77
    Public Const TIMEHASP_GET_ID_NUM As Short = 78

    ' A list of NetHASP Services.
    Public Const NET_LAST_STATUS As Short = 40
    Public Const NET_GET_HASP_CODE As Short = 41
    Public Const NET_LOGIN As Short = 42
    Public Const NET_LOGIN_PROCESS As Short = 110
    Public Const NET_LOGOUT As Short = 43
    Public Const NET_READ_WORD As Short = 44
    Public Const NET_WRITE_WORD As Short = 45
    Public Const NET_GET_ID_NUMBER As Short = 46
    Public Const NET_SET_IDLE_TIME As Short = 48
    Public Const NET_READ_MEMO_BLOCK As Short = 52
    Public Const NET_WRITE_MEMO_BLOCK As Short = 53
    Public Const NET_SET_CONFIG_FILENAME As Short = 85
    Public Const NET_SET_SERVER_BY_NAME As Short = 96
    Public Const NET_ENCODE_DATA As Short = 88
    Public Const NET_DECODE_DATA As Short = 89
    Public Const NET_QUERY_LICENSE As Short = 104

    ' NetHASP error codes
    Public Const NET_READ_ERROR As Short = 131
    Public Const NET_WRITE_ERROR As Short = 132

    ' Error Codes

    '  ~~~~~~~~~~~
    '  Some symbols use the following abbreviations:
    '
    '  HDD    HASP Device Driver
    '  MH     MemoHASP
    '  TH     TimeHASP 
    '  NH     NetHASP
    '  NHLM   NetHASP License Manager
    '  NHCF   NetHASP Configuration File
    '  SSBN   Set Server By Name
    '
    Public Const OK As Short = 0
    Public Const HASPERR_SUCCESS As Short = 0   'Operation successful
    Public Const HASPERR_MH_TIMEOUT As Short = -1   'Timeout - Write operation failed
    Public Const HASPERR_MH_INVALID_ADDRESS As Short = -2   'Address out of range
    'Public Const HASPERR_HASP_NOT_FOUND As Short = -3  'old constant
    Public Const HASPERR_MH_INVALID_PASSWORDS As Short = -3 'A HASP with specified passwords was not found
    Public Const HASPERR_NOT_A_MEMOHASP As Short = -4   'A HASP was found but it is not a MemoHASP
    Public Const HASPERR_MH_WRITE_FAIL As Short = -5    'Unsuccessful Write operation
    Public Const HASPERR_PORT_BUSY As Short = -6    'Parallel port is busy.
    'Public Const DATA_TOO_SHORT As Short = -7  'old constant
    Public Const HASPERR_DATA_TOO_SHORT As Short = -7   'The data length is too short
    'Public Const HARDWARE_NOT_SUPPORTED As Short = -8  'old constant
    Public Const HASPERR_HARDWARE_NOT_SUPPORTED As Short = -8   'The hardware does not support the service
    ' Public Const INVALID_POINTER As Short = -9    'old constant
    Public Const HASPERR_INVALID_POINTER As Short = -9  'Invalid pointer used by Encode Data
    Public Const HASPERR_TS_FOUND As Short = -10    'Terminal Server was found.
    Public Const HASPERR_TS_SP3_FOUND As Short = -11    'Terminal Server under SP3 is not supported
    Public Const HASPERR_INVALID_PARAMETER As Short = -12   'Invalid parameter
    Public Const HASPERR_TH_INVALID_DAY As Short = -20  'Invalid day
    Public Const HASPERR_TH_INVALID_MONTH As Short = -21    'Invalid month
    Public Const HASPERR_TH_INVALID_YEAR As Short = -22 'Invalid year
    Public Const HASPERR_TH_INVALID_SECOND As Short = -23   'Invalid Second
    Public Const HASPERR_TH_INVALID_MINUTE As Short = -24   'Invalid Minute
    Public Const HASPERR_TH_INVALID_HOUR As Short = -25 'Invalid Hour
    Public Const HASPERR_TH_INVALID_ADDRESS As Short = -26  'Invalid address - Address is not in 0 - 15
    Public Const HASPERR_TH_TIMEOUT As Short = -27  'Timeout - Write operation failed
    Public Const HASPERR_TH_INVALID_PASSWORDS As Short = -28    'A HASP with specified passwords was not found
    Public Const HASPERR_NOT_A_TIMEHASP As Short = -29  'A HASP was found but it is not a TimeHASP
    Public Const HASPERR_CANT_OPEN_HDD As Short = -100  'Cannot open the HASP Device Driver
    Public Const HASPERR_CANT_READ_HDD As Short = -101  'Cannot read the HASP Device Driver
    Public Const HASPERR_CANT_CLOSE_HDD As Short = -102 'Cannot close the HASP Device Driver
    Public Const HASPERR_DOS_CANT_OPEN_HDD As Short = -110  'Cannot open the HASP Device Driver
    Public Const HASPERR_DOS_CANT_READ_HDD As Short = -111  'Cannot read the HASP Device Driver
    Public Const HASPERR_DOS_CANT_CLOSE_HDD As Short = -112 'Cannot close the HASP Device Driver
    Public Const HASPERR_CANT_ALLOC_DOSMEM As Short = -120  'Cannot allocate DOS Memory
    Public Const HASPERR_CANT_FREE_DOSMEM As Short = -121   'Cannot free DOS Memory
    Public Const HASPERR_INVALID_SERVICE As Short = -999    'HASP Invalid Servic
    'HASP4 Net error codes
    Public Const HASPERR_NO_PROTOCOLS As Short = 1  'IPX, NetBIOS, or TCP/IP protocols have not been installed properly.
    Public Const HASPERR_NO_SOCKET_NUMBER As Short = 2  'Communication Error - unable to get the socket number (TCP/IP, IPX)
    Public Const HASPERR_COMM_ERROR As Short = 3    'Communication Error. 
    Public Const HASPERR_NO_NHLM As Short = 4   'No NetHASP License Manager was found.
    Public Const HASPERR_NO_NHLM_ADDRFILE As Short = 5  'Cannot read NetHASP License Manager address file
    Public Const HASPERR_CANT_CLOSE_NHLM_ADDRFILE As Short = 6  'Cannot close NetHASP License Manager address file 
    Public Const HASPERR_CANT_SEND_PACKET As Short = 7  'Communication error - failed to send packet (IPX, NetBIOS)
    Public Const HASPERR_SILENT_NHLM As Short = 8   'No answer from the NetHASP License Manager. 
    Public Const HASPERR_NO_LOGIN As Short = 10 'Service requested before LOGIN 
    Public Const HASPERR_ADAPTER_ERROR As Short = 11    'NetBIOS: Communication error - adapter error 
    Public Const HASPERR_NO_ACTIVE_NHLM As Short = 15   'No active NetHASP Licence Manager was found 
    Public Const HASPERR_SSBN_FAILED As Short = 18  'Cannot perform LOGIN - SetServerByName failed 
    Public Const HASPERR_NHCF_SYNTAX_ERROR As Short = 19    'NetHASP configuration file syntax error 
    Public Const HASPERR_NHCF_GENERIC_ERROR As Short = 20   'Error handling NetHASP configuration file 
    Public Const HASPERR_NH_ENOMEM As Short = 21    'Memory allocation error 
    Public Const HASPERR_NH_CANT_FREE_MEM As Short = 22 'Memory release error 
    Public Const HASPERR_NH_INVALID_ADDRESS As Short = 23   'Invalid NetHASP memory address 
    Public Const HASPERR_NH_ENCDEC_ERR As Short = 24    'Error trying to Encrypt/Decrypt 
    Public Const HASPERR_CANT_LOAD_WINSOCK As Short = 25    'TCP/IP: failed to load WINSOCK.DLL 
    Public Const HASPERR_CANT_UNLOAD_WINSOCK As Short = 26  'TCP/IP: failed to unload WINSOCK.DLL 
    Public Const HASPERR_WINSOCK_ERROR As Short = 28    'TCP/IP: WINSOCK.DLL startup error 
    Public Const HASPERR_CANT_CLOSE_SOCKET As Short = 30    'TCP/IP: Failed to close socket. 
    Public Const HASPERR_SETPROTOCOL_FAILED As Short = 33   'SetProtocol service requested without performing LOGOUT 
    Public Const HASPERR_NH_NOT_SUPPORTED As Short = 40 'NetHASP services are not supported 
    Public Const HASPERR_NH_HASPNOTFOUND As Short = 129 'NetHASP key is not connected to the NetHASP Licence Manager 
    Public Const HASPERR_INVALID_PROGNUM As Short = 130 'Program Number is not in the Program List of the NetHASP memory 
    Public Const HASPERR_NH_READ_ERROR As Short = 131   'Error reading NetHASP memory 
    Public Const HASPERR_NH_WRITE_ERROR As Short = 132  'Error writing NetHASP memory 
    Public Const HASPERR_NO_MORE_STATIONS As Short = 133    'Number of stations exceeded 
    Public Const HASPERR_NO_MORE_ACTIVATIONS As Short = 134 'Number of activations exceeded 
    Public Const HASPERR_LOGOUT_BEFORE_LOGIN As Short = 135 'LOGOUT called without LOGIN 
    Public Const HASPERR_NHLM_BUSY As Short = 136   'NetHASP Licence Manager is busy 
    Public Const HASPERR_NHLM_FULL As Short = 137   'No space in NetHASP Log Table 
    Public Const HASPERR_NH_INTERNAL_ERROR As Short = 138   'NetHASP Internal error 
    Public Const HASPERR_NHLM_CRASHED As Short = 139    'NetHASP Licence Manager crashed and reactivated 
    Public Const HASPERR_NHLM_UNSUPPORTED_NETWORK As Short = 140    'NetHASP Licence Manager does not support the station's network 
    Public Const HASPERR_NH_INVALID_SERVICE_2 As Short = 141    'Invalid Service 
    Public Const HASPERR_NHCF_INVALID_NHLM As Short = 142   'Invalid NetHASP Licence Manager name in configuration file 
    Public Const HASPERR_SSBN_INVALID_NHLM As Short = 150   'Invalid NetHASP Licence Manager name in SetServerByName 
    Public Const HASPERR_ENC_ERROR_NHLM As Short = 152  'Error trying to encrypt by the LM 
    Public Const HASPERR_DEC_ERROR_NHLM As Short = 153  'Error trying to decrypt by the LM 
    Public Const HASPERR_OLD_LM_VERSION_NHLM As Short = 155 'LM old version was found 

    ' Error Strings
    Public Const STR_DATA_TOO_SHORT As String = "Data to Encode/Decode is too short."
    Public Const STR_HARDWARE_NOT_SUPPORTED As String = "Not a Marvin plug."
    Public Const STR_INVALID_POINTER As String = "Buffer pointer is invalid."
    Public Const STR_ENCODE_SUCCEEDED As String = "Encode Data OK."
    Public Const STR_DECODE_SUCCEEDED As String = "Decode Data OK."

    Public Const NO_HASP As String = "HASP plug not found !"

    ' A list of LptNum codes for the different types of keys
    Public Const LPT_IBM_ALL_HASP25 As Short = 0
    Public Const LPT_IBM_ALL_HASP36 As Short = 50
    Public Const LPT_NEC_ALL_HASP36 As Short = 60

    'MemoHASP maximum block size
    Public Const MEMO_BUFFER_SIZE As Short = 20

    'TimeHASP maximum block size
    Public Const TIME_BUFFER_SIZE As Short = 16

    ' hasp() function parameters
    Public ProgramNum As Short      'the programa number for net login
    Public Service As Short         'the service to be called
    Public SeedCode As Short        'the seed code
    Public LptNum As Short          'the port address for search
    Public Passw1 As Short          'first password
    Public Passw2 As Short          'second password
    Public p3, p1, p2, p4 As Integer
    Public RetStatus&, dummy&

    Public Enum enumLock_Type
        LOCK_NONE = 0
        LOCK_WINSOCK = 1
        LOCK_LOCAL_WHITE_HASP = 2
        LOCK_LOCAL_RED_HASP = 3
        LOCK_LOCAL_USB_RED_HASP = 4
        LOCK_REMOTE_RED_HASP = 5
        LOCK_LOCAL_SENTRY = 6
        LOCK_REMOTE_SENTRY = 7
        LOCK_DISK = 8
        LOCK_LOCAL_ROCKY = 9
        'LOCK_NET_ROCKY = 10 --before uncomment this check dmode.dat "storing single digit lock type"
    End Enum
    '
    ' The HASP memory buffer type.
    '
    Structure HBuff
        <VBFixedString(500), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValTStr, SizeConst:=500)> Dim txt As String
    End Structure

    '
    ' The TimeHASP memory buffer type.
    '
    Structure HTimeBuff
        <VBFixedString(20), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValTStr, SizeConst:=20)> Dim txt As String
    End Structure

    '
    ' The HASP Enc/Dec Buffer type.
    '
    Structure HEncDecBuffer
        <VBFixedString(4096), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValTStr, SizeConst:=4096)> Dim txt As String
    End Structure

    ' The Encode\Decode buffer.
    Public EncodeDecodeBuffer As HEncDecBuffer

    ' The MemoHasp buffer.
    Public MemoHaspBuffer As HBuff

    ' The TimeHasp buffer.
    Public TimeHaspBuffer As HTimeBuff

    '
    ' Use this declaration to call the hasp() routine directly.
    '

    ' The main hasp API function
    Declare Sub hasp Lib "haspvbnet32.dll" (ByVal Service As Integer, ByVal seed As Integer, ByVal lpt As Integer, ByVal pass1 As Integer, ByVal pass2 As Integer, ByRef retcode1 As Integer, ByRef retcode2 As Integer, ByRef retcode3 As Integer, ByRef retcode4 As Integer)

    ' WriteHaspBlock function prepares memory for the WriteBlock service
    ' WriteHaspBlock for MemoHasp
    Declare Sub WriteHaspBlock Lib "haspvbnet32.dll" (ByVal Service As Integer, ByRef Buff As HBuff, ByVal Length As Integer)
    ' Overload WriteHaspBlock for TimeHasp
    Declare Sub WriteHaspBlock Lib "haspvbnet32.dll" (ByVal Service As Integer, ByRef Buff As HTimeBuff, ByVal Length As Integer)

    ' ReadHaspBlock function prepares memory for the WriteBlock service
    ' ReadHaspBlock for MemoHasp
    Declare Sub ReadHaspBlock Lib "haspvbnet32.dll" (ByVal Service As Integer, ByRef Buff As HBuff, ByVal Length As Integer)
    ' Overload  ReadHaspBlock for TimeHasp
    Declare Sub ReadHaspBlock Lib "haspvbnet32.dll" (ByVal Service As Integer, ByRef Buff As HTimeBuff, ByVal Length As Integer)

    ' SetEncDecBlock function prepares memory for the Encode and Decode services    
    Declare Sub SetEncDecBlock Lib "haspvbnet32.dll" (ByRef Buff As HEncDecBuffer, ByVal Length As Integer)

    ' GetEncDecBlock function retrieves Encoded/Decoded block from memory
    Declare Sub GetEncDecBlock Lib "haspvbnet32.dll" (ByRef Buff As HEncDecBuffer, ByVal Length As Integer)

    Function LongToInteger(ByRef X As Integer) As Short

        Dim i As Integer

        i = X And CInt(65535)
        If i > 32767 Then i = i - 65536
        LongToInteger = i
    End Function

    Public Function Is_MemoHASP() As enumLock_Type
        'Is_MemoHASP = 1
        LptNum = 0
        Service = GET_HASP_STATUS
        Call hasp(Service, SeedCode, LptNum, Passw1, Passw2, p1, p2, p3, p4&)
        If p1 = 1 Then
            Is_MemoHASP = enumLock_Type.LOCK_LOCAL_WHITE_HASP
            LptNum = p3
        ElseIf p1 = 4 Then
            LptNum = p3
            If p3 >= 200 Then
                Is_MemoHASP = enumLock_Type.LOCK_LOCAL_USB_RED_HASP
            Else
                Is_MemoHASP = enumLock_Type.LOCK_LOCAL_RED_HASP
            End If
        ElseIf p2 = 3 Then
            Is_MemoHASP = enumLock_Type.LOCK_NONE
        End If
    End Function

    Public Function Repl_String(ByVal lngLen As Long, ByVal strChr As String) As String
        Dim lngCounter As Long
        Dim strTemp As String
        strTemp = ""
        For lngCounter = 1 To lngLen
            strTemp = strTemp & strChr
        Next
        Repl_String = strTemp
    End Function
End Module