Attribute VB_Name = "GLOBAL"




' Booleans
' NOTICE - Uncomment the next two lines of code if using
'          Visual BASIC Ver. 1.0
'Public Const TRUE = -1
'Public Const FALSE = 0

'Public Variables
Public ReturnValue1 As Integer
Public ReturnValue2 As Integer
Public Port As Integer

Public GCode1 As Long
Public GCode2 As Long
Public GCode3 As Long
Public GCode4 As Long
Public GAddress As Integer

'Public Constants
Public Const LocalMachine = 0

Public Const TERMINATE As Integer = -1
Public Const KLCHECK As Integer = 1
Public Const READAUTH As Integer = 2
Public Const GETSerialNumber As Integer = 3
Public Const GETVARWORD As Integer = 4
Public Const WRITEAUTHCODE As Integer = 5
Public Const WRITEVARWORD As Integer = 6
Public Const DECREMENTMEM As Integer = 7
Public Const GETEXPDATE As Integer = 8
Public Const CKLEASEDATE As Integer = 9
Public Const SETEXPIREDATE As Integer = 10
Public Const SETMAXUSERS As Integer = 11
Public Const GETMAXUSERS As Integer = 12
Public Const REMOTEUPDUPT1 As Integer = 13
Public Const REMOTEUPDUPT2 As Integer = 14
Public Const REMOTEUPDUPT3 As Integer = 15
Public Const REMOTEUPDCPT1 As Integer = 16
Public Const REMOTEUPDCPT2 As Integer = 17
Public Const REMOTEUPDCPT3 As Integer = 18

'Lease expiration and remote update constant
Public Const BaseYear As Integer = 1990

'Customer Unique Device Codes
Public Const ValidateCode1 As Integer = -8278
Public Const ValidateCode2 As Integer = -27256
Public Const ValidateCode3 As Integer = 27690
Public Const ClientIDCode1 As Integer = -19735
Public Const ClientIDCode2 As Integer = 32662
Public Const READCODE1 As Integer = 14053
Public Const READCODE2 As Integer = -20851
Public Const READCODE3 As Integer = 1659
Public Const WRITECODE1 As Integer = 23243
Public Const WRITECODE2 As Integer = 7454
Public Const WRITECODE3 As Integer = -20312

'Remote Update Task Codes
Public Const REMOTEADD As Integer = 0
Public Const REMOTEDATEEXTEND As Integer = 1
Public Const REMOTEOR As Integer = 2
Public Const REMOTEREPLACE As Integer = 3
Public Const REMOTEGETMEMORY As Integer = 4
Public Const REMOTESETUSERCOUNT As Integer = 5
Public Const REMOTEGETUSERCOUNT As Integer = 6
Public Const REMOTEGETDATE As Integer = 7
Public Const REMOTEINVALID As Integer = 8

'The DLL named KL2DLL32 provide KEY-LOK II device communications only on the machine on
'which the protected application is running.
'The DLL named NWKL2_32 provides KEY-LOK II device communications for both local, as well
'as on a remote computer on this network, provided both machines support the NETBIOS
'protocol and the computer with the device is running the PARCLASS.EXE server application.
Declare Function KFUNC Lib "KL2DLL32.DLL" Alias "_KFUNC@16" (ByVal Arg1 As Long, ByVal Arg2 As Long, ByVal Arg3 As Long, ByVal Arg4 As Long) As Long
Declare Function KEYBD Lib "KL2DLL32.DLL" Alias "_KEYBD@4" (ByVal Arg1 As Long) As Integer
Declare Function GETLASTKEYERROR Lib "KL2DLL32.DLL" Alias "_GETLASTKEYERROR@0" () As Long
'Declare Function KFUNC Lib "NWKL2_32.DLL" Alias "_KFUNC@16" (ByVal Arg1 As Long, ByVal Arg2 As Long, ByVal Arg3 As Long, ByVal Arg4 As Long) As Long
'Declare Function KEYBD Lib "NWKL2_32.DLL" Alias "_KEYBD@4" (ByVal Arg1 As Long) As Integer
'Declare Function GETLASTKEYERROR Lib "NWKL2_32.DLL" Alias "_GETLASTKEYERROR@0" () As Long

' Starting with the Win95 version, the 32 bit API can return the error
' code to caller using GetLastError.  System-generated error codes
' are defined in error.h;  User error codes are flagged by setting
' bit 29 in the status word, and the user must parse their meaning.

' GetLastError() returns codes so 32bit api caller can see results:

Public Const KEY_ERROR_NOERROR As Long = 0
Public Const KEY_ERROR_NOKEYLOK_ALSO As Long = 2
Public Const KEY_ERROR_NO_SESSIONS As Long = 5  ' All network sessions in use
Public Const KEY_ERROR_NOKEYLOK As Long = 536870913  ' No keylok found on port/lan.
Public Const KEY_ERROR_BADVERSION As Long = 536870914  ' Driver/kfunc version mismatch
Public Const KEY_ERROR_BADFUNC As Long = 536870915  ' Illegal kfunc command
Public Const KEY_ERROR_WRONGKEYLOK As Long = 536870916  ' Authenticate doesn't match
Public Const KEY_ERROR_NOREADAUTH As Long = 536870917  ' Read Auth wrong or not done
Public Const KEY_ERROR_NOWRITEAUTH As Long = 536870918  ' Write Auth wrong or not done
Public Const KEY_ERROR_INVALIDADDRESS As Long = 536870919  ' Bad KeyLok Device Address
Public Const KEY_ERROR_NOCOUNTSLEFT As Long = 536870920  ' Decrement Counter already 0
Public Const KEY_ERROR_WRITETIMEOUT As Long = 536870921  ' D0 did not go high in time
Public Const KEY_ERROR_NOKEYTAG As Long = 536870922  ' No keytag present
Public Const KEY_ERROR_LEASEOVER As Long = 536870928  ' Lease Expiration
Public Const KEY_ERROR_SYSDATESETBK As Long = 536870929  ' System Date Setback Error
Public Const KEY_ERROR_NOLEASEDATE As Long = 536870930  ' No Lease Date
Public Const KEY_ERROR_LEASEDATEBAD As Long = 536870931  ' Lease Date Bad
Public Const KEY_ERROR_FSDATEBAD As Long = 536870932      ' FS Date Bad
