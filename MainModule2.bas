Attribute VB_Name = "Main"
Option Explicit
Public Type POINTAPI
        X As Long
        Y As Long
End Type
Private Declare Function ClientToScreen Lib "user32" _
(ByVal hWnd As Long, lpPoint As POINTAPI) As Long

Private Declare Function SetCursorPos Lib "user32" _
(ByVal X As Long, ByVal Y As Long) As Long

Public igFlg_TCR As Integer
Public igShowPass As Integer
Public bgSw As Boolean
Public bgBlinkPag As Boolean
Public agArr_Pub1() As String
Public sgParms() As String
Public bgVideoLabel As Boolean
Public bgDiscLabel  As Boolean
Public sgCmdLine As String
Public bgExit As Boolean
Public igInd_Kar As Integer
Public bgPopular As Boolean  'Si esta activado al orden de popular
Public bgVIP As Boolean      'Si esta activado al orden de VIP
Public igKeyAscii As Integer 'Almacena el còdigo de la ùltima letra presionada
Public igMax_Dis  As Integer 'Máxima cantidad de registros (DB) físicos en Discos
Public igMax_Gen  As Integer 'Máxima cantidad de registros (DB) físicos en Géneros
Public igMax_Can  As Integer 'Máxima cantidad de registros (DB) físicos en Géneros
Public igMax_RgG As Integer  'Máximo de registros por pantalla de Género
Public igMax_RgD As Integer  'Máximo de registros por pantalla de Discos
Public igMax_RgC As Integer  'Máximo de registros por pantalla de Discos
Public igTot_PgG As Integer  'Total de paginas en la consulta (Generos)
Public igAct_PgG As Integer  'Página actual (Generos)
Public igTot_PgD As Integer  'Total de paginas en la consulta (Discos)
Public igAct_PgD As Integer  'Página actual (Discos)
Public igTot_PgC As Integer  'Total de paginas en la consulta (Discos)
Public igAct_PgC As Integer  'Página actual (Cansión)
Public sgIdx_Prm As Integer  'Máximo de temas para insertar un promo
Public igFlg_SavedCR As Integer
Public igStartPlayMusic As Integer
Public sgCr_AKey As String
Public igKeep_Cred As Integer
Public igNoDuplicT As Integer
Public igMixe_Popu As Integer
Public igLeftDisk As Integer
Public sgKb_Crd1 As String
Public sgKb_Crd2 As String
Public sgKb_ResM As String
Public sgKb_ResA As String
Public sgKb_BonC As String
Public sgKb_VID  As String
Public sgKb_Del  As String
Public sgKb_Ret  As String
Public sgKb_Pop  As String
Public sgKb_VIP  As String
Public sgKb_Vef  As String
Public sgKb_UP   As String
Public sgKb_Pause As String
Public sgKb_SwK As String
Public sgKb_DN   As String
Public sgKb_SwP   As String
Public sgFec_iAc As String
Public sgFec_Fac As String
Public sgSer_Mac As String
Public sgSer_CPU As String
Public sgNom_Loc As String
Public sgWin_Key As String
Public bFlagFoc As Boolean
Public sgDir_odb  As String
Public sgDir_Tmp  As String
Public sgDir_Fls  As String
Public sgDir_Fls2  As String
Public sgDir_Img  As String
Public sgDir_Mp3  As String
Public sgDir_Pub1  As String    'Directorio de publicidad #1
Public sgDir_Pub2  As String    'Directorio de publicidad #2
Public sgFle_Fon As String
Public igLim_Cred As Integer    'Limite de creditos
Public igInd_Pub As Integer     'índice de publicidad
Public igGen_Sel As String
Public igDis_Sel As String
Public igCan_Sel As String
Public igCnt_CR As Integer      'Contador de créditos.
Public igCnt_CRP As Integer     'Contador de créditos de prueba
Public igCnt_CRG As Long        'Contador de Créditos general.
Public sParam(1 To 54) As String

Public igNext_Return_Gen As Integer
Public igDelay_Return_Gen As Integer
Public igDelay_Return_Dis As Integer
Public igNext_Bonus   As Integer
Public igLen As Integer
Public bgSw_Pub As Boolean
Public igInd_Bon As Integer     'Indicador de número de bonus
Public igTot_Pub As Integer
Public vgForeColor As Variant
Public vgForeColorI As Variant
Public vTemp As Variant

Public Type tGeneric
    ID_ORD As String * 2
    ID_GEN As String * 2
    De_Gen As String * 20
    No_POS As Integer
End Type

Public Type tRegist
    Genero(1 To 15)  As tGeneric
    No_Rgs As Integer
    No_Pag As Integer
End Type
Public aPag_Gen(1 To 3) As tRegist  'Capacidad de 3 páginas, de 15 géneros por pantalla

Public Type gDisc
    ID_ORD As String * 2
    ID_DIS As String * 6
    NOM_DIS As String * 20
    NOM_ART As String * 20
    TX_COM As String * 20
    FL_IMG As String
    C_VIDEO As Integer
    X_POS As Integer
    No_POS As Integer
    FL_NEW As Integer
End Type

Public Type tReg_Disc
    Discos(1 To 12)  As gDisc
    No_Rgs As Integer
End Type
Public aPag_Disc(1 To 15) As tReg_Disc 'Capacidad de 15 páginas, de 12 carátulas por pantalla

Public Type gCanc
    ID_CAN As String * 6
    ID_GEN As String * 2
    ID_DIS As String * 6
    ID_ORD As String * 2
    DE_CAN As String * 60
    FL_IMG As String
    No_POS As Integer
End Type

Public Type tReg_Canc
    Cancion(1 To 30)  As gCanc
    No_Rgs As Integer
End Type
Public aPag_Canc(1 To 3) As tReg_Canc  'Capacidad de 3 páginas, de 30 canciones por pantalla

Public bgKeep_On_Top As Boolean
Public bgIs_Video As Boolean
Public bgWMP_Busy As Boolean
Public igCont_Sin As Integer
Public igDelay_Bonus_Vid As Integer
Public rtn As Long
Public igScr_Alone As Integer
'Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, _
'ByVal nSize As Long) As Long
Public Const MAX_PATH = 255


Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

Declare Sub Pause_2 Lib "kernel32" (ByVal dwMilliseconds As Long)

Public ogVFP9                 'Objeto de interconexion con Visual Foxpro 9.0
Global igMainCode As String

Public Const SWP_HIDEWINDOW = &H80
Public Const SWP_SHOWWINDOW = &H40
'Control Panel
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

'=====================================LLamado a librería para desactivar screen saver==========================
Public Declare Function SystemParametersInfo _
    Lib "user32" _
    Alias "SystemParametersInfoA" _
      (ByVal uiAction As Long, _
       ByVal uiParam As Long, _
       pvParam As Any, _
       ByVal fWInIni As Long) As Boolean

Public Const SPI_GETSCREENSAVEACTIVE As Long = &H10
Public Const SPI_GETSCREENSAVERRUNNING As Long = &H72
'================================================================================================================


' ================================================================================
'   GetTheComputerName AND GetTheWindowsDirectory declaration
' ================================================================================

Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetWindowsDirectory Lib "kernel32.dll" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

'Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" _
'        (ByVal lpBuffer As String, ByVal nSize As Long) As Long

' ================================================================================
' Routine:              GetVolumeInformation
' Description:
' Algorithm:
' Parameters:           None
' Returns:
' ================================================================================
Public Declare Function GetVolumeInformation Lib "kernel32" Alias _
        "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal _
        lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, _
        lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, _
        lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal _
        nFileSystemNameSize As Long) As Long

#If Win32 Then
    'Declaraciones para 32 bits
    Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
        (ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
         ByVal lpDefault As String, ByVal lpReturnedString As String, _
         ByVal nSize As Long, ByVal lpFileName As String) As Long
    Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
        (ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
         ByVal lpString As Any, ByVal lpFileName As String) As Long
#Else
    'Declaraciones para 16 bits
    Private Declare Function GetPrivateProfileString Lib "Kernel" _
        (ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
         ByVal lpDefault As String, ByVal lpReturnedString As String, _
         ByVal nSize As Integer, ByVal lpFileName As String) As Integer
    Private Declare Function WritePrivateProfileString Lib "Kernel" _
        (ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
         ByVal lpString As Any, ByVal lplFileName As String) As Integer
#End If
'END   GetVolumeInformation

'*********************REGISTER DLL LIBRARY************************************************************
Public Declare Function GetProcAddress _
  Lib "kernel32" _
  (ByVal hModule As Long, _
   ByVal lpProcName As String) _
  As Long

Public Declare Function LoadLibrary _
   Lib "kernel32" Alias "LoadLibraryA" _
   (ByVal lpLibFileName As String) As Long

Public Declare Function CreateThread Lib "kernel32" _
(lpThreadAttributes As Long, ByVal dwStackSize As Long, _
lpStartAddress As Long, lpParameter As Any, ByVal dwCreationFlags As Long, _
lpThreadId As Long) As Long

Public Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Public Declare Function GetExitCodeThread Lib "kernel32" (ByVal hThread As Long, lpExitCode As Long) As Long
Public Declare Sub ExitThread Lib "kernel32" (ByVal dwExitCode As Long)
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
'*********************REGISTER DLL LIBRARY************************************************************
'----------------------------------------------------------------------------
'Función equivalente a GetSetting de VB4.
'GetSetting     En VB4/32bits usa el registro.
'               En VB4/16bits usa un archivo de texto.
'Pero al usar las llamadas del API, siempre se escriben en archivos de texto.
'----------------------------------------------------------------------------
Public Function Read_Ini_File(lpFileName As String, lpAppName As String, lpKeyName As String, Optional vDefault) As String
    'Los parámetros son:
    'lpFileName:    La Aplicación (fichero INI)
    'lpAppName:     La sección que suele estar entrre corchetes
    'lpKeyName:     Clave
    'vDefault:      Valor opcional que devolverá
    '               si no se encuentra la clave.
    '
    Dim lpString As String
    Dim LTmp As Long
    Dim sRetVal As String
    
    'Si no se especifica el valor por defecto,
    'asignar incialmente una cadena vacía
    If IsMissing(vDefault) Then
        lpString = ""
    Else
        lpString = vDefault
    End If
    
    sRetVal = String$(255, 0)
    
    LTmp = GetPrivateProfileString(lpAppName, lpKeyName, lpString, sRetVal, Len(sRetVal), lpFileName)
    If LTmp = 0 Then
        Read_Ini_File = lpString
    Else
        Read_Ini_File = Left(sRetVal, LTmp)
    End If
End Function

'----------------------------------------------------------------------------
'Procedimiento equivalente a SaveSetting de VB4.
'SaveSetting    En VB4/32bits usa el registro.
'               En VB4/16bits usa un archivo de texto.
'Pero al usar las llamadas del API, siempre se escriben en archivos de texto.
'----------------------------------------------------------------------------
Sub Write_Ini_File(lpFileName As String, lpAppName As String, lpKeyName As String, lpString As String)
    'Guarda los datos de configuración
    'Los parámetros son los mismos que en LeerIni
    'Siendo lpString el valor a guardar
    '
    Dim LTmp As Long

    LTmp = WritePrivateProfileString(lpAppName, lpKeyName, lpString, lpFileName)
End Sub

Public Function Lee_Serial() As Variant
Dim oFs As Object
Dim oHD As Object
Set oFs = CreateObject("Scripting.filesystemobject")
Set oHD = oFs.getdrive(oFs.getdrivename(oFs.getabsolutepathname("C:")))
Lee_Serial = Hex(oHD.serialnumber)
End Function

Public Function Proper2(ByVal sCadena As String) As String
Proper2 = StrConv(sCadena, 3)
End Function

Public Sub Save_Defa_Path(Optional ByVal pRuta As String = "")
Dim sTmp1 As Variant
Dim sTmp2 As String
Dim bUpd_Ok As Boolean
bUpd_Ok = False
If pRuta = "" Then
    pRuta = App.Path
End If
If Not FileExist(App.Path & "\PathV2.ini") Then
    Call Write_Ini_File(pRuta & "\PathV2.ini", "PATHS", "DIR_ODB", pRuta & "\ODBC")
    Call Write_Ini_File(pRuta & "\PathV2.ini", "PATHS", "DIR_TMP", pRuta & "\TMP")
    Call Write_Ini_File(pRuta & "\PathV2.ini", "PATHS", "DIR_FL1", pRuta & "\FILES1")
    Call Write_Ini_File(pRuta & "\PathV2.ini", "PATHS", "DIR_FL2", pRuta & "\FILES2")
    Call Write_Ini_File(pRuta & "\PathV2.ini", "PATHS", "DIR_IMG", pRuta & "\fotos")
    Call Write_Ini_File(pRuta & "\PathV2.ini", "PATHS", "DIR_MP3", pRuta & "\cancionero")
    Call Write_Ini_File(pRuta & "\PathV2.ini", "PATHS", "DIR_PUB1", pRuta & "\PUBLICIDAD1")
    Call Write_Ini_File(pRuta & "\PathV2.ini", "PATHS", "DIR_PUB2", pRuta & "\PUBLICIDAD2")
    '----------------------------------------------------------------------------------'
    Call Write_Ini_File(pRuta & "\PathV2.ini", "TIMES", "TM_RET_GEN", "03")
    Call Write_Ini_File(pRuta & "\PathV2.ini", "TIMES", "TM_RET_DIS", "02")
    Call Write_Ini_File(pRuta & "\PathV2.ini", "TIMES", "TM_BON_VID", "20")
    Call Write_Ini_File(pRuta & "\PathV2.ini", "TIMES", "ID_CNT_PRO", "0")
    '----------------------------------------------------------------------------------'
    Call Write_Ini_File(pRuta & "\PathV2.ini", "ROCKOLA", "FILE_BACKG", pRuta & "\fondo8.JPG")
    Call Write_Ini_File(pRuta & "\PathV2.ini", "ROCKOLA", "SERIAL_MAC", "")
    Call Write_Ini_File(pRuta & "\PathV2.ini", "ROCKOLA", "SERIAL_CPU", "")
    Call Write_Ini_File(pRuta & "\PathV2.ini", "ROCKOLA", "NOMBRE_LOC", Scramble("SIN ASIGNACIÓN!"))
    Call Write_Ini_File(pRuta & "\PathV2.ini", "ROCKOLA", "FECHA_ACTI", "")
    Call Write_Ini_File(pRuta & "\PathV2.ini", "ROCKOLA", "FECHA_ACTF", "")
    Call Write_Ini_File(pRuta & "\PathV2.ini", "ROCKOLA", "WAPLIC_KEY", "")
    Call Write_Ini_File(pRuta & "\PathV2.ini", "ROCKOLA", "CR_ACC_KEY", "")
    Call Write_Ini_File(pRuta & "\PathV2.ini", "ROCKOLA", "SHOW_VIDLA", "0")
    Call Write_Ini_File(pRuta & "\PathV2.ini", "ROCKOLA", "SHOW_DISLA", "0")
    Call Write_Ini_File(pRuta & "\PathV2.ini", "ROCKOLA", "PASSW__BOX", "0")
    Call Write_Ini_File(pRuta & "\PathV2.ini", "ROCKOLA", "RELOAD_APP", "0")
    Call Write_Ini_File(pRuta & "\PathV2.ini", "ROCKOLA", "APPRUNNING", "0")
    Call Write_Ini_File(pRuta & "\PathV2.ini", "ROCKOLA", "START_MLST", "0")
    '----------------------------------------------------------------------------------'
    Call Write_Ini_File(pRuta & "\PathV2.ini", "GENERAL", "SHOW_ONTOP", "0")
    Call Write_Ini_File(pRuta & "\PathV2.ini", "GENERAL", "SCRN_ALONE", "1")
    Call Write_Ini_File(pRuta & "\PathV2.ini", "GENERAL", "NDUP_TEMES", "0")
    Call Write_Ini_File(pRuta & "\PathV2.ini", "GENERAL", "SWITCH_PUB", "1")
    Call Write_Ini_File(pRuta & "\PathV2.ini", "GENERAL", "SWITCH_KAR", "0")
    '----------------------------------------------------------------------------------'
    Call Write_Ini_File(pRuta & "\PathV2.ini", "CREDITS", "ACU_SAVECR", "0")
    Call Write_Ini_File(pRuta & "\PathV2.ini", "CREDITS", "FLG_SAVECR", "1")
    Call Write_Ini_File(pRuta & "\PathV2.ini", "CREDITS", "KEEP_SCRED", "0")
    Call Write_Ini_File(pRuta & "\PathV2.ini", "CREDITS", "LIMCR_CRED", "20")
    Call Write_Ini_File(pRuta & "\PathV2.ini", "CREDITS", "BONUS_CRED", "0")
    Call Write_Ini_File(pRuta & "\PathV2.ini", "CREDITS", "VIDEO_CRED", "0")
    Call Write_Ini_File(pRuta & "\PathV2.ini", "CREDITS", "CEDIT_ACAN", "0")
    Call Write_Ini_File(pRuta & "\PathV2.ini", "CREDITS", "CEDIT_ACAC", "0")
    '----------------------------------------------------------------------------------'
    Call Write_Ini_File(pRuta & "\PathV2.ini", "THEMES", "POPULAR_MIXER", "1")
    '--------------------------------------------------------------------------------------
    Call Write_Ini_File(pRuta & "\PathV2.ini", "KEYBOARD", "KB_ADD_01", "+")
    Call Write_Ini_File(pRuta & "\PathV2.ini", "KEYBOARD", "KB_ADD_03", "N")
    Call Write_Ini_File(pRuta & "\PathV2.ini", "KEYBOARD", "KB_DEL_01", "D")
    Call Write_Ini_File(pRuta & "\PathV2.ini", "KEYBOARD", "KB_RET_01", ".")
    Call Write_Ini_File(pRuta & "\PathV2.ini", "KEYBOARD", "KB_RST_01", "S")
    Call Write_Ini_File(pRuta & "\PathV2.ini", "KEYBOARD", "KB_RST_03", "R")
    Call Write_Ini_File(pRuta & "\PathV2.ini", "KEYBOARD", "KB_POP_01", "P")
    Call Write_Ini_File(pRuta & "\PathV2.ini", "KEYBOARD", "KB_VIP_01", "V")
    Call Write_Ini_File(pRuta & "\PathV2.ini", "KEYBOARD", "KB_PUP_01", "/")
    Call Write_Ini_File(pRuta & "\PathV2.ini", "KEYBOARD", "KB_PDN_01", "*")
    Call Write_Ini_File(pRuta & "\PathV2.ini", "KEYBOARD", "KB_VERIFY", "H")
    Call Write_Ini_File(pRuta & "\PathV2.ini", "KEYBOARD", "KB_SWTPUB", "W")
    Call Write_Ini_File(pRuta & "\PathV2.ini", "KEYBOARD", "KB__PAUSE", "L")
    Call Write_Ini_File(pRuta & "\PathV2.ini", "KEYBOARD", "KB_SWTKAR", "K")
    '--------------------------------------------------------------------------------------
End If
End Sub

Public Sub Save_Defa_Path2(Optional ByVal pRuta As String = "")
Dim sTmp1 As Variant
Dim sTmp2 As String
Dim bUpd_Ok As Boolean
bUpd_Ok = False
If pRuta = "" Then
    pRuta = App.Path
End If
Call Write_Ini_File(pRuta & "\PathV2.ini", "PATHS", "DIR_ODB", sgDir_odb)
Call Write_Ini_File(pRuta & "\PathV2.ini", "PATHS", "DIR_TMP", sgDir_Tmp)
Call Write_Ini_File(pRuta & "\PathV2.ini", "PATHS", "DIR_FL1", sgDir_Fls)
Call Write_Ini_File(pRuta & "\PathV2.ini", "PATHS", "DIR_FL2", sgDir_Fls2)
Call Write_Ini_File(pRuta & "\PathV2.ini", "PATHS", "DIR_IMG", sgDir_Img)
Call Write_Ini_File(pRuta & "\PathV2.ini", "PATHS", "DIR_MP3", sgDir_Mp3)
Call Write_Ini_File(pRuta & "\PathV2.ini", "PATHS", "DIR_PUB1", sgDir_Pub1)
Call Write_Ini_File(pRuta & "\PathV2.ini", "PATHS", "DIR_PUB2", sgDir_Pub2)
Call Write_Ini_File(pRuta & "\PathV2.ini", "PATHS", "FILE_BACKG", sgFle_Fon)
'---------------------------------------------------------------------------------------------------------------
Call Write_Ini_File(pRuta & "\PathV2.ini", "TIMES", "TM_RET_GEN", VBA.Format(igDelay_Return_Gen, "#####0"))
Call Write_Ini_File(pRuta & "\PathV2.ini", "TIMES", "TM_RET_DIS", VBA.Format(igDelay_Return_Dis, "#####0"))
Call Write_Ini_File(pRuta & "\PathV2.ini", "TIMES", "TM_BON_VID", VBA.Format(igDelay_Bonus_Vid, "#####0"))
Call Write_Ini_File(pRuta & "\PathV2.ini", "TIMES", "ID_CNT_PRO", VBA.Format(sgIdx_Prm, "#####0"))
'---------------------------------------------------------------------------------------------------------------
Call Write_Ini_File(pRuta & "\PathV2.ini", "ROCKOLA", "SHOW_VIDLA", VBA.Format(IIf(bgVideoLabel = True, 1, 0), "#####0"))
Call Write_Ini_File(pRuta & "\PathV2.ini", "ROCKOLA", "SHOW_DISLA", VBA.Format(IIf(bgDiscLabel = True, 1, 0), "#####0"))
Call Write_Ini_File(pRuta & "\PathV2.ini", "ROCKOLA", "SHOW_DISLA", VBA.Format(IIf(bgDiscLabel = True, 1, 0), "#####0"))
Call Write_Ini_File(pRuta & "\PathV2.ini", "ROCKOLA", "CR_ACC_KEY", Scramble(sgCr_AKey))
Call Write_Ini_File(pRuta & "\PathV2.ini", "ROCKOLA", "PASSW__BOX", VBA.Format(igShowPass, "#####0"))
Call Write_Ini_File(pRuta & "\PathV2.ini", "ROCKOLA", "START_MLST", VBA.Format(igStartPlayMusic, "#####0"))
'---------------------------------------------------------------------------------------------------------------
Call Write_Ini_File(pRuta & "\PathV2.ini", "CREDITS", "LIMCR_CRED", VBA.Format(igLim_Cred, "#####0"))
Call Write_Ini_File(pRuta & "\PathV2.ini", "CREDITS", "KEEP_SCRED", VBA.Format(igKeep_Cred, "0"))
Call Write_Ini_File(pRuta & "\PathV2.ini", "CREDITS", "BONUS_CRED", VBA.Format(sgKb_BonC, "#####0"))
Call Write_Ini_File(pRuta & "\PathV2.ini", "CREDITS", "VIDEO_CRED", VBA.Format(sgKb_VID, "#####0"))
Call Write_Ini_File(pRuta & "\PathV2.ini", "CREDITS", "VIDEO_CRED", VBA.Format(sgKb_VID, "#####0"))
Call Write_Ini_File(pRuta & "\PathV2.ini", "CREDITS", "FLG_SAVECR", VBA.Format(igFlg_SavedCR, "#####0"))
'---------------------------------------------------------------------------------------------------------------
Call Write_Ini_File(pRuta & "\PathV2.ini", "THEMES", "POPULAR_MIXER", VBA.Format(igMixe_Popu, "#####0"))
'---------------------------------------------------------------------------------------------------------------
Call Write_Ini_File(pRuta & "\PathV2.ini", "KEYBOARD", "KB_ADD_01", sgKb_Crd1)
Call Write_Ini_File(pRuta & "\PathV2.ini", "KEYBOARD", "KB_ADD_03", sgKb_Crd2)
Call Write_Ini_File(pRuta & "\PathV2.ini", "KEYBOARD", "KB_DEL_01", sgKb_Del)
Call Write_Ini_File(pRuta & "\PathV2.ini", "KEYBOARD", "KB_RET_01", sgKb_Ret)
Call Write_Ini_File(pRuta & "\PathV2.ini", "KEYBOARD", "KB_RST_01", sgKb_ResM)
Call Write_Ini_File(pRuta & "\PathV2.ini", "KEYBOARD", "KB_RST_03", sgKb_ResA)
Call Write_Ini_File(pRuta & "\PathV2.ini", "KEYBOARD", "KB_POP_01", sgKb_Pop)
Call Write_Ini_File(pRuta & "\PathV2.ini", "KEYBOARD", "KB_VIP_01", sgKb_VIP)
Call Write_Ini_File(pRuta & "\PathV2.ini", "KEYBOARD", "KB_PUP_01", sgKb_UP)
Call Write_Ini_File(pRuta & "\PathV2.ini", "KEYBOARD", "KB_PDN_01", sgKb_DN)
Call Write_Ini_File(pRuta & "\PathV2.ini", "KEYBOARD", "KB_VERIFY", sgKb_Vef)
Call Write_Ini_File(pRuta & "\PathV2.ini", "KEYBOARD", "KB_SWTPUB", sgKb_SwP)
Call Write_Ini_File(pRuta & "\PathV2.ini", "KEYBOARD", "KB__PAUSE", sgKb_Pause)
Call Write_Ini_File(pRuta & "\PathV2.ini", "KEYBOARD", "KB_SWTKAR", sgKb_SwK)
'---------------------------------------------------------------------------------------------------------------
Call Write_Ini_File(pRuta & "\PathV2.ini", "GENERAL", "SCRN_ALONE", VBA.Format(igScr_Alone, "#####0"))
Call Write_Ini_File(pRuta & "\PathV2.ini", "GENERAL", "SHOW_ONTOP", VBA.Format(bgKeep_On_Top, "#####0"))
Call Write_Ini_File(pRuta & "\PathV2.ini", "GENERAL", "NDUP_TEMES", VBA.Format(igNoDuplicT, "#####0"))
Call Write_Ini_File(pRuta & "\PathV2.ini", "GENERAL", "SWITCH_PUB", VBA.Format(IIf(bgSw_Pub = True, 1, 0), "#####0"))
Call Write_Ini_File(pRuta & "\PathV2.ini", "GENERAL", "SWITCH_KAR", VBA.Format(igInd_Kar, "#####0"))
End Sub

Public Sub Upd_Cnt()
Call Write_Ini_File(App.Path & "\PathV2.ini", "GENERAL", "SWITCH_PUB", VBA.Format(IIf(bgSw_Pub = True, 1, 0), "#####0"))
End Sub

Public Sub Upd_Path(ByVal pRuta As String, ByRef pVariables() As String)
Call Write_Ini_File(pRuta & "\PathV2.ini", "Rockola", "FECHA_ACTI", Scramble(pVariables(7)))
Call Write_Ini_File(pRuta & "\PathV2.ini", "Rockola", "FECHA_ACTF", Scramble(pVariables(8)))
Call Write_Ini_File(pRuta & "\PathV2.ini", "Rockola", "SERIAL_CPU", Scramble(pVariables(14)))
Call Write_Ini_File(pRuta & "\PathV2.ini", "Rockola", "SERIAL_MAC", Scramble(pVariables(9)))
Call Write_Ini_File(pRuta & "\PathV2.ini", "Rockola", "NOMBRE_LOC", Scramble(pVariables(10)))
Call Write_Ini_File(pRuta & "\PathV2.ini", "Rockola", "WAPLIC_KEY", Scramble(pVariables(43)))
End Sub

Public Sub Get_System_Path(ByRef paPath() As String, Optional ByVal pRuta As String = "")
If pRuta = "" Then
    pRuta = App.Path
End If
paPath(1) = Read_Ini_File(pRuta & "\PathV2.ini", "PATHS", "DIR_ODB", pRuta)
paPath(2) = Read_Ini_File(pRuta & "\PathV2.ini", "PATHS", "DIR_TMP", pRuta)
paPath(3) = Read_Ini_File(pRuta & "\PathV2.ini", "PATHS", "DIR_FL1", pRuta)
sgDir_Fls2 = Read_Ini_File(pRuta & "\PathV2.ini", "PATHS", "DIR_FL2", pRuta)
paPath(4) = Read_Ini_File(pRuta & "\PathV2.ini", "PATHS", "DIR_IMG", pRuta)
paPath(5) = Read_Ini_File(pRuta & "\PathV2.ini", "PATHS", "DIR_MP3", pRuta)
paPath(6) = Read_Ini_File(pRuta & "\PathV2.ini", "PATHS", "DIR_PUB1", pRuta)
paPath(50) = Read_Ini_File(pRuta & "\PathV2.ini", "PATHS", "DIR_PUB2", pRuta)
paPath(13) = Read_Ini_File(pRuta & "\PathV2.ini", "PATHS", "FILE_BACKG", "")
'--------------------------------------------------------------------------------------
paPath(7) = Read_Ini_File(pRuta & "\PathV2.ini", "ROCKOLA", "FECHA_ACTI", "")
paPath(7) = UnScramble(paPath(7))
paPath(8) = Read_Ini_File(pRuta & "\PathV2.ini", "ROCKOLA", "FECHA_ACTF", "")
paPath(8) = UnScramble(paPath(8))
paPath(9) = Read_Ini_File(pRuta & "\PathV2.ini", "ROCKOLA", "SERIAL_MAC", "")
paPath(9) = UnScramble(paPath(9))
paPath(10) = Read_Ini_File(pRuta & "\PathV2.ini", "ROCKOLA", "NOMBRE_LOC", "")
paPath(10) = UnScramble(paPath(10))
paPath(14) = Read_Ini_File(pRuta & "\PathV2.ini", "ROCKOLA", "SERIAL_CPU", "")
paPath(14) = UnScramble(paPath(14))
paPath(43) = Read_Ini_File(pRuta & "\PathV2.ini", "ROCKOLA", "WAPLIC_KEY", "")
paPath(43) = UnScramble(paPath(43))
paPath(44) = Read_Ini_File(pRuta & "\PathV2.ini", "ROCKOLA", "SHOW_VIDLA", "0")
paPath(45) = Read_Ini_File(pRuta & "\PathV2.ini", "ROCKOLA", "SHOW_DISLA", "0")
paPath(52) = Read_Ini_File(pRuta & "\PathV2.ini", "ROCKOLA", "CR_ACC_KEY", "")
paPath(52) = UnScramble(paPath(52))
igShowPass = Read_Ini_File(pRuta & "\PathV2.ini", "ROCKOLA", "PASSW__BOX", "0")
igStartPlayMusic = VBA.Val(Read_Ini_File(pRuta & "\PathV2.ini", "ROCKOLA", "START_MLST", "0"))
'--------------------------------------------------------------------------------------
paPath(46) = Read_Ini_File(pRuta & "\PathV2.ini", "GENERAL", "SHOW_ONTOP", "0")
paPath(47) = Read_Ini_File(pRuta & "\PathV2.ini", "GENERAL", "SCRN_ALONE", "0")
paPath(48) = Read_Ini_File(pRuta & "\PathV2.ini", "GENERAL", "NDUP_TEMES", "0")
paPath(49) = Read_Ini_File(pRuta & "\PathV2.ini", "GENERAL", "SWITCH_PUB", "0")
igInd_Kar = VBA.Val(Read_Ini_File(pRuta & "\PathV2.ini", "GENERAL", "SWITCH_KAR"))
'--------------------------------------------------------------------------------------
paPath(15) = Read_Ini_File(pRuta & "\PathV2.ini", "TIMES", "TM_RET_GEN", "0")
paPath(16) = Read_Ini_File(pRuta & "\PathV2.ini", "TIMES", "TM_RET_DIS", "0")
paPath(17) = Read_Ini_File(pRuta & "\PathV2.ini", "TIMES", "TM_BON_VID", "20")
paPath(53) = Read_Ini_File(pRuta & "\PathV2.ini", "TIMES", "ID_CNT_PRO", "0")
'--------------------------------------------------------------------------------------
igFlg_SavedCR = Val(Read_Ini_File(pRuta & "\PathV2.ini", "CREDITS", "FLG_SAVECR", "0"))
paPath(20) = Read_Ini_File(pRuta & "\PathV2.ini", "CREDITS", "LIMCR_CRED", "80")
paPath(21) = Read_Ini_File(pRuta & "\PathV2.ini", "CREDITS", "KEEP_SCRED", "0")
paPath(38) = Read_Ini_File(pRuta & "\PathV2.ini", "CREDITS", "BONUS_CRED", "0")
paPath(36) = Read_Ini_File(pRuta & "\PathV2.ini", "CREDITS", "VIDEO_CRED", "0")
'--------------------------------------------------------------------------------------
paPath(22) = Read_Ini_File(pRuta & "\PathV2.ini", "THEMES", "POPULAR_MIXER", "0")
'--------------------------------------------------------------------------------------
paPath(23) = Read_Ini_File(pRuta & "\PathV2.ini", "KEYBOARD", "KB_ADD_01", "+")
paPath(25) = Read_Ini_File(pRuta & "\PathV2.ini", "KEYBOARD", "KB_ADD_03", "N")
paPath(27) = Read_Ini_File(pRuta & "\PathV2.ini", "KEYBOARD", "KB_DEL_01", "D")
paPath(29) = Read_Ini_File(pRuta & "\PathV2.ini", "KEYBOARD", "KB_RET_01", ".")
paPath(31) = Read_Ini_File(pRuta & "\PathV2.ini", "KEYBOARD", "KB_RST_01", "S")
paPath(33) = Read_Ini_File(pRuta & "\PathV2.ini", "KEYBOARD", "KB_RST_03", "R")
paPath(35) = Read_Ini_File(pRuta & "\PathV2.ini", "KEYBOARD", "KB_POP_01", "P")
paPath(37) = Read_Ini_File(pRuta & "\PathV2.ini", "KEYBOARD", "KB_VIP_01", "V")
paPath(39) = Read_Ini_File(pRuta & "\PathV2.ini", "KEYBOARD", "KB_PUP_01", "/")
paPath(41) = Read_Ini_File(pRuta & "\PathV2.ini", "KEYBOARD", "KB_PDN_01", "*")
paPath(40) = Read_Ini_File(pRuta & "\PathV2.ini", "KEYBOARD", "KB_VERIFY", "H")
paPath(51) = Read_Ini_File(pRuta & "\PathV2.ini", "KEYBOARD", "KB_SWTPUB", "W")
sgKb_Pause = Read_Ini_File(pRuta & "\PathV2.ini", "KEYBOARD", "KB__PAUSE", "L")
sgKb_SwK = Read_Ini_File(pRuta & "\PathV2.ini", "KEYBOARD", "KB_SWTKAR", "K")
End Sub

Public Function Scramble(Text As String) As String
Dim i As Integer
Dim c As Integer
Dim Temp As String
Temp = ""
For i = 1 To Len(Text)
    c = Asc(Mid(Text, i, 1))
    c = c + 10
         
    'If the character was near the end of the ASCII character set then
    'adding 10 to the ASCII value can tip it over 255 which is the last
    'character - in this case, find the difference and that's the new
    'character (ie. at the start of the character set
    If c > 255 Then c = c - 255
        Temp = Temp & Chr(c)
    Next i
    Scramble = Temp
End Function

Public Function UnScramble(Text As String) As String
Dim i As Integer
Dim c As Integer
Dim Temp As String
Temp = ""
For i = 1 To Len(Text)
    c = Asc(Mid(Text, i, 1))
    c = c - 10
    'If the character was near the start of the ASCII character set then
    'subtracting 10 from the ASCII value can take it under 0 which is the
    'first character - in this case, find the difference and that's the
    'new character (ie. at the start of the character set
    'You will notice that we add rather than subtract from 255 because if
    'you minus a minus it becomes a plus (!)
    If c < 0 Then c = 256 + c
        Temp = Temp & Chr(c)
    Next i
    UnScramble = Temp
End Function

Public Function FileExist(ByVal strPathName As String) As Boolean
Dim intFileNum As Integer
On Error Resume Next
'
'Remove any trailing directory separator character
'
If Right$(strPathName, 1) = "\" Then
    strPathName = Left$(strPathName, Len(strPathName) - 1)
End If
'
'Attempt to open the file, return value of this function is False
'if an error occurs on open, True otherwise
'
intFileNum = FreeFile
Open strPathName For Input As intFileNum
FileExist = IIf(Err, False, True)
Close intFileNum
Err = 0
End Function

' ================================================================================
' Routine:              GetTheComputerName
' Description:
' Algorithm:
' Parameters:           None
' Returns:
' ================================================================================
Public Function GetTheComputerName() As String
    On Error GoTo errorhandler
    Dim strComputerName As String ' Variable to return the path of computer name
    strComputerName = Space(250) ' Initilize the buffer to receive the string
    GetComputerName strComputerName, Len(strComputerName)
    strComputerName = Mid(Trim$(strComputerName), 1, Len(Trim$(strComputerName)) - 1)
    GetTheComputerName = strComputerName
    Exit Function
 
errorhandler:
    Err.Raise Err.Number, Err.Source & "/Utils.GetTheComputerName", Err.Description
End Function

' ================================================================================
' Routine:              GetTheWindowsDirectory
' Description:
' Algorithm:
' Parameters:           None
' Returns:
' ================================================================================
Public Function GetTheWindowsDirectory() As String
    On Error GoTo errorhandler
    Dim strWindowsDir As String        ' Variable to return the path of Windows Directory
    Dim lngWindowsDirLength As Long    ' Variable to return the the lenght of the path
    strWindowsDir = Space(250)         ' Initilize the buffer to receive the string
    lngWindowsDirLength = GetWindowsDirectory(strWindowsDir, 250) ' Read the path of the windows directory
    strWindowsDir = Left(strWindowsDir, lngWindowsDirLength) ' Extract the windows path from the buffer
    GetTheWindowsDirectory = strWindowsDir
    Exit Function
 
errorhandler:
    Err.Raise Err.Number, Err.Source & "/Utils.GetTheWindowsDirectory", Err.Description
End Function

Public Function CheckValue(ByVal s As Variant) As String
   If VarType(s) = vbNull Then s = ""
   s = Trim(s)
   If s = "" Then CheckValue = "Not available" Else CheckValue = s
End Function

Public Function ObPlayer_Ocupado(oPlayer As Object) As Boolean
If (oPlayer.playState = wmppsWaiting Or _
    oPlayer.playState = wmppsPlaying Or _
    oPlayer.playState = wmppsBuffering Or _
    oPlayer.playState = wmppsTransitioning) Then
    ObPlayer_Ocupado = True
Else
    ObPlayer_Ocupado = False
End If
End Function

Public Sub pMPlayer_Change()
If bgIs_Video = True Then
    If ObPlayer_Ocupado(Video_Form.MediaPlayer3) = True Then
        Exit Sub
    End If
Else
    If ObPlayer_Ocupado(Main_Form.MediaPlayer1) = True Then
        Exit Sub
    End If
End If
If Main_Form.oLst_A_Tocar.ListCount > 0 Then
    Main_Form.oLst_A_Tocar.RemoveItem (0)
    Main_Form.oLst_A_Tocar.Refresh
End If
End Sub

Public Sub Muestra_Tema_Det()
Dim aDet() As String
If Main_Form.oLst_A_Tocar.List(0) = "" Then
    Main_Form.otTema_Act.Text = ""
    Main_Form.olTema_Act.Caption = ""
    Exit Sub
End If
aDet = VBA.Split(Main_Form.oLst_A_Tocar.List(0), ",", , vbTextCompare)
Main_Form.otTema_Act.Text = aDet(0)
If VBA.Trim(aDet(4)) = "*" Then
    Main_Form.olCred_Msg.Tag = Main_Form.olCred_Msg.Caption
    Main_Form.olCred_Msg.Caption = "DISCO PROMO..."
    Main_Form.Image2.Picture = LoadPicture(aDet(5))
Else
    If Main_Form.olCred_Msg.Caption <> "INSERTE ¢ 0.25" Then
        Main_Form.olCred_Msg.Caption = "INSERTE ¢ 0.25"
    End If
    Call Refresh_Creditos(Main_Form)
End If
Main_Form.olTema_Act.Caption = VBA.Trim(aDet(2))
End Sub

Public Sub Remove_Temes()
If Main_Form.oLst_A_Tocar.ListCount > 0 Then
    Main_Form.oLst_A_Tocar.RemoveItem (0)
    Main_Form.oLst_A_Tocar.Refresh
    Call Muestra_Tema_Det
End If
End Sub

Public Sub Sincronoze_Media()
Main_Form.MediaPlayer2.Controls.currentPosition = Video_Form.MediaPlayer3.Controls.currentPosition
Main_Form.MediaPlayer2.Controls.play
End Sub

Public Function Refresh_Creditos(oForm As Object)
If igKeep_Cred = 1 Then
    oForm.olCreditos.Caption = "CRÉDITOS GRATIS"
    oForm.oTimer_Moneda.Enabled = False
    oForm.olCred_Msg.Visible = False
    Exit Function
End If
If igCnt_CR > 0 Then
    oForm.oTimer_Moneda.Enabled = False
    oForm.olCred_Msg.Visible = False
Else
    oForm.oTimer_Moneda.Enabled = True
    oForm.olCred_Msg.Visible = True
End If
oForm.olCreditos.Caption = "CREDITOS (" + VBA.Trim(VBA.Str(igCnt_CR)) & ")"
'If oForm.olMetros.Visible = False Then
'    oForm.olLabMetros.Visible = True
'    oForm.olMetros.Visible = True
'End If
oForm.olMetros.Caption = PADL(igCnt_CRG, 6, "0")
End Function

Public Sub ControlError()
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Dim oFs As Object
Dim oTx As Object
Dim sCa As String
Dim sFe As String
Main_Form.olMessage.Visible = True
Main_Form.olMessage.Caption = "ERROR:" & VBA.Trim(Err.Description)
Main_Form.oTime_Mensajes.Enabled = True
sCa = "Error: " & VBA.Format(VBA.Date, "dd-mm-yyyy") & " - " & _
    VBA.Format(VBA.Time(), "hh:mm:ss AMPM") & " - " & _
    VBA.Str(Err.Number) & " - " & _
    PADR(VBA.Trim(Err.Description), 25, " ") & " - " & _
    VBA.Trim(Err.LastDLLError) & " - " & _
    PADR(VBA.Trim(Err.Source), 25, " ")
    
sFe = App.Path & "\ERROR.TXT"
Set oFs = CreateObject("Scripting.FileSystemObject")
If oFs.FileExists(sFe) = False Then
    Set oTx = oFs.CreateTextFile(sFe, True)
Else
    Set oTx = oFs.OpenTextFile(sFe, ForAppending, -2)
End If
oTx.WriteLine (sCa)
oTx.Close
End Sub

Public Sub Sleep(Seconds)
Dim PauseTime, Start, Finish
    PauseTime = Seconds   ' Set duration.
    Start = Timer   ' Set start time.
    Do While Timer < Start + Seconds
        'DoEvents    ' Yield to other processes.
    Loop
    Finish = Timer  ' Set end time.
End Sub

Public Function PADL(vIn_Val As Variant, iIn_Len As Integer, sIn_Simbol As String)
'Esta función; es genérica para todos los programas y formularios que la soliciten y se encarga
'de rellenar los espacios a la derecha la cantidad de veces que se decee, con el caracter que se desee
'Sintaxis: PADR(20, 6, "0") el resultado es "20000000"
'Diseñado por Jhonn G. Gutiérrez A.
'Última modificasión: 03 Febrero del 2002
Dim sCadena_Val As String
Dim iCadena_Len As Integer
If IsEmpty(vIn_Val) Then
    PADL = ""
    Exit Function
End If
If IsEmpty(sIn_Simbol) Then
    sIn_Simbol = " "
End If
If IsNumeric(vIn_Val) Then
    sCadena_Val = VBA.Trim(VBA.Str(vIn_Val))
Else
    If VBA.VarType(vIn_Val) = vbVariant Or vbString Then
        sCadena_Val = VBA.Trim(vIn_Val)
    Else
        MsgBox "Tipo de dato desconocido, o no manejable por esta función", vbOKOnly, "Error"
        PADL = ""
        Exit Function
    End If
End If
iCadena_Len = Len(sCadena_Val)
If iCadena_Len > iIn_Len Then
    sCadena_Val = VBA.Left(sCadena_Val, iIn_Len)
Else
    sCadena_Val = VBA.String((iIn_Len - iCadena_Len), sIn_Simbol) + sCadena_Val
End If
PADL = sCadena_Val
End Function

Public Function PADR(vIn_Val As Variant, iIn_Len As Integer, sIn_Simbol As String)
'Esta función; es genérica para todos los programas y formularios que la soliciten y se encarga
'de rellenar los espacios a la izquierda la cantidad de veces que se decee, con el caracter que se desee
'Sintaxis: PADR(20, 6, "0") el resultado es "00000020"
'Diseñado por Jhonn G. Gutiérrez A.
'Última modificasión: 03 Febrero del 2002
Dim sCadena_Val As String
Dim iCadena_Len As Integer
If IsEmpty(vIn_Val) Then
    PADR = ""
    Exit Function
End If
If IsEmpty(sIn_Simbol) Then
    sIn_Simbol = " "
End If
If IsNumeric(vIn_Val) Then
    sCadena_Val = VBA.Trim(VBA.Str(vIn_Val))
Else
    If VBA.VarType(vIn_Val) = vbVariant Or vbString Then
        sCadena_Val = VBA.Trim(vIn_Val)
    Else
        MsgBox "Tipo de dato desconocido, o no manejable por esta función", vbOKOnly, "Error"
        PADR = ""
        Exit Function
    End If
End If
iCadena_Len = Len(sCadena_Val)
If iCadena_Len > iIn_Len Then
    sCadena_Val = VBA.Left(sCadena_Val, iIn_Len)
Else
    sCadena_Val = sCadena_Val + VBA.String((iIn_Len - iCadena_Len), sIn_Simbol)
End If
PADR = sCadena_Val
End Function

Function Proper(ByVal TXT As String) As String
Dim need_cap As Boolean
Dim i As Integer
Dim ch As String
    TXT = VBA.LCase(TXT)
    need_cap = True
    For i = 1 To Len(TXT)
        ch = Mid$(TXT, i, 1)
        If ch >= "a" And ch <= "z" Then
            If need_cap Then
                Mid$(TXT, i, 1) = VBA.UCase(ch)
                need_cap = False
            End If
        Else
            need_cap = True
        End If
    Next i
    Proper = TXT
End Function

Public Function GenerateRandomFileName() As String
Const MASKNUM As String = "_0123456789"
Const MASKCHR As String = "abcdefghijklmnoprstuvwxyz"
Const MASK As String = MASKCHR + MASKNUM
Const MINLEN As Integer = 4
Const MAXLEN As Integer = 12

Dim nMask As Long
Dim nFile As Long
Dim sFile As String
Dim sExt As String
Dim i As Long
Dim nChr As Long

nFile = MINLEN + (MAXLEN - MINLEN) * Rnd()
nMask = Len(MASK)
For i = 1 To nFile
nChr = Int(nMask * Rnd()) + 1
sFile = sFile + Mid$(MASK, nChr, 1)
Next
nMask = Len(MASKCHR)
For i = 1 To 3
nChr = Int(nMask * Rnd()) + 1
sExt = sExt + Mid$(MASKCHR, nChr, 1)
Next

GenerateRandomFileName = sFile + "." + sExt
End Function

Public Sub ControlPanels(filename As String)
Dim rtn As Double
On Error Resume Next
rtn = Shell(filename, 5)
End Sub

Public Function GetWindowsDir() As String
    Dim sRet As String, lngRet As Long
    sRet = String$(MAX_PATH, 0)
    lngRet = GetWindowsDirectory(sRet, MAX_PATH)
    GetWindowsDir = Left(sRet, lngRet)
End Function

Public Sub sCenterForm(tmpF As Form)
Dim X As Integer, Y As Integer
Y = (Screen.Height - tmpF.Height) \ 2
X = (Screen.Width - tmpF.Width) \ 2
tmpF.Move X, Y
End Sub


Public Function MBCPUNumber() As String
Dim oWMI As Object
Dim oCpu As Object
Dim sCpuId As String
sCpuId = ""
Set oWMI = CreateObject("winmgmts:")
For Each oCpu In oWMI.instancesof("Win32_Processor")
    sCpuId = sCpuId & oCpu.ProcessorID
Next
MBCPUNumber = sCpuId
Set oWMI = Nothing
Set oCpu = Nothing
End Function

Function VBRegSvr32(ByVal sServerPath As String, _
                                          Optional fRegister = True) As Boolean
Dim hMod As Long            ' module handle
Dim lpfn As Long            ' reg/unreg function address
Dim lpThreadId As Long      ' dummy var that get's filled
Dim hThread As Long         ' thread handle
Dim fSuccess As Boolean     ' if things worked
Dim dwExitCode As Long      ' thread's exit code if it doesn't finish
Dim WAIT_OBJECT_0   As Long
WAIT_OBJECT_0 = 0
   Screen.MousePointer = vbHourglass

   ' Load the server into memeory
   hMod = LoadLibrary(sServerPath)
       
   ' Get the specified function's address and our msgbox string.
   If fRegister Then
       lpfn = GetProcAddress(hMod, "DllRegisterServer")
   Else
       lpfn = GetProcAddress(hMod, "DllUnregisterServer")
   End If
     
   ' If we got a function address...
   If lpfn Then
       ' Create an alive thread and execute the function.
       hThread = CreateThread(ByVal 0, 0, ByVal lpfn, ByVal 0, 0, lpThreadId)
         
       ' If we got the thread handle...
       If hThread Then
           ' Wait 10 secs for the thread to finish (the function may take a while...)
           fSuccess = (WaitForSingleObject(hThread, 10000) = WAIT_OBJECT_0)
           
           ' If it didn't finish in 5 seconds...
           If Not fSuccess Then
               ' Something unlikely happened, lose the thread.
               Call GetExitCodeThread(hThread, dwExitCode)
               Call ExitThread(dwExitCode)
           End If
     
       ' Lose the thread handle
       Call CloseHandle(hThread)
       End If   ' hThread
   End If   ' lpfn
     
   ' Free server if we loaded it.
   If hMod Then Call FreeLibrary(hMod)
     
   Screen.MousePointer = vbDefault
     
   If fSuccess Then
       VBRegSvr32 = True
   Else
       MsgBox ("Error: " & sServerPath)
       VBRegSvr32 = False
   End If
   
End Function

Public Function SetCursorPosition(Window As Object, xPos As Long, yPos As Long) As Boolean



'AUTHOR: FreeVBCode.com (http://www.freevbcode.com)

'Usage: Window = window for which you want
'to set the cursof for (form or control)

'window must support hwnd property; function will
'return false if it doesn't (e.g., won't work
'with labels)

'x = xPosition in twips
'y = yPosition in twips

'if you are not using twips as your scale mode
'you must convert the value to pixels from your
'metric. There is a function on FreeVBCode.com
'that does this, under the System/API
'category

On Error GoTo errorhandler
Dim X As Long, Y As Long
Dim lRet As Long
Dim lHandle As Long
Dim typPoint As POINTAPI

lHandle = Window.hWnd
With Screen
    X = CLng(xPos / .TwipsPerPixelX)
    Y = CLng(yPos / .TwipsPerPixelY)
End With
 
typPoint.X = X
typPoint.Y = Y

lRet = ClientToScreen(lHandle, typPoint)
lRet = SetCursorPos(typPoint.X, typPoint.Y)
SetCursorPosition = (lRet <> 0)
Exit Function

errorhandler:

SetCursorPosition = False
Exit Function

End Function


Public Sub Desordenar_array(ByRef vArray As Variant, _
                         startIndex As Variant, _
                         endIndex As Variant)
   
    Dim i As Long
    Dim rndIndex As Long
    Dim Temp As Variant
    
    Randomize
    
    startIndex = LBound(vArray)
    endIndex = UBound(vArray)
    
    For i = startIndex To endIndex
        rndIndex = Int((endIndex - startIndex + 1) * Rnd() + startIndex)

        Temp = vArray(i)
        vArray(i) = vArray(rndIndex)
        vArray(rndIndex) = Temp
    Next i
End Sub


