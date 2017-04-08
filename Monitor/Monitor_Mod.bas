Attribute VB_Name = "Main"
Option Explicit

'Declaraciones para 32 bits
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
        (ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
         ByVal lpDefault As String, ByVal lpReturnedString As String, _
         ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
        (ByVal lpApplicationName As String, ByVal lpKeyName As Any, _
         ByVal lpString As Any, ByVal lpFileName As String) As Long


'----------------------------------------------------------------------------
'Funci�n equivalente a GetSetting de VB4.
'GetSetting     En VB4/32bits usa el registro.
'               En VB4/16bits usa un archivo de texto.
'Pero al usar las llamadas del API, siempre se escriben en archivos de texto.
'----------------------------------------------------------------------------
Public Function Read_Ini_File(lpFileName As String, lpAppName As String, lpKeyName As String, Optional vDefault) As String
    'Los par�metros son:
    'lpFileName:    La Aplicaci�n (fichero INI)
    'lpAppName:     La secci�n que suele estar entrre corchetes
    'lpKeyName:     Clave
    'vDefault:      Valor opcional que devolver�
    '               si no se encuentra la clave.
    '
    Dim lpString As String
    Dim LTmp As Long
    Dim sRetVal As String
    
    'Si no se especifica el valor por defecto,
    'asignar incialmente una cadena vac�a
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
    'Guarda los datos de configuraci�n
    'Los par�metros son los mismos que en LeerIni
    'Siendo lpString el valor a guardar
    '
    Dim LTmp As Long

    LTmp = WritePrivateProfileString(lpAppName, lpKeyName, lpString, lpFileName)
End Sub

