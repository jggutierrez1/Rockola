Option Explicit
Public Name As String
Private mvarImporteADesglosar As Double
Private mvarMonedas As String
Private mvarDesglose As String
Private TMonedas() As Double
Private TDesglose() As Long
Private Nmonedas As Integer
'  Declaración de Constantes
Private Const UDL_LINEA1 = "[oledb]"
Private Const UDL_LINEA2 = "; Everything after this line is an OLE DB initstring"
Private Const KEY_ALL_CLASSES As Long = &HF0063
Private Const REG_SZ = 1
Private Const ERROR_SUCCESS = 0&
Private Const MAX_FILENAME_LEN = 256
Private Const RESOURCETYPE_DISK = &H1
Private Const MAX_COMPUTERNAME_LENGTH = 255
Private Const MAX_PREFERRED_LENGTH        As Long = -1
Private Const ERROR_MORE_DATA             As Long = 234&
Private Const SV_TYPE_ALL                 As Long = &HFFFFFFFF
'Configuración Regional
Private Const LOCALE_SDECIMAL = &HE         'separador de decimales NUMERO
Private Const LOCALE_STHOUSAND = &HF        'separador de miles NUMERO
Private Const LOCALE_SDATE = &H1D        'separador de fecha
Private Const LOCALE_STIME = &H1E        'separador de hora
Private Const LOCALE_SSHORTDATE = &H1F        'cadena de formato de fecha corta
Private Const LOCALE_SLONGDATE = &H20         'cadena de formato de fecha larga
Private Const LOCALE_STIMEFORMAT = &H1003     'cadena de formato de hora
Private Const LOCALE_USER_DEFAULT = &H400
'  Declaración de tipos de datos
Private Type NETRESOURCE
   dwScope As Long
   dwType As Long
   dwDisplayType As Long
   dwUsage As Long
   lpLocalName As String
   lpRemoteName As String
   lpComment As String
   lpProvider As String
End Type
Private Type SERVER_INFO_100
  sv100_platform_id As Long
  sv100_name As Long
End Type
'  Declaración de Funciones API
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pTo As Any, uFrom As Any, ByVal lSize As Long)
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetVolumeInformation& Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal pVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long)
Private Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long          ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function WNetConnectionDialog Lib "mpr.dll" (ByVal hwnd As Long, ByVal dwType As Long) As Long
Private Declare Function WNetDisconnectDialog Lib "mpr.dll" (ByVal hwnd As Long, ByVal dwType As Long) As Long
Private Declare Function WNetAddConnection2 Lib "mpr.dll" Alias "WNetAddConnection2A" (lpNetResource As NETRESOURCE, ByVal lpPassword As String, ByVal lpUserName As String, ByVal dwFlags As Long) As Long
Private Declare Function WNetCancelConnection2 Lib "mpr.dll" Alias "WNetCancelConnection2A" (ByVal lpName As String, ByVal dwFlags As Long, ByVal fForce As Long) As Long
Private Declare Function NetServerEnum Lib "Netapi32" (ByVal servername As Long, ByVal level As Long, buf As Any, ByVal prefmaxlen As Long, entriesread As Long, totalentries As Long, ByVal servertype As Long, ByVal domain As Long, resume_handle As Long) As Long
Private Declare Function NetApiBufferFree Lib "netapi32.dll" (ByVal Buffer As Long) As Long
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
'para la siguiente, activar la Microsoft Scriptin Runtime
Dim ObjetoSistemaFicheros As New Scripting.FileSystemObject
Public Sub ModoNumerico(ByVal hWndTextBox As Long)
    SetWindowLong hWndTextBox, -16, GetWindowLong(hWndTextBox, -16) Or 8192
End Sub
Public Sub ModoMayusculas(ByVal hWndTextBox As Long)
    SetWindowLong hWndTextBox, -16, GetWindowLong(hWndTextBox, -16) Or 8
End Sub
Public Sub ModoMinusculas(ByVal hWndTextBox As Long)
    SetWindowLong hWndTextBox, -16, GetWindowLong(hWndTextBox, -16) Or &HA
End Sub
Public Function CrearDirectorio(Sendero As String, CrearControl As Long, Mensaje As Boolean) As Boolean
    On Error Resume Next
    If Len(Trim(Sendero)) > 0 Then
        Err.Clear
        ObjetoSistemaFicheros.CreateFolder Trim(Sendero)
        If Err.Number <> 0 Then
            If Mensaje Then MsgBox "Imposible Crear " & Trim(Sendero) & vbCrLf & Err.Description
            CrearDirectorio = Err.Number
        Else
            CrearDirectorio = 0
            If CrearControl Then ObjetoSistemaFicheros.CreateTextFile Trim(Sendero) & "\JdmCreate.dir"
            Err.Clear
        End If
    End If
End Function
Public Function TeclasAlfaNumericas(TeclaPulsada As Integer) As Integer
   Select Case TeclaPulsada
          Case 65 To 90, 97 To 122  'Ascii de A a Z y a a z
          Case 209, 241             'Ascii de Ñ y ñ
          Case 48 To 57             'Ascii de 0 a 9
          Case 8                    'Ascii de BackSpace
          Case 32                   'Ascii de Space
          Case 13                   'enter
          Case Else
               TeclaPulsada = 0         'Cualquier otra tecla es anulada
   End Select
   TeclasAlfaNumericas = TeclaPulsada
End Function
Public Function TeclasAlfa(TeclaPulsada As Integer) As Integer
   Select Case TeclaPulsada
          Case 65 To 90, 97 To 122  'Ascii de A a Z y a a z
          Case 209, 241             'Ascii de Ñ y ñ
          Case 8                    'Ascii de BackSpace
          Case 32                   'Ascii de Space
          Case 13                   'enter
          Case Else
               TeclaPulsada = 0         'Cualquier otra tecla es anulada
   End Select
   TeclasAlfa = TeclaPulsada
End Function
Public Function TeclasNumericas(TeclaPulsada As Integer) As Integer
   Select Case TeclaPulsada
          Case 48 To 57             'Ascii de 0 a 9
          Case 8                    'Ascii de BackSpace
          Case 13                   'enter
          Case Else
               TeclaPulsada = 0         'Cualquier otra tecla es anulada
   End Select
   TeclasNumericas = TeclaPulsada
End Function
'Iconos e imagenes
Public Sub CargaIcono(Destino As Object, Icono As String)
   On Error Resume Next
   If Len(Icono) > 0 Then
      Open Icono For Input As #1
      If Err.Number = 0 Then
         Close #1
         Destino.Icon = LoadPicture("")
         Destino.Icon = LoadPicture(Icono)
      Else
         Destino.Icon = LoadPicture("")
         Destino.Icon = LoadPicture("Aplicación.Ico")
      End If
   Else
        Destino.Icon = LoadPicture("")
        Destino.Icon = LoadPicture("Aplicación.Ico")
   End If
   Err.Clear
End Sub
Public Sub CargaImage(Destino As Object, Fichero As String)
   On Error Resume Next
   Open Fichero For Input As #1
   If Err.Number = 0 Then
      Close #1
      Destino.Picture = LoadPicture("")
      Destino.Picture = LoadPicture(Fichero)
   End If
   Err.Clear
End Sub
Public Sub CargaPictureBox(Destino As Object, Fichero As String)
   On Error Resume Next
   Open Fichero For Input As #1
   If Err.Number = 0 Then
      Close #1
      Destino.Picture = LoadPicture("")
      Destino.Picture = LoadPicture(Fichero)
   End If
   Err.Clear
End Sub

Public Function DiasNaturales(FechaIni As Date, FechaFin As Date) As Integer
Dim I As Long, FechaBucle As Date, numD As Long
    DiasNaturales = 0
    FechaBucle = FechaIni
    numD = DateDiff("d", FechaIni, FechaFin, vbUseSystemDayOfWeek, vbUseSystemDayOfWeek)
    If FechaIni < FechaFin Then
    For I = 1 To DateDiff("d", FechaIni, FechaFin, vbUseSystemDayOfWeek, vbUseSystemDayOfWeek)
        FechaBucle = Format(FechaIni + I, "Short date")
        If Weekday(FechaBucle, vbUseSystemDayOfWeek) >= 1 And Weekday(FechaBucle, vbUseSystemDayOfWeek) <= 5 Then
            DiasNaturales = DiasNaturales + 1
        End If
    Next
    ElseIf FechaIni > FechaFin Then
        For I = 1 To DateDiff("d", FechaFin, FechaIni, vbUseSystemDayOfWeek, vbUseSystemDayOfWeek)
            FechaBucle = Format(FechaIni - I, "Short date")
            If Weekday(FechaBucle, vbUseSystemDayOfWeek) >= 1 And Weekday(FechaBucle, vbUseSystemDayOfWeek) <= 5 Then
                DiasNaturales = DiasNaturales + 1
            End If
        Next
    End If
End Function
Public Function DiaSemana(Fecha As Date) As String
    Select Case Weekday(Fecha)
        Case 1 'Domingo
            DiaSemana = "Domingo"
        Case 2 'Lunes
            DiaSemana = "Lunes"
        Case 3 'Martes
            DiaSemana = "Martes"
        Case 4 'Miercoles
            DiaSemana = "Miercoles"
        Case 5 'Jueves
            DiaSemana = "Jueves"
        Case 6 'Viernes
            DiaSemana = "Viernes"
        Case 7 'Sabado
            DiaSemana = "Sabado"
        Case Else
            DiaSemana = ""
    End Select
End Function
Public Function NombreMes(ByVal Entrada As Integer) As String
    Select Case Entrada
        Case 1
            NombreMes = "Enero"
        Case 2
            NombreMes = "Febrero"
        Case 3
            NombreMes = "Marzo"
        Case 4
            NombreMes = "Abril"
        Case 5
            NombreMes = "Mayo"
        Case 6
            NombreMes = "Junio"
        Case 7
            NombreMes = "Julio"
        Case 8
            NombreMes = "Agosto"
        Case 9
            NombreMes = "Septiembre"
        Case 10
            NombreMes = "Octubre"
        Case 11
            NombreMes = "Noviembre"
        Case 12
            NombreMes = "Diciembre"
        Case Else
            NombreMes = "INVALIDO"
    End Select
End Function
Public Function NumeroMes(ByVal Entrada As String) As Integer
Dim I As Integer
    NumeroMes = 0
    For I = 1 To 12
        If (StrComp(Trim(Entrada), NombreMes(I), vbTextCompare)) Then
            NumeroMes = I
            Exit Function
        End If
    Next I
End Function
Public Function DiasEnMes(Fecha As Date) As Byte
    If Not IsDate(Fecha) Then
        DiasEnMes = 0
    Else
        Select Case Month(Fecha)
        Case 2
            If Bisiesto(Year(Fecha)) Then
                DiasEnMes = 29
            Else
                DiasEnMes = 28
            End If
        Case 4, 6, 9, 11
            DiasEnMes = 30
        Case 1, 3, 5, 7, 8, 10, 12
            DiasEnMes = 31
        End Select
    End If
End Function
Private Function Bisiesto(YYYY As Integer) As Integer
    Bisiesto = YYYY Mod 4 = 0 And (YYYY Mod 100 <> 0 Or YYYY Mod 400 = 0)
End Function
Public Function IsTime(ByVal Entrada As String) As Boolean
Dim Largo As Integer
    Largo = Len(Entrada)
    Select Case Largo
        Case 5
            If Mid(Entrada, 3, 1) = ":" Then
                If CInt(Mid(Entrada, 1, 2)) < 0 Or CInt(Mid(Entrada, 1, 2)) > 23 Then
                    IsTime = False
                Else
                    If CInt(Mid(Entrada, 4, 2)) < 0 Or CInt(Mid(Entrada, 4, 2)) > 59 Then
                        IsTime = False
                    Else
                        IsTime = True
                    End If
                End If
            Else
                IsTime = False
            End If
        Case 8
            If Mid(Entrada, 3, 1) = ":" Then
                If CInt(Mid(Entrada, 1, 2)) < 0 Or CInt(Mid(Entrada, 1, 2)) > 23 Then
                    IsTime = False
                Else
                    If CInt(Mid(Entrada, 4, 2)) < 0 Or CInt(Mid(Entrada, 4, 2)) > 59 Then
                        IsTime = False
                    Else
                        IsTime = True
                    End If
                End If
            Else
                IsTime = False
            End If
        Case Else
            IsTime = False
    End Select
End Function
Public Function QuitaHoras(ByVal Entrada As Date) As Date
    If CLng(Entrada) > CDbl(Entrada) Then
        QuitaHoras = CDate(CLng(Entrada - 1))
    Else
        QuitaHoras = CDate(CLng(Entrada))
    End If
End Function
Public Function LimpiaHora(ByVal Entrada As String) As String
Dim Auxiliar As String
    Auxiliar = Trim(Entrada)
    If IsNull(Auxiliar) Or Len(Auxiliar) = 0 Then
        LimpiaHora = "00:00:00"
    Else
        Select Case (Len(Auxiliar))
            Case 1
                LimpiaHora = Auxiliar & "0:00:00"
            Case 2
                LimpiaHora = Auxiliar & ":00:00"
            Case 3
                LimpiaHora = Auxiliar & "00:00"
            Case 4
                LimpiaHora = Auxiliar & "0:00"
            Case 5
                LimpiaHora = Auxiliar & ":00"
            Case 6
                LimpiaHora = Auxiliar & "00"
            Case 7
                LimpiaHora = "0" & Auxiliar
            Case 8
                LimpiaHora = Auxiliar
            Case Else
                LimpiaHora = Mid(Auxiliar, 1, 8)
        End Select
    End If
End Function
Public Function VerificaFecha(ByVal Entrada As String) As String
Dim Auxiliar As String
    Auxiliar = Trim(Entrada)
    If IsNull(Auxiliar) Or Len(Auxiliar) = 0 Then
        Auxiliar = Format(QuitaHoras(Now()), "Short Date")
    Else
        Auxiliar = Format(Auxiliar, "Short Date")
    End If
    If Len(Auxiliar) = 7 Then
        VerificaFecha = "0" & Auxiliar
    Else
        VerificaFecha = Auxiliar
    End If
End Function
Public Function Horas(ByVal Entrada As Date) As Long
    Horas = Hour(Entrada) + Round(Minute(Entrada) / 60, 0) + Round(Second(Entrada) / 3600, 0)
End Function
Public Function Minutos(ByVal Entrada As Date) As Long
    Minutos = (Hour(Entrada) * 60) + Minute(Entrada) + Round(Second(Entrada) / 60, 0)
End Function
Public Function Segundos(ByVal Entrada As Date) As Long
    Segundos = (Hour(Entrada) * 3600) + (Minute(Entrada) * 60) + Second(Entrada)
End Function
Public Function Edad(FechaNacimiento, FechaActual) As Integer
    If Month(FechaActual) < Month(FechaNacimiento) Or (Month(FechaActual) = Month(FechaNacimiento) And Day(FechaActual) < Day(FechaNacimiento)) Then
        Edad = Year(FechaActual) - Year(FechaNacimiento) - 1
    Else
        Edad = Year(FechaActual) - Year(FechaNacimiento)
    End If
End Function
Public Function SoloFecha(ByVal Entrada As Date) As String
    SoloFecha = Day(Entrada) & "/" & Month(Entrada) & "/" & Year(Entrada)
End Function
Public Function SoloHora(ByVal Entrada As Date) As Date
    SoloHora = Hour(Entrada) & ":" & Minute(Entrada) & ":" & Second(Entrada)
End Function
Public Function FechaSql(ByVal Entrada As Date) As String
    FechaSql = Year(Entrada) & "-" & Month(Entrada) & "-" & Day(Entrada)
End Function
Public Function FechaHoraSql(ByVal Entrada As Date) As String
    FechaHoraSql = Year(Entrada) & "-" & Month(Entrada) & "-" & Day(Entrada) & " " & Hour(Entrada) & ":" & Minute(Entrada) & ":" & Second(Entrada)
End Function
Public Function PreparaFechaSql(ByVal Entrada As Date) As String
    PreparaFechaSql = "CONVERT(DATETIME, '" & FechaHoraSql(Entrada) & "', 102)"
End Function
'Digitos de control
Public Function DcEan(Codigo) ' sistema ean, upc, itf
Dim dc As Integer, n As Integer
    dc = 0
    For n = Len(Codigo) To 1 Step -1
        dc = dc + (Val(Mid(Codigo, n, 1)) * 3)
        n = n - 1
        If n > 0 Then dc = dc + Val(Mid(Codigo, n, 1))
    Next
    DcEan = Right(Str(10 - (dc Mod 10)), 1)
End Function
Public Function Dc39(Codigo) 'code39
Const C39CHK = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ-. $/+%"
Dim n As Integer, suma As Long
    For n = 1 To Len(Codigo)
        suma = suma + InStr(C39CHK, Mid(Codigo, n, 1)) - 1
    Next
    Dc39 = Mid(C39CHK, (suma Mod 43) + 1, 1)
End Function
Public Function DcBancario(Cuenta As String)
    If Len(Cuenta) <> 18 Then
        DcBancario = ""
    Else
        DcBancario = DcBancario2(Mid(Cuenta, 1, 8)) & DcBancario2(Mid(Cuenta, 9))
    End If
End Function
Private Function DcBancario2(Serie As String)
Const PESOS = "06030709100508040201"
Dim l As Long, I As Integer
    For I = 1 To Len(Serie)
        l = l + (Mid(Serie, Len(Serie) - I + 1, 1) * Mid(PESOS, ((I - 1) * 2) + 1, 2))
    Next
    Select Case 11 - (l Mod 11)
    Case 11
        DcBancario2 = 0
    Case 10
        DcBancario2 = 1
    Case Else
        DcBancario2 = 11 - (l Mod 11)
    End Select
End Function
Public Function LetraNIF(ByVal lNIF As Long) As String
   LetraNIF = Mid("TRWAGMYFPDXBNJZSQVHLCKET", lNIF Mod 23 + 1, 1)
End Function
'Moneda
Public Function Euros(Pesetas As Double, Decimales As Integer) As Double
    Euros = Round(Pesetas / 166.386, Decimales)
End Function
Public Function Importe(Unidades As Double, Precio As Double) As Double
    If Unidades = 0 Or IsNull(Unidades) Or IsNull(Precio) Then
        Importe = 0
    Else
        Importe = Precio * Unidades
    End If
End Function
Public Function Pesetas(Euros As Double) As Double
    Pesetas = Round(Euros * 166.386, 0)
End Function
Public Function Precio(Unidades As Double, Importe As Double) As Double
    If Unidades = 0 Or IsNull(Unidades) Or IsNull(Importe) Then
        Precio = 0
    Else
        Precio = Importe / Unidades
    End If
End Function
Public Function Redondeando(ByVal numero As Double, ByVal Decimales As Integer) As Double
Dim Entero1 As Long, Entero2 As Long, Entero3 As Long, Resultado As Double
    Resultado = numero * (10 ^ Decimales)
    Entero1 = Int(Resultado)
    Resultado = numero * (10 ^ (Decimales + 1))
    Entero2 = Int(Resultado)
    Entero3 = Entero1 * 10
    If (Entero2 - Entero3) > 4 Then
        Redondeando = (Entero1 + 1) / (10 ^ Decimales)
    Else
        Redondeando = Entero1 / (10 ^ Decimales)
    End If
End Function
Public Function Unidades(Precio As Double, Importe As Double) As Double
    If Precio = 0 Or IsNull(Importe) Or IsNull(Precio) Then
        Unidades = 0
    Else
        Unidades = Importe / Precio
    End If
End Function
'Strings
Public Function Apellidos(ByVal A1, ByVal A2, ByVal NN) As String
    If IsNull(A1) Then A1 = "???"
    If IsNull(A2) Then A2 = "???"
    If IsNull(NN) Then NN = "???"
    Apellidos = NombrePropio(A1 & " " & A2 & "; " & NN)
End Function
Public Function ArreglaCRLF(ByVal Entrada As String) As String
Dim Temp As String, C As String, Salida As String, I As Integer
    Entrada = Trim(Entrada)
    If IsNull(Entrada) Or Len(Entrada) = 0 Then
        ArreglaCRLF = ""
    Else
        Temp = Trim(CStr(Entrada))
        Salida = ""
        For I = 1 To Len(Temp)
            C = Mid(Temp, I, 1)
            If C = vbCr Or C = vbLf Then
                Salida = Salida & " "
            Else
                Salida = Salida & C
            End If
        Next I
        ArreglaCRLF = QuitaDuplicados(Salida, " ")
    End If
End Function
Public Function ArreglaSignos(ByVal Entrada As String) As String
Dim C As String, I As Integer, Barrita As String * 1
    Barrita = "\"
    Entrada = Trim(Entrada)
    ArreglaSignos = ""
    If Not IsNull(Entrada) Then
        For I = 1 To Len(Entrada)
            C = Mid(Entrada, I, 1)
            Select Case C
                Case "§"
                    ArreglaSignos = ArreglaSignos & "º"
                Case "¥"
                    ArreglaSignos = ArreglaSignos & "Ñ"
                Case Barrita
                    ArreglaSignos = ArreglaSignos & "Ñ"
                Case "¤"
                    ArreglaSignos = ArreglaSignos & "ñ"
                Case "¦"
                    ArreglaSignos = ArreglaSignos & "ª"
                Case "@"
                    ArreglaSignos = ArreglaSignos & "ª"
                Case "{"
                    ArreglaSignos = ArreglaSignos & "ª"
                Case "š"
                    ArreglaSignos = ArreglaSignos & "ü"
                Case Else
                    ArreglaSignos = ArreglaSignos & C
            End Select
        Next I
    End If
End Function
Public Function MayusculaInicial(ByVal Entrada As String)
    Entrada = LCase(Trim(Entrada))
    MayusculaInicial = UCase(Left(Entrada, 1)) & Mid(Entrada, 2)
End Function
Public Function NombrePropio(Entrada As String) As String
Dim ptr As Integer
Dim Cadena As String
Dim CActual As String, CAnterior As String
    Entrada = Trim(Entrada)
    NombrePropio = ""
    If Not IsNull(Entrada) Or Len(Entrada) = 0 Then
        Cadena = QuitaDuplicados(CStr(Entrada), " ")
        For ptr = 1 To Len(Cadena)
            CActual = Mid(Cadena, ptr, 1)
            Select Case CAnterior
                Case "0" To "9", "A" To "Z", "a" To "z", "Ñ", "ñ", "Á", "á", "É", "é", "Í", "í", "Ó", "ó", "Ú", "ú", "Ü", "ü", "ç", "Ç"
                    Mid(Cadena, ptr, 1) = LCase(CActual)
                Case Else
                    Mid(Cadena, ptr, 1) = UCase(CActual)
            End Select
            CAnterior = CActual
        Next ptr
        NombrePropio = Cadena
    End If
End Function
Public Function PadL(ByVal Entrada As String, Largo As Integer, Caracter As String) As String
   If Len(Caracter) = 0 Then Caracter = " "
   Entrada = Trim(Entrada)
   PadL = Entrada
   If Len(Entrada) < Largo Then
      PadL = String((Largo - Len(Entrada)), Caracter) & PadL
   End If
End Function
Public Function PadR(ByVal Entrada As String, Largo As Integer, Caracter As String) As String
   If Len(Caracter) = 0 Then Caracter = " "
   Entrada = Trim(Entrada)
   PadR = Entrada
   If Len(Entrada) < Largo Then
      PadR = PadR & String((Largo - Len(Entrada)), Caracter)
   End If
End Function
Public Function RightAlign(ByVal Entrada As String, ByVal Largo As Long) As String
    If (Len(Entrada) > Largo) Then
        RightAlign = Right(Entrada, Largo)
    Else
        RightAlign = String(Largo - Len(Entrada), " ") & Entrada
    End If
End Function
Public Function QuitaBlancoComa(Entrada As String) As String
    Entrada = Trim(Entrada)
    QuitaBlancoComa = Entrada
    If IsNull(Entrada) Then
        QuitaBlancoComa = ""
    Else
        QuitaBlancoComa = Replace(Entrada, " ,", ",")
    End If
End Function
Public Function QuitaCaracter(Entrada As String, Caracter As String) As String
Dim C As String, I As Integer
    Entrada = Trim(Entrada)
    QuitaCaracter = ""
    Caracter = Mid(Trim(Caracter), 1, 1)
    For I = 1 To Len(Entrada)
        C = Mid(Entrada, I, 1)
        If Val(C) <> Val(Caracter) Then QuitaCaracter = QuitaCaracter & C
    Next I
End Function
Public Function QuitaDuplicados(Entrada As String, Caracter As String) As String
    Dim Doble As String
    Entrada = Trim(Entrada)
    Caracter = Mid(Trim(Caracter), 1, 1)
    Doble = Caracter & Caracter
    QuitaDuplicados = Entrada
    If Not IsNull(Entrada) Or Len(Entrada) > 0 Then
        If Len(Caracter) > 0 Then QuitaDuplicados = Replace(Entrada, Doble, Caracter)
    End If
End Function
Public Function QuitarSignos(Entrada As String)
Dim Temp As String, C As String, I As Integer
    QuitarSignos = ""
    If Not IsNull(Entrada) Then
        Temp = CStr(Trim(Entrada))
        For I = 1 To Len(Temp)
            C = Mid(Temp, I, 1)
            Select Case C
                Case "0" To "9", "A" To "Z", "a" To "z", "Ñ", "ñ", "Á", "á", "É", "é", "Í", "í", "Ó", "ó", "Ú", "ú", "Ü", "ü", "ç", "Ç", " "
                    QuitarSignos = QuitarSignos & C
            End Select
        Next I
    End If
End Function
Public Function SoloNumeros(Entrada As String) As String
Dim I As Integer, X As String
    SoloNumeros = ""
    If Not IsNull(Entrada) And Len(Entrada) > 0 And Not IsNull(Len(Entrada)) Then
        Entrada = Trim(Entrada)
        For I = 1 To Len(Entrada)
            X = Mid(Entrada, I, 1)
            If (X <= "9" And X >= "0") Then
            SoloNumeros = SoloNumeros & X
            End If
        Next I
    End If
End Function
Public Function SoloNumerico(Entrada As String) As String
Dim I As Integer, X As String, Signo As Boolean, Coma As Boolean, Punto As Boolean
    SoloNumerico = ""
    If Not IsNull(Entrada) And Len(Entrada) > 0 And Not IsNull(Len(Entrada)) Then
        Signo = True
        Coma = True
        Punto = True
        Entrada = Trim(Entrada)
        For I = 1 To Len(Entrada)
            X = Mid(Entrada, I, 1)
            Select Case X
                Case "0" To "9"
                    SoloNumerico = SoloNumerico & X
                Case "."
                    If Punto Then
                        Punto = False
                        SoloNumerico = SoloNumerico & X
                    End If
                Case ","
                    If Coma Then
                        Coma = False
                        SoloNumerico = SoloNumerico & X
                    End If
                Case "+"
                    If Signo And (I = 1 Or I = Len(Entrada)) Then
                        Signo = False
                        SoloNumerico = SoloNumerico & X
                    End If
                Case "-"
                    If Signo And (I = 1 Or I = Len(Entrada)) Then
                        Signo = False
                        SoloNumerico = SoloNumerico & X
                    End If
            End Select
        Next I
    End If
End Function
Public Function VerificaTextoNumerico(Entrada As String, Tipo As Byte, Signo As Boolean) As String
    VerificaTextoNumerico = ""
    Entrada = Trim(Entrada)
    If Not IsNull(Entrada) And Len(Entrada) > 0 And Not IsNull(Len(Entrada)) Then
        If Not Signo Then
            Entrada = Replace(Entrada, "+", "")
            Entrada = Replace(Entrada, "-", "")
        End If
        Entrada = SoloNumerico(Entrada)
        On Error Resume Next
        Select Case Tipo
        Case 1 ' Byte
            Err.Clear
            VerificaTextoNumerico = CByte(Entrada)
            If Err.Number <> 0 Then VerificaTextoNumerico = ""
        Case 2
            Err.Clear
            VerificaTextoNumerico = CInt(Entrada)
            If Err.Number <> 0 Then VerificaTextoNumerico = ""
        Case 3
            Err.Clear
            VerificaTextoNumerico = CLng(Entrada)
            If Err.Number <> 0 Then VerificaTextoNumerico = ""
        Case 4
            Err.Clear
            VerificaTextoNumerico = CSng(Entrada)
            If Err.Number <> 0 Then VerificaTextoNumerico = ""
        Case 5
            Err.Clear
            VerificaTextoNumerico = CDbl(Entrada)
            If Err.Number <> 0 Then VerificaTextoNumerico = ""
        Case Else
            VerificaTextoNumerico = Entrada
        End Select
    End If
End Function
Public Function RemplazaCaracter(Entrada As String, Anterior As String, Nuevo As String) As String
    Entrada = Trim(Entrada)
    RemplazaCaracter = Entrada
    Anterior = Mid(Trim(Anterior), 1, 1)
    Nuevo = Mid(Trim(Nuevo), 1, 1)
    If Not IsNull(RemplazaCaracter) Or Len(RemplazaCaracter) > 0 Then
        If Len(Anterior) > 0 And Len(Nuevo) > 0 Then RemplazaCaracter = Replace(Entrada, Anterior, Nuevo)
    End If
End Function
Public Function Trunca(ByRef Entrada As String, Separador As String) As String
Dim Posicion As Integer
    Posicion = InStr(1, Entrada, Separador)
    If Posicion > 0 Then
        Trunca = Mid(Entrada, 1, Posicion - 1)
        Entrada = Mid(Entrada, Posicion + 1)
    Else
        If Len(Entrada) > 0 Then
            Trunca = Entrada
            Entrada = ""
        End If
    End If
End Function
Public Sub LimpiaSeparador(ByRef Entrada As String, Separador As String)
    Separador = Mid(Separador, 1, 1)
    Entrada = QuitaDuplicados(Trim(Entrada), Separador)
    If Left(Entrada, 1) = Separador Then Entrada = Mid(Entrada, 2)
    If Right(Entrada, 1) = Separador Then Entrada = Mid(Entrada, 1, Len(Entrada) - 1)
End Sub
Public Function CuentaSeparador(Entrada As String, Separador As String) As Single
Dim I As Long
    Separador = Mid(Separador, 1, 1)
    LimpiaSeparador Entrada, Separador
    CuentaSeparador = 0
    For I = 1 To Len(Entrada)
        If Mid(Entrada, I, 1) = Separador Then CuentaSeparador = CuentaSeparador + 1
    Next I
End Function
Public Function LPSTRToVBString(ByVal Entrada As String) As String
    Dim NullPos As Long
    NullPos = InStr(Entrada, Chr(0))
    LPSTRToVBString = ""
    If NullPos > 0 Then LPSTRToVBString = Left(Entrada, NullPos - 1)
End Function
Public Function HexToInt(StrHex As String) As Integer
    HexToInt = CInt("&H" & StrHex)
End Function
Public Function HiByte(ByVal wParam As Integer)
    HiByte = wParam \ &H100 And &HFF&
End Function
Public Function LoByte(ByVal wParam As Integer)
    LoByte = wParam And &HFF&
End Function
'Desgloses
Public Property Let Desglose(ByVal vData As String)
    mvarDesglose = vData
End Property
Public Property Get Desglose() As String
    Desglose = mvarDesglose
End Property
Public Property Let Monedas(ByVal vData As String)
    mvarMonedas = Dimensiona(Trim(vData))
End Property
Public Property Get Monedas() As String
    Monedas = mvarMonedas
End Property
Public Property Let ImporteADesglosar(ByVal vData As Double)
    mvarImporteADesglosar = vData
    If Len(mvarMonedas) = 0 Then Monedas = "500;200;100;50;20;10;5;2;1;0,50;0,20;0,10;0,05;0,02;0,01"
    Desglose = Desglosa(mvarImporteADesglosar)
End Property
Public Property Get ImporteADesglosar() As Double
    ImporteADesglosar = mvarImporteADesglosar
End Property
Private Function Dimensiona(Entrada As String) As String
   Dim I As Integer, Auxiliar As String
    If Mid(Entrada, 1, 1) = ";" Then Entrada = Mid(Entrada, 2, Len(Entrada))
    If Mid(Entrada, Len(Entrada), 1) = ";" Then Entrada = Mid(Entrada, 1, Len(Entrada) - 1)
    Nmonedas = 0
    For I = 1 To Len(Entrada)
        If Mid(Entrada, I, 1) = ";" Then Nmonedas = Nmonedas + 1
    Next I
    Nmonedas = Nmonedas + 1
    Auxiliar = Entrada
    ReDim TMonedas(Nmonedas)
    ReDim TDesglose(Nmonedas)
    For I = 0 To Nmonedas - 1
        TMonedas(I) = TruncaDesglose(Auxiliar)
        TDesglose(I) = 0
    Next I
    Dimensiona = Entrada
End Function
Private Function TruncaDesglose(ByRef Entrada As String) As Double
Dim Posicion As Integer
    Posicion = InStr(1, Entrada, ";")
    If Posicion = 0 Then Posicion = Len(Entrada) + 1
    TruncaDesglose = CDbl(Mid(Entrada, 1, Posicion - 1))
    Entrada = Mid(Entrada, Posicion + 1, Len(Entrada))
End Function
Private Function Desglosa(Entrada As Double) As String
Dim Auximporte As Double, I As Integer, Salida As String, AuxDesglose As Double
    Auximporte = Entrada
    Salida = ""
    For I = 0 To Nmonedas - 1
        AuxDesglose = Auximporte / TMonedas(I)
        TDesglose(I) = Fix(AuxDesglose)
        Salida = Salida & TDesglose(I)
        If I < (Nmonedas - 1) Then Salida = Salida & ";"
        Auximporte = Auximporte - (TDesglose(I) * TMonedas(I))
    Next I
    Desglosa = Salida
End Function
'Services
Public Function AbreConectarUnidadDeRed(Formulario As Object) As Long
    AbreConectarUnidadDeRed = WNetConnectionDialog(Formulario.hwnd, RESOURCETYPE_DISK)
End Function
Public Function AbreDesconectarUnidadDeRed(Formulario As Object) As Long
    AbreDesconectarUnidadDeRed = WNetDisconnectDialog(Formulario.hwnd, RESOURCETYPE_DISK)
End Function
Public Function GetSerialNumber(sDrive As String) As Long
   Dim ser As Long
   Dim s As String * MAX_FILENAME_LEN
   Dim s2 As String * MAX_FILENAME_LEN
   Dim I As Long
   Dim j As Long
   Call GetVolumeInformation(sDrive + ":\" & Chr$(0), s, MAX_FILENAME_LEN, ser, I, j, s2, MAX_FILENAME_LEN)
   GetSerialNumber = ser
End Function
Public Function ComputerName() As String
Dim sComputerName As String, ComputerNameLength As Long
    sComputerName = String(MAX_COMPUTERNAME_LENGTH + 1, 0)
    ComputerNameLength = MAX_COMPUTERNAME_LENGTH
    Call GetComputerName(sComputerName, ComputerNameLength)
    ComputerName = Trim(Mid(sComputerName, 1, ComputerNameLength))
End Function
Public Function ConectarUnidadRed(Unidad As String, UnidadDisponible As String) As Long
Dim Disco As NETRESOURCE
    Disco.lpRemoteName = Unidad 'Dispositivo al que conectarse en formato UNC
    Disco.dwType = RESOURCETYPE_DISK ' Tipo de dispositivo
    Disco.lpLocalName = UnidadDisponible
    ConectarUnidadRed = WNetAddConnection2(Disco, "", "", 0)
End Function
Public Function DesconectarUnidadRed(Unidad As String) As Long
    DesconectarUnidadRed = WNetCancelConnection2(Unidad, 0, True)  ' Puedes enviar como Unidad un UNC o una normal "F:"
End Function
Public Function ExisteObjecto(Nombre As String) As Boolean
    Dim Objeto As Object
    On Error Resume Next
    ExisteObjecto = False
    Err.Clear
    Set Objeto = CreateObject(Nombre)
    If Err.Number = 0 Then ExisteObjecto = True
    Set Objeto = Nothing
End Function
Public Function GetDefPrinter() As String
Dim Def As String, Di As Long
    Def = String(128, 0)
    Di = GetProfileString("WINDOWS", "DEVICE", "", Def, 127)
    For Di = 1 To 128
        If Mid(Def, Di, 1) = "," Then Exit For
        GetDefPrinter = GetDefPrinter & Mid(Def, Di, 1)
    Next Di
    GetDefPrinter = Trim(GetDefPrinter)
End Function
Public Function GetUsuario() As String
Dim Resultado As Long, Buffer As String, Largo As Long
    Largo = 255
    Buffer = String(Largo, 0)
    Resultado = GetUserName(Buffer, Largo)
    GetUsuario = LPSTRToVBString(Left(Buffer, Largo))
    If GetUsuario = "" Then
        GetUsuario = RegGetString(&H80000002, "NETWORK\LOGON", "username")
    End If
End Function
Private Function RegGetString(hInKey As Long, ByVal SubKey As String, ByVal valname As String) As String
Dim retval As String, hSubKey As Long, dwType As Long, SZ As Long, v, r As Long
    retval = ""
    r = RegOpenKeyEx(hInKey, SubKey, 0, KEY_ALL_CLASSES, hSubKey)
    If r <> ERROR_SUCCESS Then GoTo Quit_Now
    SZ = 256: v = String(SZ, 0)
    r = RegQueryValueEx(hSubKey, valname, 0, dwType, ByVal v, SZ)
    If r = ERROR_SUCCESS And dwType = REG_SZ Then
        retval = Left(v, SZ - 1)
    Else
        retval = ""
    End If
    If hInKey = 0 Then r = RegCloseKey(hSubKey)
Quit_Now:
    RegGetString = retval
End Function
Private Function GetPointerToByteStringW(ByVal dwData As Long) As String
   Dim tmp() As Byte, tmplen As Long
   If dwData <> 0 Then
      tmplen = lstrlenW(dwData) * 2
      If tmplen <> 0 Then
         ReDim tmp(0 To (tmplen - 1)) As Byte
         CopyMemory tmp(0), ByVal dwData, tmplen
         GetPointerToByteStringW = tmp
     End If
   End If
End Function
Public Function GetServers(ByRef Lista As String) As Long
Dim bufptr As Long, dwEntriesread As Long, dwTotalentries     As Long, dwResumehandle     As Long, se100              As SERVER_INFO_100, success            As Long, nStructSize        As Long, cnt                As Long
    nStructSize = LenB(se100)
    Lista = ""
    success = NetServerEnum(0&, 100, bufptr, MAX_PREFERRED_LENGTH, dwEntriesread, dwTotalentries, SV_TYPE_ALL, 0&, dwResumehandle)
    If success = ERROR_SUCCESS And success <> ERROR_MORE_DATA Then
        For cnt = 0 To dwEntriesread - 1
            CopyMemory se100, ByVal bufptr + (nStructSize * cnt), nStructSize
            If Len(Lista) > 0 Then Lista = Lista & ";"
            Lista = Lista & GetPointerToByteStringW(se100.sv100_name)
        Next
    End If
    Call NetApiBufferFree(bufptr)
    GetServers = dwEntriesread
End Function
Public Function GetSeparadorMiles() As String
Dim Retorno As String * 100, Resultado As Long
    GetSeparadorMiles = ""
    Resultado = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_STHOUSAND, Retorno, 99)
    If Resultado > 0 Then GetSeparadorMiles = LPSTRToVBString(Retorno)
End Function
Public Function GetSeparadorDecimal() As String
Dim Retorno As String * 100, Resultado As Long
    GetSeparadorDecimal = ""
    Resultado = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SDECIMAL, Retorno, 99)
    If Resultado > 0 Then GetSeparadorDecimal = LPSTRToVBString(Retorno)
End Function
Public Function GetSeparadorFecha() As String
Dim Retorno As String * 100, Resultado As Long
    GetSeparadorFecha = ""
    Resultado = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SDATE, Retorno, 99)
    If Resultado > 0 Then GetSeparadorFecha = LPSTRToVBString(Retorno)
End Function
Public Function GetSeparadorHora() As String
Dim Retorno As String * 100, Resultado As Long
    GetSeparadorHora = ""
    Resultado = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SLONGDATE, Retorno, 99)
    If Resultado > 0 Then GetSeparadorHora = LPSTRToVBString(Retorno)
End Function
Public Function GetFormatoFechaCorta() As String
Dim Retorno As String * 100, Resultado As Long
    GetFormatoFechaCorta = ""
    Resultado = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SSHORTDATE, Retorno, 99)
    If Resultado > 0 Then GetFormatoFechaCorta = LPSTRToVBString(Retorno)
End Function
Public Function GetFormatoFechaLarga() As String
Dim Retorno As String * 100, Resultado As Long
    GetFormatoFechaLarga = ""
    Resultado = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_STIME, Retorno, 99)
    If Resultado > 0 Then GetFormatoFechaLarga = LPSTRToVBString(Retorno)
End Function
Public Function GetFormatoHora() As String
Dim Retorno As String * 100, Resultado As Long
    GetFormatoHora = ""
    Resultado = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_STIMEFORMAT, Retorno, 99)
    If Resultado > 0 Then GetFormatoHora = LPSTRToVBString(Retorno)
End Function
Public Sub ReportException(Aplicación As String, Puesto As String, Usuario As String, Exception As String)
Dim fHandle As Integer
    fHandle = FreeFile
    Open "Errors.Log" For Append As #fHandle
    Print #fHandle, Now() & " *" & Trim(Mid(Usuario, 1, 20)) & "*" & Trim(Mid(Puesto, 1, 20)) & "*" & Trim(Mid(Aplicación, 1, 20)) & "* " & Exception
    Close #fHandle
End Sub
Public Function Spell(Texto As String) As String
Dim ObjetoWord As Object
    Spell = ""
    Set ObjetoWord = CreateObject("Word.Application")
    With ObjetoWord
        .Documents.Add
        .Selection.TypeText Trim(Texto)
        .Visible = True
        .ActiveDocument.CheckSpelling
        Spell = Trim(.ActiveDocument.StoryRanges(1))
        .Quit 0
    End With
    Set ObjetoWord = Nothing
End Function
Public Function CerrarObjetoOle(Objeto As Object) As Boolean
    On Error Resume Next
    Err.Clear
    Objeto.Visible = True
    If Err.Number <> 0 Then
        MsgBox "Error al hacer visible el objeto " & Objeto.Name & vbCrLf & "Error: " & Err.Description
        CerrarObjetoOle = False
    Else
        Err.Clear
        Objeto.Quit
        If Err.Number <> 0 Then
            MsgBox "Error al hacer Cerrar el objeto " & Objeto.Name & vbCrLf & "Error: " & Err.Description
            CerrarObjetoOle = False
        Else
            Set Objeto = Nothing
            CerrarObjetoOle = True
        End If
    End If
    DoEvents
End Function
Public Function CerrarLibroExcel(ByRef Objeto As Object, Nombre As String, Salvar As Boolean) As Boolean
    Objeto.Visible = True
    If Salvar Then
        Objeto.ActiveWorkbook.SaveAs Nombre, -4143, "", "", False, False, True
        Objeto.ActiveWorkbook.Close (False)
    Else
        Objeto.ActiveWorkbook.Close (False)
    End If
End Function
Public Function CrearObjetoExcel(ByRef Objeto As Object, Visible As Boolean) As Boolean
    On Error Resume Next
    Err.Clear
    Set Objeto = CreateObject("Excel.application")
    If Err.Number <> 0 Then
        MsgBox "Error al crear el objeto Excel" & vbCrLf & "Error: " & Err.Description
        CrearObjetoExcel = False
    Else
        If Visible Then Objeto.Visible = True
        CrearObjetoExcel = True
    End If
End Function
Public Function CrearLibroExcel(ByRef Objeto As Object, Plantilla As String) As Boolean
    On Error Resume Next
    Dim TextoPlantilla As String
    If Len(Trim(" " & Plantilla)) > 0 Then
        Err.Clear
        Objeto.Workbooks.Add (Trim(Plantilla))
    Else
        Err.Clear
        Objeto.Workbooks.Add
    End If
    If Err.Number <> 0 Then
        MsgBox "Error al crear el Libro en " & Objeto.Name & vbCrLf & "Error: " & Err.Description
        CrearLibroExcel = False
    Else
        CrearLibroExcel = True
    End If
End Function
Public Function TablaAdoAExcel(Cursor As ADODB.Recordset, Objeto As Object, Posicion As String) As Boolean
Dim NumeroDeCampos As Long, I As Long
    TablaAdoAExcel = False
    If Cursor.EOF And Cursor.BOF Then
        MsgBox "NO hay datos que descargar", vbCritical, "Descarga de Tablas"
    Else
        On Error Resume Next
        NumeroDeCampos = Cursor.Fields.Count
        Cursor.MoveFirst
        Objeto.Range(Posicion).Select
        For I = 0 To NumeroDeCampos - 1
            Objeto.ActiveCell = Cursor.Fields(I).Name
            Objeto.ActiveCell.Offset(0, 1).Range("A1").Select
        Next I
        Objeto.ActiveCell.Offset(1, (NumeroDeCampos * -1)).Range("A1").Select
        DoEvents
        Do Until Cursor.EOF
            For I = 0 To NumeroDeCampos - 1
                Objeto.ActiveCell = Cursor.Fields(I).Value
                Objeto.ActiveCell.Offset(0, 1).Range("A1").Select
            Next I
            Objeto.ActiveCell.Offset(1, (NumeroDeCampos * -1)).Range("A1").Select
            DoEvents
            Cursor.MoveNext
        Loop
        Objeto.Cells.Select
        Objeto.Cells.EntireColumn.AutoFit
        Objeto.Range(Posicion).Select
        TablaAdoAExcel = True
    End If
End Function
Public Function CerrarDocumentoWord(ByRef Objeto As Object, Nombre As String, Salvar As Boolean) As Boolean
    Objeto.Visible = True
    If Salvar Then
        Objeto.ActiveDocument.SaveAs Nombre, 0, False, "", True, "", False, False, False, False, False
        Objeto.ActiveDocument.Close (False)
    Else
        Objeto.ActiveDocument.Close (False)
    End If
End Function
Public Function CrearObjetoWord(ByRef Objeto As Object, Visible As Boolean) As Boolean
    On Error Resume Next
    Err.Clear
    Set Objeto = CreateObject("Word.application")
    If Err.Number <> 0 Then
        MsgBox "Error al crear el objeto Word" & vbCrLf & "Error: " & Err.Description
        CrearObjetoWord = False
    Else
        If Visible Then Objeto.Visible = True
        CrearObjetoWord = True
    End If
End Function
Public Function CrearDocumentoWord(ByRef Objeto As Object, Plantilla As String) As Boolean
    On Error Resume Next
    Err.Clear
    Objeto.Documents.Add Plantilla, False
    If Err.Number <> 0 Then
        MsgBox "Error al crear el documento en " & Objeto.Name & vbCrLf & "Error: " & Err.Description
        CrearDocumentoWord = False
    Else
        CrearDocumentoWord = True
    End If
End Function
Public Function TablaAdoAWord(Cursor As ADODB.Recordset, Objeto As Object) As Boolean
Dim NumeroDeCampos As Long, I As Long
    TablaAdoAWord = False
    If Cursor.EOF And Cursor.BOF Then
        MsgBox "NO hay datos que descargar", vbCritical, "Descarga de Tablas"
    Else
        On Error Resume Next
        NumeroDeCampos = Cursor.Fields.Count
        Cursor.MoveFirst
        Objeto.ActiveDocument.Tables.Add Objeto.Selection.Range, 1, NumeroDeCampos
        For I = 0 To NumeroDeCampos - 1
            Objeto.Selection.TypeText Trim(Cursor.Fields(I).Name)
            If I < NumeroDeCampos - 1 Then
                Objeto.Selection.MoveRight 12
            Else
                Objeto.Selection.MoveRight 12
                Objeto.Selection.MoveLeft 12
                Objeto.Selection.SelectRow
                Objeto.Selection.Rows.HeadingFormat = 9999998
                Objeto.Selection.MoveRight 12
            End If
        Next I
        Objeto.Selection.MoveRight 12, NumeroDeCampos
        Do Until Cursor.EOF
            For I = 0 To NumeroDeCampos - 1
                ' Escribir el valor en la celda
                Select Case Cursor.Fields(I).Type
                    Case 2, 18  ' Enteros
                        Objeto.Selection.ParagraphFormat.Alignment = 2
                        Objeto.Selection.TypeText Trim(Format(Cursor.Fields(I).Value, "##,##0"))
                    Case 3, 19  ' Enteros Largos
                        Objeto.Selection.ParagraphFormat.Alignment = 2
                        Objeto.Selection.TypeText Trim(Format(Cursor.Fields(I).Value, "##,###,###,##0"))
                    Case 4   ' Decimal
                        Objeto.Selection.TypeText Trim(Format(Cursor.Fields(I).Value, "Standard"))
                        Objeto.Selection.ParagraphFormat.Alignment = 2
                    Case 5, 14, 131 ' Doble y Exact
                        Objeto.Selection.ParagraphFormat.Alignment = 2
                        Objeto.Selection.TypeText Trim(Format(Cursor.Fields(I).Value, "Standard"))
                    Case 6   ' Moneda
                        Objeto.Selection.ParagraphFormat.Alignment = 2
                        Objeto.Selection.TypeText Trim(Format(Cursor.Fields(I).Value, "Currency"))
                    Case 11  ' Bolean
                        Objeto.Selection.ParagraphFormat.Alignment = 2
                        Objeto.Selection.TypeText Trim(Format(Cursor.Fields(I).Value, "Yes/No"))
                    Case 17  ' byte
                        Objeto.Selection.ParagraphFormat.Alignment = 2
                        Objeto.Selection.TypeText Trim(Format(Cursor.Fields(I).Value, "##0"))
                    Case 16, 20, 21 ' Big Int
                        Objeto.Selection.ParagraphFormat.Alignment = 2
                        Objeto.Selection.TypeText Trim(Format(Cursor.Fields(I).Value, "##,###,###,###,###,###,##0"))
                    Case 7, 133, 134, 135 ' Fecha Hora
                        Objeto.Selection.ParagraphFormat.Alignment = 1
                        Objeto.Selection.TypeText Trim(Format(Cursor.Fields(I).Value, "General Date"))
                    Case 103, 129, 136, 200, 201, 202, 203 ' Texto
                        Objeto.Selection.TypeText Trim(Cursor.Fields(I).Value)
                    Case Else ' Otros
                        Objeto.Selection.TypeText Cursor.Fields(I).Value
                    End Select
                'pasar a la siguente celda
                Objeto.Selection.MoveRight 12
            Next I
            DoEvents
            Cursor.MoveNext
        Loop
        Objeto.Selection.SelectRow
        Objeto.Selection.Rows.Delete
        Objeto.Selection.Tables(1).Select
        Objeto.Selection.Cells.HeightRule = 0
        Objeto.Selection.Cells.AutoFit
        TablaAdoAWord = True
    End If
End Function
Public Function TablaAdoAFichero(Cursor As ADODB.Recordset, FicheroSalida As String, Separador As String, Formatear As Boolean) As Boolean
Dim CanalFichero As Integer, NumeroDeCampos As Long, TextoMensaje As String, respuesta As Integer, I As Long, RegistroSalida As String, TextoCampo As String
Dim PosicionCursor As Variant
    TablaAdoAFichero = False
    FicheroSalida = Trim(FicheroSalida)
    If Len(FicheroSalida) = 0 Then FicheroSalida = "C:\Downloads\Descarga.txt"
    Separador = Trim(Separador)
    If Len(Separador) = 0 Then
        Separador = "|"
    Else
        Separador = Mid(Separador, 1, 1)
    End If
    If Cursor.EOF And Cursor.BOF Then
        MsgBox "NO hay datos que descargar", vbCritical, "Descarga de Tablas"
    Else
        On Error Resume Next
        CanalFichero = FreeFile
        Err.Clear
        Open FicheroSalida For Input As #CanalFichero
        If Err.Number = 0 Then
            TextoMensaje = "El fichero '" & vbCrLf & FicheroSalida & "' ya existe." & vbCrLf & " ¿Lo Remplazo?"
            respuesta = MsgBox(TextoMensaje, 547, "Descarga de Tablas")
            Close #CanalFichero
            CanalFichero = FreeFile
            Select Case respuesta
                Case 6 ' SI
                    Open FicheroSalida For Output As #CanalFichero
                Case 7 ' NO
                    Open FicheroSalida For Append As #CanalFichero
                Case 2 ' CANCELAR
                    Exit Function
                End Select
        Else
            Open FicheroSalida For Output As #CanalFichero
        End If
        NumeroDeCampos = Cursor.Fields.Count
        Cursor.MoveFirst
        Do Until Cursor.EOF
            RegistroSalida = ""
            For I = 0 To NumeroDeCampos - 1
                If Formatear Then
                    Select Case Cursor.Fields(I).Type
                        Case 2, 18  ' Enteros
                            TextoCampo = Trim(Format(Cursor.Fields(I).Value, "##,##0"))
                        Case 3, 19  ' Enteros Largos
                            TextoCampo = Trim(Format(Cursor.Fields(I).Value, "##,###,###,##0"))
                        Case 4   ' Decimal
                            TextoCampo = Trim(Format(Cursor.Fields(I).Value, "Standard"))
                        Case 5, 14, 131 ' Doble y Exact
                            TextoCampo = Trim(Format(Cursor.Fields(I).Value, "Standard"))
                        Case 6   ' Moneda
                            TextoCampo = Trim(Format(Cursor.Fields(I).Value, "Currency"))
                        Case 11  ' Bolean
                            TextoCampo = Trim(Format(Cursor.Fields(I).Value, "Yes/No"))
                        Case 17  ' byte
                            TextoCampo = Trim(Format(Cursor.Fields(I).Value, "##0"))
                        Case 16, 20, 21 ' Big Int
                            TextoCampo = Trim(Format(Cursor.Fields(I).Value, "##,###,###,###,###,###,##0"))
                        Case 7, 133, 134, 135 ' Fecha Hora
                            TextoCampo = Trim(Format(Cursor.Fields(I).Value, "General Date"))
                        Case 103, 129, 136, 200, 201, 202, 203 ' Texto
                            TextoCampo = Trim(Cursor.Fields(I).Value)
                        Case Else ' Otros
                            TextoCampo = Cursor.Fields(I).Value
                    End Select
                Else
                    TextoCampo = Cursor.Fields(I).Value
                End If
                If I = NumeroDeCampos - 1 Then
                    RegistroSalida = RegistroSalida & TextoCampo
                Else
                    RegistroSalida = RegistroSalida & TextoCampo & Separador
                End If
            Next I
            Print #CanalFichero, RegistroSalida
            DoEvents
            Cursor.MoveNext
        Loop
        Close #CanalFichero
        Cursor.Bookmark = PosicionCursor
        TablaAdoAFichero = True
    End If
End Function
Public Sub GuardarUDL(ByVal strConnectionString As String, ByVal strRutaUDL As String)
Dim CanalFichero As Integer
    strConnectionString = UDL_LINEA1 & vbCrLf & UDL_LINEA2 & vbCrLf & strConnectionString & vbCrLf
    strConnectionString = Chr$(&HFF) & Chr$(&HFE) & StrConv(strConnectionString, vbUnicode)
    CanalFichero = FreeFile
    Open strRutaUDL For Binary As #CanalFichero
    Put #CanalFichero, , strConnectionString
    Close #CanalFichero
End Sub


