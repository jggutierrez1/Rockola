Attribute VB_Name = "CopyFastM"
Option Explicit

'---Copy a file the using SHFileOperation API call so that Windows copy progress dialog appears---
Private Const FOF_FILESONLY = &H80& 'Perform the operation
Public Type SHFILEOPSTRUCT
hWnd As Long
wFunc As Long
pFrom As String
pTo As String
fFlags As Integer
fAnyOperationsAborted As Boolean
hNameMappings As Long
lpszProgressTitle As String
End Type

Public Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Public Const FO_COPY = &H2
Public Const FOF_ALLOWUNDO = &H40

'EXAMPLE:
'dim bSuccess as boolean
'bSuccess = APIFileCopy ("C:\MyFile.txt", "D:\MyFile.txt")
Public Declare Function CopyFile Lib "kernel32" _
  Alias "CopyFileA" (ByVal lpExistingFileName As String, _
  ByVal lpNewFileName As String, ByVal bFailIfExists As Long) _
  As Long
  
Public Function CopyFileWindowsWay(SourceFile As String, DestinationFile As String, sError As String, Optional iFlag As Integer = 0) As Boolean
Dim lngReturn As Long
Dim typFileOperation As SHFILEOPSTRUCT
With typFileOperation
    .hWnd = 0
    .wFunc = FO_COPY
    .pFrom = SourceFile & vbNullChar & vbNullChar 'source file
    .pTo = DestinationFile & vbNullChar & vbNullChar 'destination file
    .fFlags = VBA.IIf(iFlag = 0, FOF_ALLOWUNDO, FOF_FILESONLY)
    .fFlags = FOF_ALLOWUNDO
End With
If iFlag = 0 Then
    If FileExist(SourceFile) = False Then
        sError = "Status: Error origen " & SourceFile & ", no existe!"
        CopyFileWindowsWay = False
        Exit Function
    End If
End If
lngReturn = SHFileOperation(typFileOperation)
If lngReturn <> 0 Then 'Operation failed
    sError = "Status: " & Err.Description
    CopyFileWindowsWay = False
Else 'Aborted
    sError = "Status: OK..."
    CopyFileWindowsWay = True
    If typFileOperation.fAnyOperationsAborted = True Then
        sError = "Status: Falla encontrada: "
        CopyFileWindowsWay = False
    End If
End If
End Function

Public Function CopyFast(Source As String, Destination As String, ProgressBar As Object, Optional sError As String = "")
'=========================================================================================
'This sub copies a file, and displays a progress bar of your choice whilst it is doing it.
'It is dead useful for saving large files in your program.
'
'When calling it, you need to pass to it the source and destination files,
'and the name of the progressbar, e.g:
'
'CopyFast "file1.txt", "file2.txt", FrmMain.ProgressBar1
'
'it reads the file in 20k chunks, which is much faster than reading byte
'by byte and displaying progress.
'You'll have to do your own errorhandling! e.g. if destination file cannot be accessed.
'
'by Andrew Beverley
'email: andrew.beverley@parthus.com
'=========================================================================================

Dim Buffer As String 'buffer used for copying
Dim Pointer As Long 'position of pointer in file
Dim X As Integer 'used in for...next loop
Dim Whole As Integer
Dim Part As Integer

If Dir(Destination) <> "" Then Kill (Destination) 'checks if destination file exists. if it does, delete it

Open Source For Binary As #1
Open Destination For Binary As #2

On Error GoTo Solve_error

Whole = LOF(1) \ 20000 'number of whole 20k chunks
Part = LOF(1) Mod 20000 'remainder at the end
Buffer = String$(20000, 0) 'buffer
Pointer = 1 'start at position 1
ProgressBar.Max = IIf(LOF(1) = 0, 1, LOF(1))
If LOF(1) = 0 Then
    Debug.Print "Archivo vacío: " & VBA.Trim(Source)
End If
ProgressBar.Visible = True
For X = 1 To Whole
    Get #1, Pointer, Buffer 'get data
    Put #2, Pointer, Buffer 'put it to destination file
    Pointer = Pointer + 20000 'put pointer 20k later
    If Pointer < ProgressBar.Max Then ProgressBar.value = Pointer Else ProgressBar.value = ProgressBar.Max 'update progressbar, make sure it doesn't overflow
Next X

Buffer = String$(Part, 0) 'copy the last bit
Get #1, Pointer, Buffer        'get the remaining bytes at the end
Put #2, Pointer, Buffer       'put them
ProgressBar.value = ProgressBar.Max
Close

Buffer = "" 'reclaim space
CopyFast = True
sError = ""
Exit Function

Solve_error:
sError = Err.Description
CopyFast = False
End Function
