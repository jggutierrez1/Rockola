VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Video_Form 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   8475
   ClientLeft      =   12240
   ClientTop       =   -1455
   ClientWidth     =   10140
   LinkTopic       =   "Form2"
   Moveable        =   0   'False
   ScaleHeight     =   565
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   676
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Image oImg_Logo1 
      Height          =   8400
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   10080
   End
   Begin WMPLibCtl.WindowsMediaPlayer MediaPlayer3 
      Height          =   8415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10095
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   0   'False
      currentMarker   =   0
      invokeURLs      =   0   'False
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   -1  'True
      windowlessVideo =   0   'False
      enabled         =   0   'False
      enableContextMenu=   0   'False
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   17806
      _cy             =   14843
   End
End
Attribute VB_Name = "Video_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
Call AlwaysOnTop(Main_Form, False)
Main_Form.oSetFocus_Codigo.SetFocus
End Sub

Private Sub MediaPlayer3_MediaError(ByVal pMediaObject As Object)
Main_Form.olMessage.Visible = True
Main_Form.olMessage.Caption = "TEMA NO DISPONIBLE"
Main_Form.oTime_Mensajes.Enabled = True
Call Remove_Temes
igCnt_CR = igCnt_CR + 1
Call Refresh_Creditos(Main_Form)
Sleep 3 '* 1000 'Implements a 3 second delay
VBA.SendKeys ("S")
End Sub

Private Sub MediaPlayer3_PlayStateChange(ByVal NewState As Long)
Select Case NewState
Case Is = wmppsMediaEnded
   bgWMP_Busy = False
    Call Muestra_Tema_Det
    'Video_Form.MediaPlayer3.Close
    Main_Form.MediaPlayer2.Close
    igCont_Sin = 0
    If igDelay_Bonus_Vid > 0 Then
        igNext_Bonus = (Hour(Time()) * 60) + (Minute(Time()) + igDelay_Bonus_Vid)
    End If
    If bgIs_Video = False Then
        Exit Sub
    Else
        If igScr_Alone = 0 Then
            bfTm = False
        End If
    End If
    Call Remove_Temes
Case Is = wmppsPlaying
    If igCont_Sin > 0 Then
        Exit Sub
    End If
    If bgIs_Video = True Then
        If igScr_Alone = 0 Then
            bfTm = True
            Main_Form.oTM_Box.Enabled = True
        End If
    End If
    bgWMP_Busy = True
    Main_Form.MediaPlayer2.URL = Video_Form.MediaPlayer3.URL
    Main_Form.MediaPlayer2.Controls.currentPosition = Video_Form.MediaPlayer3.Controls.currentPosition
    Main_Form.MediaPlayer2.settings.mute = True
    Main_Form.MediaPlayer2.Controls.play
    igCont_Sin = igCont_Sin + 1
End Select
End Sub

