VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{EEE78583-FE22-11D0-8BEF-0060081841DE}#1.0#0"; "XVoice.dll"
Begin VB.UserControl SpeakingHead 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   ClientHeight    =   7185
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6990
   ScaleHeight     =   7185
   ScaleWidth      =   6990
   ToolboxBitmap   =   "SpeakingHead.ctx":0000
   Begin VB.CommandButton cmdStop 
      Caption         =   "STOP"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   6000
      Width           =   1095
   End
   Begin MSComctlLib.ImageList imgLSpeak 
      Left            =   4080
      Top             =   5640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   151
      ImageHeight     =   230
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SpeakingHead.ctx":0312
            Key             =   "aa"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SpeakingHead.ctx":1952
            Key             =   "aah"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SpeakingHead.ctx":2FB2
            Key             =   "bmp"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SpeakingHead.ctx":45DA
            Key             =   "dst"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SpeakingHead.ctx":5C36
            Key             =   "ee"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SpeakingHead.ctx":72AE
            Key             =   "eh"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SpeakingHead.ctx":8912
            Key             =   "fv"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SpeakingHead.ctx":9F62
            Key             =   "i"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SpeakingHead.ctx":B5E6
            Key             =   "k"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SpeakingHead.ctx":CC3E
            Key             =   "n"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SpeakingHead.ctx":E29A
            Key             =   "oh"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SpeakingHead.ctx":F8D2
            Key             =   "q"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SpeakingHead.ctx":10F2E
            Key             =   "r"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SpeakingHead.ctx":12582
            Key             =   "th"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SpeakingHead.ctx":13BF6
            Key             =   "w"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SpeakingHead.ctx":1524E
            Key             =   "silent"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Slider Ratesldr 
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   3840
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   873
      _Version        =   393216
      LargeChange     =   1
      Min             =   -10
   End
   Begin VB.CommandButton cmdSpeak 
      Caption         =   "SPEAK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   6000
      Width           =   1215
   End
   Begin VB.TextBox MainTxtBox 
      Height          =   1695
      HideSelection   =   0   'False
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "SpeakingHead.ctx":16852
      Top             =   4320
      Width           =   2295
   End
   Begin VB.ComboBox VoiceCB 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   3480
      Width           =   2295
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3960
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   151
      ImageHeight     =   230
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   45
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SpeakingHead.ctx":16869
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SpeakingHead.ctx":17E71
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SpeakingHead.ctx":19475
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SpeakingHead.ctx":1AA61
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SpeakingHead.ctx":1C069
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SpeakingHead.ctx":1D671
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SpeakingHead.ctx":1EC89
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SpeakingHead.ctx":2029D
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SpeakingHead.ctx":218B1
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SpeakingHead.ctx":22EC5
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SpeakingHead.ctx":244DD
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SpeakingHead.ctx":25AE5
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SpeakingHead.ctx":270F5
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SpeakingHead.ctx":28705
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SpeakingHead.ctx":29CF1
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SpeakingHead.ctx":2B305
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SpeakingHead.ctx":2C919
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SpeakingHead.ctx":2DF35
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SpeakingHead.ctx":2F541
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SpeakingHead.ctx":30B41
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SpeakingHead.ctx":3214D
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SpeakingHead.ctx":33769
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SpeakingHead.ctx":34D7D
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SpeakingHead.ctx":36391
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SpeakingHead.ctx":379A9
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SpeakingHead.ctx":38FC5
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SpeakingHead.ctx":3A5E9
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SpeakingHead.ctx":3BBF9
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SpeakingHead.ctx":3D21D
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SpeakingHead.ctx":3E839
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SpeakingHead.ctx":3FE4D
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SpeakingHead.ctx":4145D
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SpeakingHead.ctx":42A71
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SpeakingHead.ctx":44081
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SpeakingHead.ctx":4566D
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SpeakingHead.ctx":46C81
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SpeakingHead.ctx":48291
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SpeakingHead.ctx":49899
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SpeakingHead.ctx":4AEA9
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SpeakingHead.ctx":4C4B1
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SpeakingHead.ctx":4DAB1
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SpeakingHead.ctx":4F0B1
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SpeakingHead.ctx":506AD
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SpeakingHead.ctx":51CC9
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SpeakingHead.ctx":532C9
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   4920
      Top             =   4200
   End
   Begin ACTIVEVOICEPROJECTLibCtl.DirectSS DirectSS1 
      Height          =   705
      Left            =   4920
      OleObjectBlob   =   "SpeakingHead.ctx":548CD
      TabIndex        =   4
      Top             =   5160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image imgIdle 
      Height          =   3450
      Left            =   0
      Picture         =   "SpeakingHead.ctx":54925
      Stretch         =   -1  'True
      ToolTipText     =   "http://interpret.co.za"
      Top             =   0
      Width           =   2265
   End
End
Attribute VB_Name = "SpeakingHead"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim j As Integer
Dim check As String, QQ As String

Private Sub cmdSpeak_Click()
DirectSS1.VolumeLeft = (100 * 655.35)
DirectSS1.VolumeRight = (100 * 655.35)
DirectSS1.Speed = 100


DirectSS1.Speak MainTxtBox.Text

End Sub

Private Sub cmdStop_Click()
DirectSS1.AudioReset
End Sub

Private Sub DirectSS1_AudioStart(ByVal hi As Long, ByVal lo As Long)
Timer1.Enabled = False
End Sub

Private Sub DirectSS1_AudioStop(ByVal hi As Long, ByVal lo As Long)
Timer1.Enabled = True
'MsgBox check
check = ""
End Sub

Private Sub DirectSS1_Visual(ByVal timehi As Long, ByVal timelo As Long, ByVal Phoneme As Integer, ByVal EnginePhoneme As Integer, ByVal hints As Long, ByVal MouthHeight As Integer, ByVal bMouthWidth As Integer, ByVal bMouthUpturn As Integer, ByVal bJawOpen As Integer, ByVal TeethUpperVisible As Integer, ByVal TeethLowerVisible As Integer, ByVal TonguePosn As Integer, ByVal LipTension As Integer)
check = check & ", " & MouthHeight & ":" & bMouthWidth


If InStr(QQ, "Afrikaans") Or InStr(QQ, "Dutch") Then
If MouthHeight = 0 And bMouthWidth = 0 Then
imgIdle.Picture = imgLSpeak.Overlay("silent", "silent")
End If
If MouthHeight = 64 And bMouthWidth = 128 Then
imgIdle.Picture = imgLSpeak.Overlay("ee", "ee")
End If
If MouthHeight = 60 And bMouthWidth = 100 Then
imgIdle.Picture = imgLSpeak.Overlay("ee", "ee")
End If
If MouthHeight = 32 And bMouthWidth = 128 Then
imgIdle.Picture = imgLSpeak.Overlay("n", "n")
End If
If MouthHeight = 48 And bMouthWidth = 128 Then
imgIdle.Picture = imgLSpeak.Overlay("dst", "dst")
End If
If MouthHeight = 10 And bMouthWidth = 240 Then
imgIdle.Picture = imgLSpeak.Overlay("w", "w")
End If
If MouthHeight = 40 And bMouthWidth = 220 Then
imgIdle.Picture = imgLSpeak.Overlay("w", "w")
End If
If MouthHeight = 64 And bMouthWidth = 128 Then
imgIdle.Picture = imgLSpeak.Overlay("k", "k") '@r
End If
If MouthHeight = 60 And bMouthWidth = 100 Then
imgIdle.Picture = imgLSpeak.Overlay("k", "k") '@r
End If
If MouthHeight = 40 And bMouthWidth = 220 Then
imgIdle.Picture = imgLSpeak.Overlay("i", "i")
End If
If MouthHeight = 32 And bMouthWidth = 112 Then
imgIdle.Picture = imgLSpeak.Overlay("i", "i") 'h
End If
If MouthHeight = 40 And bMouthWidth = 50 Then
imgIdle.Picture = imgLSpeak.Overlay("q", "q") 'oe
End If
If MouthHeight = 116 And bMouthWidth = 114 Then
imgIdle.Picture = imgLSpeak.Overlay("i", "i") 'g
End If
If MouthHeight = 128 And bMouthWidth = 150 Then
imgIdle.Picture = imgLSpeak.Overlay("aah", "aah")
End If
If MouthHeight = 128 And bMouthWidth = 160 Then
imgIdle.Picture = imgLSpeak.Overlay("oh", "oh") 'ui
End If
If MouthHeight = 40 And bMouthWidth = 128 Then
imgIdle.Picture = imgLSpeak.Overlay("dst", "dst")
End If
If MouthHeight = 0 And bMouthWidth = 150 Then
imgIdle.Picture = imgLSpeak.Overlay("bmp", "bmp")
End If
If MouthHeight = 80 And bMouthWidth = 160 Then
imgIdle.Picture = imgLSpeak.Overlay("eh", "eh")
End If
If MouthHeight = 96 And bMouthWidth = 160 Then
imgIdle.Picture = imgLSpeak.Overlay("oh", "oh")
End If
If MouthHeight = 20 And bMouthWidth = 100 Then
imgIdle.Picture = imgLSpeak.Overlay("fv", "fv")
End If

If MouthHeight = 100 And bMouthWidth = 160 Then
imgIdle.Picture = imgLSpeak.Overlay("aa", "aa")
End If
If MouthHeight = 0 And bMouthWidth = 128 Then
imgIdle.Picture = imgLSpeak.Overlay("bmp", "bmp")
End If
If MouthHeight = 127 And bMouthWidth = 98 Then
imgIdle.Picture = imgLSpeak.Overlay("k", "k")
End If
If MouthHeight = 189 And bMouthWidth = 159 Then
imgIdle.Picture = imgLSpeak.Overlay("oh", "oh")
End If
If MouthHeight = 80 And bMouthWidth = 90 Then
imgIdle.Picture = imgLSpeak.Overlay("oh", "oh")
End If
If MouthHeight = 40 And bMouthWidth = 64 Then
imgIdle.Picture = imgLSpeak.Overlay("oh", "oh")
End If
If MouthHeight = 60 And bMouthWidth = 100 Then
imgIdle.Picture = imgLSpeak.Overlay("aa", "aa")
End If
If MouthHeight = 48 And bMouthWidth = 128 Then
imgIdle.Picture = imgLSpeak.Overlay("k", "k")
End If
Else




If MouthHeight = 0 And bMouthWidth = 117 Then
imgIdle.Picture = imgLSpeak.Overlay("silent", "silent")
End If
If MouthHeight = 179 And bMouthWidth = 0 Then
imgIdle.Picture = imgLSpeak.Overlay("q", "q")
End If
If MouthHeight = 48 And bMouthWidth = 208 Then
imgIdle.Picture = imgLSpeak.Overlay("aa", "aa")
End If
If MouthHeight = 32 And bMouthWidth = 208 Then
imgIdle.Picture = imgLSpeak.Overlay("n", "n")
End If
If MouthHeight = 37 And bMouthWidth = 46 Then
imgIdle.Picture = imgLSpeak.Overlay("th", "th")
End If
If MouthHeight = 32 And bMouthWidth = 160 Then
imgIdle.Picture = imgLSpeak.Overlay("ee", "ee")
End If
If MouthHeight = 0 And bMouthWidth = 64 Then
imgIdle.Picture = imgLSpeak.Overlay("bmp", "bmp")
End If
If MouthHeight = 53 And bMouthWidth = 117 Then
imgIdle.Picture = imgLSpeak.Overlay("k", "k") 'l
End If
If MouthHeight = 16 And bMouthWidth = 208 Then
imgIdle.Picture = imgLSpeak.Overlay("dst", "dst") 's
End If
If MouthHeight = 48 And bMouthWidth = 160 Then
imgIdle.Picture = imgLSpeak.Overlay("eh", "eh")
End If
If MouthHeight = 64 And bMouthWidth = 128 Then
imgIdle.Picture = imgLSpeak.Overlay("r", "r") '@
End If
If MouthHeight = 32 And bMouthWidth = 192 Then
imgIdle.Picture = imgLSpeak.Overlay("q", "q") 'your
End If
If MouthHeight = 112 And bMouthWidth = 128 Then
imgIdle.Picture = imgLSpeak.Overlay("oh", "oh") 'ou
End If
If MouthHeight = 64 And bMouthWidth = 0 Then
imgIdle.Picture = imgLSpeak.Overlay("r", "r")
End If
If MouthHeight = 16 And bMouthWidth = 208 Then
imgIdle.Picture = imgLSpeak.Overlay("dst", "dst") 'x
End If
If MouthHeight = 128 And bMouthWidth = 128 Then
imgIdle.Picture = imgLSpeak.Overlay("i", "i") 'w
End If
If MouthHeight = 243 And bMouthWidth = 0 Then
imgIdle.Picture = imgLSpeak.Overlay("q", "q")
End If
If MouthHeight = 144 And bMouthWidth = 128 Then
imgIdle.Picture = imgLSpeak.Overlay("oh", "oh")
End If
If MouthHeight = 64 And bMouthWidth = 192 Then
imgIdle.Picture = imgLSpeak.Overlay("n", "n")
End If
If MouthHeight = 16 And bMouthWidth = 160 Then
imgIdle.Picture = imgLSpeak.Overlay("fv", "fv")
End If

If MouthHeight = 34 And bMouthWidth = 38 Then
imgIdle.Picture = imgLSpeak.Overlay("dst", "dst")
End If
If MouthHeight = 53 And bMouthWidth = 117 Then
imgIdle.Picture = imgLSpeak.Overlay("k", "k")
End If
If MouthHeight = 112 And bMouthWidth = 176 Then
imgIdle.Picture = imgLSpeak.Overlay("aa", "aa")
End If
If MouthHeight = 179 And bMouthWidth = 0 Then
imgIdle.Picture = imgLSpeak.Overlay("q", "q")
End If
If MouthHeight = 243 And bMouthWidth = 0 Then
imgIdle.Picture = imgLSpeak.Overlay("oh", "oh")
End If
If MouthHeight = 48 And bMouthWidth = 208 Then
imgIdle.Picture = imgLSpeak.Overlay("aa", "aa")
End If
If MouthHeight = 179 And bMouthWidth = 0 Then
imgIdle.Picture = imgLSpeak.Overlay("oh", "oh")
End If
If MouthHeight = 80 And bMouthWidth = 176 Then
imgIdle.Picture = imgLSpeak.Overlay("ee", "ee")
End If

End If

End Sub

Private Sub Timer1_Timer()
imgIdle.Picture = ImageList1.Overlay(j, j)
j = j + 1
If j = 45 Then j = 1
End Sub

Private Sub UserControl_Initialize()
Dim ModeName As String

j = 1
'
    'On Error Resume Next
    engine = DirectSS1.Find("Mfg=Microsoft;Gender=1")
    DirectSS1.Select engine
       For i = 1 To DirectSS1.CountEngines
       LanguageID = DirectSS1.LanguageID(i)
       ModeName = DirectSS1.ModeName(i)
       VoiceCB.AddItem ModeName
       
    Next i
   ' VoiceCB.ListIndex = 0
    VoiceCB.ListIndex = DirectSS1.CurrentMode - 1

Ratesldr.Value = 6

UserControl.Width = 2265
UserControl.Height = 3450 + VoiceCB.Height + MainTxtBox.Height + cmdSpeak.Height + Ratesldr.Height

End Sub
Private Sub UserControl_Resize()
 On Error Resume Next
 UserControl.imgIdle.Width = UserControl.Width
 UserControl.MainTxtBox.Width = UserControl.Width
 UserControl.Ratesldr.Width = UserControl.Width
 UserControl.cmdSpeak.Width = UserControl.Width \ 2
 UserControl.cmdStop.Width = UserControl.Width - UserControl.cmdSpeak.Width
 UserControl.cmdStop.Left = UserControl.cmdSpeak.Width
 UserControl.VoiceCB.Width = UserControl.Width
 UserControl.cmdSpeak.Top = UserControl.Height - UserControl.cmdSpeak.Height
 UserControl.cmdStop.Top = UserControl.Height - UserControl.cmdStop.Height
 UserControl.MainTxtBox.Top = UserControl.Height - UserControl.cmdSpeak.Height - UserControl.MainTxtBox.Height
 UserControl.Ratesldr.Top = UserControl.Height - UserControl.cmdSpeak.Height - UserControl.MainTxtBox.Height - UserControl.Ratesldr.Height
 UserControl.VoiceCB.Top = UserControl.Height - UserControl.cmdSpeak.Height - UserControl.MainTxtBox.Height - UserControl.Ratesldr.Height - UserControl.VoiceCB.Height
 UserControl.imgIdle.Height = UserControl.VoiceCB.Top

End Sub
Private Sub VoiceCB_Click()
QQ = UserControl.VoiceCB.Text

   Dim i As Integer
    
   DirectSS1.CurrentMode = VoiceCB.ListIndex + 1
   DirectSS1.Speak "1 2 3"
   
  
   
End Sub
