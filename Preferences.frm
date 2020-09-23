VERSION 5.00
Object = "{CFCDAA00-8BE4-11CF-B84B-0020AFBBCCFA}#1.0#0"; "RMOC3260.DLL"
Begin VB.Form Preferencesmp3 
   Caption         =   "MP3 Preferences"
   ClientHeight    =   3930
   ClientLeft      =   3645
   ClientTop       =   2670
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3930
   ScaleWidth      =   4680
   Begin RealAudioObjectsCtl.RealAudio RealAudio1 
      Height          =   0
      Left            =   4620
      TabIndex        =   6
      Top             =   3840
      Visible         =   0   'False
      Width           =   0
      _ExtentX        =   0
      _ExtentY        =   0
      AUTOSTART       =   0   'False
      SHUFFLE         =   0   'False
      PREFETCH        =   0   'False
      NOLABELS        =   0   'False
      LOOP            =   0   'False
      NUMLOOP         =   0
      CENTER          =   0   'False
      MAINTAINASPECT  =   0   'False
      BACKGROUNDCOLOR =   "#000000"
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Set Directory"
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   3000
      Width           =   1455
   End
   Begin VB.DirListBox Dir1 
      Height          =   1890
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   4215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "About RealPlayer"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   3480
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   4215
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   15
      Left            =   720
      TabIndex        =   5
      Top             =   3480
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Directory For Opening MP3s:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "Preferencesmp3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    RealAudio1.AboutBox

End Sub

Private Sub Command2_Click()

    Open "C:\mp3pref.txt" For Output As #1
    
        Print #1, Text1.Text
          
    Close (1)
    Paint.CommonDialog2.InitDir = Text1.Text

End Sub

Private Sub Dir1_Change()

    Text1.Text = Dir1.Path

End Sub

Private Sub Form_Load()

        On Error GoTo error
        Open "C:\mp3pref.txt" For Input As #1
        
        Do Until EOF(1)
        Input #1, astring
               
        Text1.Text = astring
        Dir1.Path = astring
                
        Loop
        Close (1)
        GoTo skip
error:
    Text1.Text = "C:\"
    Dir1.Path = "C:\"
skip:
    

End Sub

Private Sub Form_Terminate()
    

    
    Paint.Visible = True
    Paint.Enabled = True

End Sub

Private Sub Form_Unload(Cancel As Integer)

    
    Paint.Visible = True
    Paint.Enabled = True

End Sub
