VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{CFCDAA00-8BE4-11CF-B84B-0020AFBBCCFA}#1.0#0"; "RMOC3260.DLL"
Begin VB.Form Paint 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000004&
   Caption         =   "Ryan's Paint"
   ClientHeight    =   7980
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11880
   FillStyle       =   7  'Diagonal Cross
   LinkTopic       =   "Form1"
   MouseIcon       =   "Paint.frx":0000
   ScaleHeight     =   532
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   792
   Begin VB.Frame Frame1 
      Caption         =   "MP3 Player"
      Height          =   1470
      Left            =   2280
      TabIndex        =   47
      Top             =   6480
      Width           =   6255
      Begin VB.CommandButton Command9 
         Height          =   285
         Left            =   5250
         Picture         =   "Paint.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton PauseRA 
         Height          =   285
         Left            =   4635
         Picture         =   "Paint.frx":0784
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton stop 
         Height          =   285
         Left            =   4035
         Picture         =   "Paint.frx":097E
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   240
         Width           =   615
      End
      Begin VB.Frame Frame2 
         Caption         =   "Information"
         Height          =   1215
         Left            =   75
         TabIndex        =   53
         Top             =   180
         Width           =   2895
         Begin VB.Label Label4 
            Height          =   255
            Left            =   120
            TabIndex        =   56
            Top             =   525
            Width           =   2655
         End
         Begin VB.Label Label3 
            Height          =   255
            Left            =   120
            TabIndex        =   55
            Top             =   300
            Width           =   2655
         End
         Begin VB.Label Label12 
            Height          =   255
            Left            =   120
            TabIndex        =   54
            Top             =   765
            Width           =   2655
         End
      End
      Begin VB.CommandButton play 
         Height          =   285
         Left            =   3435
         Picture         =   "Paint.frx":0B78
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   240
         Width           =   615
      End
      Begin MSComctlLib.ProgressBar Progress 
         Height          =   255
         Left            =   3195
         TabIndex        =   48
         Top             =   600
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   4080
         TabIndex        =   58
         Top             =   885
         Width           =   1095
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Height          =   270
         Left            =   3435
         TabIndex        =   57
         Top             =   1110
         Width           =   2415
      End
   End
   Begin VB.Timer Timer3 
      Interval        =   100
      Left            =   3660
      Top             =   6675
   End
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   9855
      Top             =   -180
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RealAudioObjectsCtl.RealAudio RealAudio 
      Height          =   30
      Left            =   11625
      TabIndex        =   46
      Top             =   7695
      Visible         =   0   'False
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   53
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
   Begin VB.PictureBox Picture3 
      Height          =   15
      Left            =   10560
      ScaleHeight     =   15
      ScaleWidth      =   15
      TabIndex        =   44
      Top             =   405
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000009&
      CausesValidation=   0   'False
      DrawWidth       =   3
      Height          =   6135
      Left            =   2280
      ScaleHeight     =   405
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   414
      TabIndex        =   43
      Top             =   0
      Width           =   6270
   End
   Begin VB.CommandButton Command3 
      Height          =   375
      Left            =   735
      Picture         =   "Paint.frx":0D72
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   1680
      Width           =   735
   End
   Begin VB.ListBox List2 
      Height          =   255
      Left            =   9660
      TabIndex        =   40
      Top             =   -75
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Play Once"
      Height          =   255
      Left            =   10245
      TabIndex        =   39
      Top             =   5205
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Loop Play"
      Height          =   255
      Left            =   9045
      TabIndex        =   38
      Top             =   5220
      Width           =   1095
   End
   Begin VB.PictureBox Picturebox 
      BackColor       =   &H80000009&
      Height          =   0
      Left            =   9300
      ScaleHeight     =   0
      ScaleWidth      =   495
      TabIndex        =   37
      Top             =   45
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox picundo 
      BackColor       =   &H80000009&
      Height          =   0
      Left            =   9300
      ScaleHeight     =   0
      ScaleWidth      =   615
      TabIndex        =   36
      Top             =   330
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ComboBox Fillstyle2 
      Height          =   315
      ItemData        =   "Paint.frx":1264
      Left            =   360
      List            =   "Paint.frx":126E
      TabIndex        =   29
      Text            =   "None"
      ToolTipText     =   "Fill Style"
      Top             =   6120
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Height          =   375
      Left            =   1080
      Picture         =   "Paint.frx":127F
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Ellipses"
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton Command15 
      Height          =   375
      Left            =   1080
      Picture         =   "Paint.frx":1771
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Connected Lines"
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   1080
      Picture         =   "Paint.frx":1BEB
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Rectangles"
      Top             =   600
      Width           =   735
   End
   Begin VB.CommandButton Command8 
      Height          =   375
      Left            =   1080
      Picture         =   "Paint.frx":2065
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Lines"
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Height          =   375
      Left            =   360
      Picture         =   "Paint.frx":253F
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Spray Paint"
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton Command16 
      Height          =   375
      Left            =   360
      Picture         =   "Paint.frx":2A31
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Paintbrush"
      Top             =   960
      Width           =   735
   End
   Begin VB.PictureBox picprint 
      Height          =   255
      Left            =   9870
      ScaleHeight     =   195
      ScaleWidth      =   435
      TabIndex        =   27
      Top             =   60
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Paint.frx":2FC3
      Left            =   360
      List            =   "Paint.frx":2FDC
      TabIndex        =   25
      Text            =   "Solid"
      ToolTipText     =   "Draw Style"
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Remove"
      Height          =   255
      Left            =   10200
      TabIndex        =   2
      ToolTipText     =   "Remove Item"
      Top             =   6600
      Width           =   975
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Add"
      Height          =   255
      Left            =   10200
      TabIndex        =   3
      ToolTipText     =   "Add Item"
      Top             =   6360
      Width           =   975
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   10740
      Top             =   -435
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   10260
      Top             =   -435
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command13 
      Height          =   255
      Left            =   9000
      Picture         =   "Paint.frx":3025
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Stop"
      Top             =   6600
      Width           =   975
   End
   Begin VB.CommandButton Command12 
      Height          =   255
      Left            =   9000
      Picture         =   "Paint.frx":321F
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Play"
      Top             =   6360
      Width           =   975
   End
   Begin VB.CommandButton Command20 
      Height          =   255
      Left            =   9000
      Picture         =   "Paint.frx":3419
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Pause"
      Top             =   6360
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   9720
      TabIndex        =   22
      Text            =   "100"
      ToolTipText     =   "Time Between Frames"
      Top             =   5760
      Width           =   735
   End
   Begin VB.ListBox List 
      Height          =   4155
      ItemData        =   "Paint.frx":3613
      Left            =   8715
      List            =   "Paint.frx":3615
      TabIndex        =   20
      Top             =   645
      Width           =   2865
   End
   Begin MSComctlLib.Slider Slider 
      Height          =   255
      Left            =   360
      TabIndex        =   19
      ToolTipText     =   "Draw Width"
      Top             =   3000
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   393216
   End
   Begin VB.PictureBox linebox 
      BackColor       =   &H80000009&
      Height          =   495
      Left            =   360
      ScaleHeight     =   435
      ScaleWidth      =   1395
      TabIndex        =   18
      Top             =   3360
      Width           =   1455
   End
   Begin VB.PictureBox chosen3 
      BackColor       =   &H80000009&
      Height          =   375
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   17
      ToolTipText     =   "Fill Color"
      Top             =   5160
      Width           =   375
   End
   Begin VB.PictureBox chosen2 
      BackColor       =   &H80000009&
      Height          =   375
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   16
      ToolTipText     =   "BackGround Color"
      Top             =   4800
      Width           =   375
   End
   Begin VB.PictureBox chosen 
      BackColor       =   &H80000006&
      Height          =   375
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   15
      ToolTipText     =   "Line Color"
      Top             =   4440
      Width           =   375
   End
   Begin VB.CommandButton Fill 
      Caption         =   "Set Fill Color"
      Height          =   375
      Left            =   840
      TabIndex        =   14
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CommandButton back 
      Caption         =   "Set Back Color"
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CommandButton line 
      Caption         =   "Set Line Color"
      Height          =   375
      Left            =   840
      TabIndex        =   13
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Clear"
      Height          =   255
      Left            =   9600
      TabIndex        =   11
      ToolTipText     =   "Clear List"
      Top             =   6120
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   840
      Top             =   3840
   End
   Begin VB.CommandButton Command6 
      Height          =   375
      Left            =   360
      Picture         =   "Paint.frx":3617
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Eraser"
      Top             =   600
      Width           =   735
   End
   Begin VB.CommandButton Command5 
      Height          =   375
      Left            =   360
      Picture         =   "Paint.frx":39B9
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Freehand"
      Top             =   240
      Width           =   735
   End
   Begin VB.Frame Animation 
      Caption         =   "Animation List"
      Height          =   4500
      Left            =   8610
      TabIndex        =   35
      Top             =   405
      Width           =   3105
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   4335
      TabIndex        =   45
      Top             =   6210
      Width           =   2415
   End
   Begin VB.Label Frame 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   9495
      TabIndex        =   41
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label11 
      Caption         =   "Label11"
      Height          =   135
      Left            =   9900
      TabIndex        =   34
      Top             =   165
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label10 
      Caption         =   "Label10"
      Height          =   135
      Left            =   9900
      TabIndex        =   33
      Top             =   -75
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      Height          =   30
      Left            =   11280
      TabIndex        =   32
      Top             =   330
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Label8 
      Caption         =   "Label8"
      Height          =   135
      Left            =   11100
      TabIndex        =   31
      Top             =   -75
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Fill Style"
      Height          =   255
      Left            =   240
      TabIndex        =   30
      Top             =   5880
      Width           =   1695
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Draw Style"
      Height          =   255
      Left            =   360
      TabIndex        =   26
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   7590
      TabIndex        =   24
      Top             =   6210
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Timer Interval"
      Height          =   255
      Left            =   9600
      TabIndex        =   21
      Top             =   5520
      Width           =   975
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu Pictures 
         Caption         =   "Pictures"
         Begin VB.Menu New 
            Caption         =   "New"
         End
         Begin VB.Menu Load 
            Caption         =   "Load"
         End
         Begin VB.Menu save 
            Caption         =   "Save"
         End
         Begin VB.Menu Print 
            Caption         =   "Print"
         End
      End
      Begin VB.Menu Animations 
         Caption         =   "Animations"
         Begin VB.Menu Loada 
            Caption         =   "Load"
         End
         Begin VB.Menu savea 
            Caption         =   "Save"
         End
      End
      Begin VB.Menu Quit 
         Caption         =   "Quit"
      End
   End
   Begin VB.Menu Edit 
      Caption         =   "Edit"
      Begin VB.Menu Undo 
         Caption         =   "Undo"
      End
   End
   Begin VB.Menu MP3 
      Caption         =   "MP3 Player"
      Begin VB.Menu Preferences 
         Caption         =   "Preferences"
      End
   End
   Begin VB.Menu Help 
      Caption         =   "Help"
      Begin VB.Menu About 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Paint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ExtFloodFill Lib "Gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long
Private Declare Function Ellipse Lib "Gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Dim number2 As Integer
Dim startx As Single
Dim starty As Single
Dim z As Integer
Dim blnmouseisdown As Boolean
Dim a As Integer
Dim somevar
Dim oldx As Single
Dim oldy As Single
Dim n As Integer
Dim o As Integer
Dim mystr As String
Dim intupperbound As Integer
Dim intlowerbound As Integer
Dim intxcoord As Integer
Dim intycoord As Integer
Dim intradius As Integer
Dim namey As String
Dim q As Integer
Dim astring As String
Dim r As Integer
Dim timetopause As Integer
Dim anothervariable As Integer
Dim xx As Integer
Dim slidervar As Integer
Dim oldx1 As Variant
Dim oldx2 As Variant
Dim oldy1 As Variant
Dim oldy2 As Variant
Dim avar As Integer
Dim string234 As String
Dim aye As String
Dim number As String
Dim aye2 As String
Dim temp
Dim Picturesbln As Boolean
Dim strstring As String
Dim timevariable
Dim timevar
Dim ro
Dim progressvar


Public Sub Remove()

    If List.ListCount = 0 Then Exit Sub
    If avar > List.ListCount - 1 Then List.ListIndex = avar - 1
    If avar <= List.ListCount - 1 Then List.ListIndex = avar
    

End Sub



Private Sub Pause(pausetime)
On Error GoTo error
    For i = 0 To pausetime
    DoEvents
    Next i
    
error:
End Sub


Private Sub About_Click()

    frmAbout.Label2.Caption = "d"
    frmAbout.Visible = True
    Me.Enabled = False

End Sub

Private Sub back_Click()
On Error GoTo error
    CommonDialog.ShowColor
    chosen2.BackColor = CommonDialog.Color
    Picture1 = LoadPicture
    Picture1.BackColor = chosen2.BackColor
error:
End Sub

Private Sub comm_Click()

On Error GoTo lblcomm
    Picture1 = LoadPicture(filetext.Text)
lblcomm:
End Sub

Private Sub chosen_DblClick()

On Error GoTo chosenerror
        
    CommonDialog.ShowColor
    chosen.BackColor = CommonDialog.Color
    Picture1.ForeColor = chosen.BackColor
    linebox.Cls
    linebox.Line (0, 250)-(1350, 250)
    
    
chosenerror:

End Sub

Private Sub chosen_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = 2 Then Call chosen_DblClick

End Sub

Private Sub chosen2_DblClick()

On Error GoTo chosenerror

    CommonDialog.ShowColor
    chosen2.BackColor = CommonDialog.Color
    Picture1 = LoadPicture
    Picture1.Cls
    Picture1.BackColor = chosen2.BackColor
chosenerror:
End Sub

Private Sub chosen2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = 2 Then Call chosen2_DblClick

End Sub

Private Sub chosen3_DblClick()

On Error GoTo chosenerror

    If Picture1.FillStyle <> 1 Then
    CommonDialog.ShowColor
    chosen3.BackColor = CommonDialog.Color
    Picture1.FillColor = chosen3.BackColor
    End If


chosenerror:
End Sub

Private Sub chosen3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = 2 Then Call chosen3_DblClick

End Sub

Private Sub Combo1_Change()

    Combo1.Text = aye
    
End Sub

Private Sub Combo1_Click()

    Picture1.DrawWidth = 1
    Picture1.DrawStyle = Combo1.ListIndex
    linebox.DrawStyle = Combo1.ListIndex
    linebox.DrawWidth = 1
    Slider.Value = 1
    linebox.Cls
    linebox.Line (0, 250)-(1350, 250)
    aye = Combo1.Text
    number = Combo1.ListIndex
    

End Sub

Private Sub Command1_Click()

    
    If ro = 1 Then chosen3.BackColor = RGB(255, 255, 255)
    Fillstyle2.ListIndex = ro
    If ro = 1 Then Fillstyle2.Text = "None"
    Picture1.FillStyle = ro
    If ro = 1 Then Fill.Enabled = False
    Picture1.MousePointer = 2
    z = 1

End Sub



Private Sub Command10_Click()


    If List.ListIndex <> -1 Then
    avar = List.ListIndex
    List2.ListIndex = List.ListIndex
    List.RemoveItem List.ListIndex
    List2.RemoveItem List2.ListIndex
    Call Remove
    End If

End Sub




Private Sub Command11_Click()

On Error GoTo lblerror3
    CommonDialog.CancelError = True
    CommonDialog.ShowOpen
    List.AddItem (List.ListCount + 1 & ".  " & CommonDialog.FileName)
    List2.AddItem (CommonDialog.FileName)
    CommonDialog.CancelError = False
lblerror3:
    
    
End Sub

Private Sub Command12_Click()

    If List.ListCount <> 0 Then
    
    
    Timer1.Interval = Text1.Text
    If Timer1.Interval >= 1 Then
    o = 1
    Command12.Visible = False
    Command20.Visible = True
    End If
    End If

End Sub

Private Sub Command13_Click()

    o = 2
    n = 0
    Command20.Visible = False
    Command12.Visible = True
    
    
End Sub

Private Sub Command15_Click()

    If ro = 1 Then chosen3.BackColor = RGB(255, 255, 255)
    Fillstyle2.ListIndex = ro
    If ro = 1 Then Fillstyle2.Text = "None"
    Picture1.FillStyle = ro
    If ro = 1 Then Fill.Enabled = False
    'If z = 55 Then Picture1.Picture = Picture2.Image
    'Picture2.Visible = False
    'Picture1.Visible = True
    Picture1.MousePointer = 2
    xx = 1
    z = 27

End Sub

Private Sub Command16_Click()

    If ro = 1 Then chosen3.BackColor = RGB(255, 255, 255)
    Fillstyle2.ListIndex = ro
    If ro = 1 Then Fillstyle2.Text = "None"
    Picture1.FillStyle = ro
    If ro = 1 Then Fill.Enabled = False
    Picture1.MousePointer = 2
    z = 32

End Sub

Private Sub Command2_Click()
    
    If ro = 1 Then chosen3.BackColor = RGB(255, 255, 255)
    Fillstyle2.ListIndex = ro
    If ro = 1 Then Fillstyle2.Text = "None"
    Picture1.FillStyle = ro
    If ro = 1 Then Fill.Enabled = False
    Picture1.MousePointer = 2
   
    z = 2

End Sub

Private Sub Command20_Click()

    o = 2
    Command12.Visible = True
    Command20.Visible = False

End Sub

Private Sub Command3_Click()
On Error GoTo error
    ro = Picture1.FillStyle
    z = 55
    Fillstyle2.ListIndex = 0
    Picture1.MousePointer = 99
    Picture1.MouseIcon = LoadPicture("C:\Program Files\Microsoft Visual Studio\VB98\FillRgn.cur")
    
    
    Picture1.FillStyle = 0
    Picture1.FillColor = chosen3.BackColor
    Fill.Enabled = True
    Exit Sub

error:
    ro = Picture1.FillStyle
    z = 55
    Fillstyle2.ListIndex = 0
    Picture1.MousePointer = 0
    
    Picture1.FillStyle = 0
    Picture1.FillColor = chosen3.BackColor
    Fill.Enabled = True

End Sub

Private Sub Command4_Click()
On Error GoTo error
    
    If ro = 1 Then chosen3.BackColor = RGB(255, 255, 255)
    Fillstyle2.ListIndex = ro
    If ro = 1 Then Fillstyle2.Text = "None"
    Picture1.FillStyle = ro
    If ro = 1 Then Fill.Enabled = False

    Picture1.MousePointer = 99
    Picture1.MouseIcon = LoadPicture("C:\Program Files\Microsoft Visual Studio\VB98\brush.cur")
    
    z = 4
    Exit Sub
error:
    If ro = 1 Then chosen3.BackColor = RGB(255, 255, 255)
    Fillstyle2.ListIndex = ro
    Picture1.FillStyle = ro
        If ro = 1 Then Fillstyle2.Text = "None"

    z = 4
    Picture1.MousePointer = 0


End Sub

Private Sub Command5_Click()
On Error GoTo error
    
    Fillstyle2.ListIndex = ro
    Picture1.FillStyle = ro
    If ro = 1 Then Fill.Enabled = False
    If ro = 1 Then chosen3.BackColor = RGB(255, 255, 255)
   
    Picture1.MousePointer = 99
    Picture1.MouseIcon = LoadPicture("C:\Program Files\Microsoft Visual Studio\VB98\pencil.cur")
    z = 5
    Picture1.DrawWidth = 1
    Slider.Value = 1
    linebox.Line (0, 250)-(1455, 250)
    Exit Sub
    
error:
    Picture1.FillStyle = ro
    If ro = 1 Then chosen3.BackColor = RGB(255, 255, 255)
    Fillstyle2.ListIndex = ro
    If ro = 1 Then Fill.Enabled = False
    z = 5
    Picture1.DrawWidth = 1
    Picture1.MousePointer = 0
    Slider.Value = 1
    linebox.Line (0, 250)-(1455, 250)

End Sub

Private Sub Command6_Click()
On Error GoTo error
    If ro = 1 Then chosen3.BackColor = RGB(255, 255, 255)
    Fillstyle2.ListIndex = ro
    If ro = 1 Then Fillstyle2.Text = "None"
    Picture1.FillStyle = ro
    If ro = 1 Then Fill.Enabled = False

    Picture1.MousePointer = 99
    Picture1.MouseIcon = LoadPicture("C:\Program Files\Microsoft Visual Studio\VB98\erase.cur")
    
    z = 10
    Exit Sub
error:
    If ro = 1 Then chosen3.BackColor = RGB(255, 255, 255)
    Fillstyle2.ListIndex = ro
    Picture1.FillStyle = ro
    If ro = 1 Then Fillstyle2.Text = "None"
    If ro = 1 Then Fill.Enabled = False
  
    z = 10
    Picture1.MousePointer = 0

End Sub

Private Sub Command7_Click()

    List.Clear
    List2.Clear

End Sub

Private Sub Command8_Click()
    
    Fillstyle2.ListIndex = ro
    Picture1.FillStyle = ro
    If ro = 1 Then chosen3.BackColor = RGB(255, 255, 255)
    If ro = 1 Then Fill.Enabled = False
    'Picture2.Visible = False
    'Picture1.Visible = True
    'If z = 55 Then Picture1.Picture = Picture2.Image
    Picture1.MousePointer = 2
    z = 8

End Sub




Private Sub default_Click()
    
    Picturesbln = True
    frmDefault.Visible = True
    Paint.Enabled = False
       

End Sub

Private Sub defaulta_Click()
    
    Picturesbln = False
    frmDefault.Visible = True
    Paint.Enabled = False


End Sub

Private Sub Command9_Click()

On Error GoTo error
    CommonDialog2.ShowOpen
    RealAudio.SetSource (CommonDialog2.FileName)
    
error:

End Sub

Private Sub Disabled_Click()

    Disabled.Caption = "Disabled"
    Enabled2.Caption = "Enable"
    Disabled.Checked = True
    Enabled2.Checked = False
    

End Sub

Private Sub Enabled2_Click()

    Disabled.Caption = "Disable"
    Enabled2.Caption = "Enabled"

    Enabled2.Checked = True
    Disabled.Checked = False

End Sub

Private Sub fill_Click()
On Error GoTo error
    
    CommonDialog.ShowColor
    chosen3.BackColor = CommonDialog.Color
    Picture1.FillColor = chosen3.BackColor
    
    
error:
End Sub

Private Sub Fillstyle2_Change()

    Fillstyle2.Text = aye2

End Sub

Private Sub Fillstyle2_Click()

On Error GoTo fillstyleerror
    
    Picture1.FillStyle = Fillstyle2.ListIndex
    aye2 = Fillstyle2.Text
    If Fillstyle2.ListIndex <> 1 Then Fill.Enabled = True
    If Fillstyle2.ListIndex = 1 Then Fill.Enabled = False
    number2 = Fillstyle2.ListIndex
    If z <> 55 Then ro = Picture1.FillStyle
   
   
fillstyleerror:

End Sub

Private Sub Form_Load()
    
        
    On Error GoTo error
        Open "C:\mp3pref.txt" For Input As #1
        
        Do Until EOF(1)
        Input #1, astring
               
        CommonDialog2.InitDir = astring
                
        Loop
        Close (1)
        GoTo skip
error:
    CommonDialog2.InitDir = "C:\"
skip:
    progressvar = 2
    ro = 1
    'CommonDialog2.InitDir = "C:\Program Files\"
    CommonDialog2.Filter = "MP3s|*.mp3"
        
 
    CommonDialog2.CancelError = True



    Picture1.FillStyle = 1
    Picture3.Height = Picture1.Height
    Picture3.Width = Picture1.Width
    Option1 = True
    number2 = 1
    aye = "Solid"
    aye2 = "None"
    Combo1.ListIndex = 0
    CommonDialog.CancelError = True
    CommonDialog1.CancelError = True
    CommonDialog.Filter = "Pictures|*.bmp;*.jpg;*.jpeg"
    CommonDialog.InitDir = "C:\My Documents"
    CommonDialog.Flags = cdlOFNFileMustExist + cdlOFNOverwritePrompt + cdlOFNPathMustExist
    CommonDialog1.InitDir = "C:\My Documents"
    CommonDialog1.Flags = CommonDialog.Flags
    CommonDialog1.Filter = "Text Files|*.txt"
    
    Picture1.ForeColor = chosen.BackColor
    Slider.Min = 1
    Slider.Value = 1
    intradius = 100
    Command20.Visible = False
    Text1.Text = Timer1.Interval
    
    o = 0
    n = 0
    Fill.Enabled = False
    'Command3.Visible = False
    'linebox.Width = 1455
    'linebox.Height = 500
    
    linebox.Cls
    linebox.DrawWidth = 1
    linebox.Line (0, 250)-(1455, 250)
    

    Picture1.FillColor = Picture1.BackColor


    Picture1.ScaleHeight = Picture1.Height
    Picture1.ScaleWidth = Picture1.Width

    Paint.FillStyle = 1
    Paint.ScaleHeight = 1000
    Paint.ScaleWidth = 1000
    
    r = 1

End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    Label5.Caption = ""

End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    Label5.Caption = ""

End Sub

Private Sub line_Click()
On Error GoTo error
    
    CommonDialog.ShowColor
    
    chosen.BackColor = CommonDialog.Color
    
    Picture1.ForeColor = chosen.BackColor
    linebox.ForeColor = chosen.BackColor
    linebox.Line (0, 250)-(1455, 250)
    
error:
End Sub

Private Sub linebox_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    Label5.Caption = ""

End Sub

Private Sub List_DblClick()

    List2.ListIndex = List.ListIndex
    Picture1 = LoadPicture(List2.Text)

End Sub

Private Sub List_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyDelete Then Call Command10_Click

End Sub

Private Sub List_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If List.ListCount <> 0 Then
    List.ToolTipText = "Double Click to View Picture"
    End If
    If List.ListCount = 0 Then
    List.ToolTipText = ""
    End If
    Label5.Caption = ""

End Sub

Private Sub Load_Click()
On Error GoTo lblerror2
    CommonDialog.CancelError = True
    CommonDialog.ShowOpen
    Picture1 = LoadPicture(CommonDialog.FileName)
    CommonDialog.CancelError = False
lblerror2:

End Sub

Private Sub Loada_Click()

On Error GoTo anothererror
    CommonDialog1.ShowOpen
        List.Clear
        List2.Clear
        Open CommonDialog1.FileName For Input As #1
        
        Do Until EOF(1)
        Input #1, astring
               
        List2.AddItem (astring)
        List.AddItem (List.ListCount + 1 & ".  " & astring)
        Loop
        
    
    Close (1)
    

anothererror:


End Sub

Private Sub New_Click()

    Picture1 = LoadPicture
    Picture1.Cls
    
    Picture1.BackColor = chosen2.BackColor

End Sub

Private Sub Pause_Click()

    RealAudio.DoPause

End Sub

Private Sub PauseRA_Click()

    RealAudio.DoPause

End Sub

Private Sub Picture1_DblClick()

    If z <> 30 Then
    q = z
    z = 30
    End If

End Sub



Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    
    On Error GoTo error
    
    
    picundo.Height = Picture1.Height
    picundo.Width = Picture1.Width
    picundo.Picture = Picture1.Image
    
       
    If Button = 1 And z = 55 Then
    'build a random fill color
    'Randomize
    Picture1.FillColor = chosen3.BackColor
    'call the dll using a temporary variable
    ExtFloodFill Picture1.hdc, x, y, Picture1.Point(x, y), 1
    End If
    
    
        If Button = 1 Then
         blnmouseisdown = True
    End If
    
        If z = 4 Then
    
        Picture1.DrawMode = 13
        Picture1.DrawWidth = 1
        intradius = 8 + (2 * Slider.Value)
        timetopause = -intradius + 135
        
        Do While blnmouseisdown
        Pause timetopause
        If blnmouseisdown = False Then
        Exit Sub
        End If
        DoEvents

        intxcoord = Int((((x + intradius) - (x - intradius)) * Rnd + 1) + (x - intradius))
        
        intupperbound = y + Int(Sqr(Abs(intradius ^ 2 - (intxcoord - x) ^ 2)))
        intlowerbound = y - Int(Sqr(Abs(intradius ^ 2 - (intxcoord - x) ^ 2)))
        intycoord = Int(((intupperbound - intlowerbound) * Rnd + 1) + intlowerbound)
        
        Picture1.PSet (intxcoord, intycoord)
        'Pause timetopause
        Loop
       
    
    
    End If
    
    
    
    
    If number2 >= 2 Then
    Picture1.FillStyle = number2
    Fill.Enabled = True
    End If
    If number2 = 0 Then
    Picture1.FillStyle = 0
    Fill.Enabled = True
    End If
    If number2 = 1 Then
    Picture1.FillStyle = 1
    Fill.Enabled = False
    End If
    If z = 55 Then
    Fill.Enabled = True
    End If
    
    If z <> 55 Then ro = Picture1.FillStyle
    Picture1.FillStyle = 1
    
    If z = 30 Then
    z = q
    End If
    Picture1.AutoRedraw = True
    Picture1.DrawMode = 2
    
    If z <> 27 Then
    startx = x
    starty = y
    End If

    oldx = x
    oldy = y
        
    If z = 27 And xx = 1 Then
    startx = x
    starty = y
    End If

    anothervariable = 1
    

    
    If z = 23 Then Picture1.DrawStyle = 2
    If z = 24 Then Picture1.DrawStyle = 2

    If Button = 2 And z = 27 Then z = -132
    
    Label8.Caption = x
    Label9.Caption = y
    Label10.Caption = x
    Label11.Caption = y
        
        X1 = Label8.Caption / 45 * 3.14
        Y1 = Label9.Caption / 45 * 3.14
        X2 = Label10.Caption / 45 * 3.14
        Y2 = Label11.Caption / 45 * 3.14
    
    
        oldx1 = X1
        oldx2 = X2
        oldy1 = Y1
        oldy2 = Y2
        
    If z = 5 Then Picture1.DrawWidth = 1
    Exit Sub

error:
    blnmouseisdown = False
    blnmouseisdown = True
    Me.Cls

End Sub


Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)


On Error GoTo error

    
    
    somevar = 1
    
    Label10.Caption = x
    Label11.Caption = y

    Label5.Caption = "" & Round(x, 0) & ", " & Round(y, 0)
    
    If z = 32 And Button = 1 Then
    
        Picture1.DrawMode = 13
        Picture1.Line (startx, starty)-(x, y)
    
        startx = x
        starty = y
    End If
    

        
    If z = 27 And Button = 1 Then
        
        Picture1.ForeColor = &H0&
        Picture1.Line (startx, starty)-(x, y)
        Picture1.Line (startx, starty)-(oldx, oldy)
        If anothervariable = 1 Then Picture1.Line (startx, starty)-(oldx, oldy)
        anothervariable = 2
        Label2.Caption = "Length: " & Round(Sqr((startx - x) ^ 2 + (starty - y) ^ 2))
        oldx = x
        oldy = y

     End If
    
    
    
    
    If z = 23 And Button = 1 Then
    
        Picture1.ForeColor = &H0&
        Picture1.DrawWidth = 1
        Picture1.Line (startx, starty)-(x, y), , B
        Picture1.Line (startx, starty)-(oldx, oldy), , B
        
        
        
        oldx = x
        oldy = y
        
        
    End If
    
    
    If z = 1 And Button = 1 Then
        
        Picture1.ForeColor = &H0&
        
        Picture1.Line (startx, starty)-(x, y), , B
        Picture1.Line (startx, starty)-(oldx, oldy), , B
        
        Label2.Caption = "Height: " & Round(Abs(y - starty), 0) & ", Width: " & Round(Abs(x - startx), 0)
    
    
        oldx = x
        oldy = y
    
    End If
    
    If z = 2 And Button = 1 Then
    
        Picture1.ForeColor = &H0&
        
        If blnmouseisdown Then
        
        X1 = Label8.Caption / 3 * 3.14
        Y1 = Label9.Caption / 3 * 3.14
        X2 = Label10.Caption / 3 * 3.14
        Y2 = Label11.Caption / 3 * 3.14
        
               
        Call Ellipse(Picture1.hdc, X1, Y1, X2, Y2)
        Call Ellipse(Picture1.hdc, oldx1, oldy1, oldx2, oldy2)
        Label2.Caption = "Height: " & Round(Abs(Y2 - Y1)) & ", Width: " & Round(Abs(X2 - X1))
        Picture1.Refresh
       
        
        
        oldx1 = X1
        oldx2 = X2
        oldy1 = Y1
        oldy2 = Y2
        End If
    
    End If
    
    
    
    
    
    If z = 4 Then
    
        Picture1.DrawMode = 13
        'Picture1.DrawWidth = 1
        intradius = 8 + (2 * Slider.Value)
        timetopause = -intradius + 135
        
        Do While blnmouseisdown
        Pause timetopause
        If blnmouseisdown = False Then
        Exit Sub
        End If
        
        DoEvents

        intxcoord = Int((((x + intradius) - (x - intradius)) * Rnd + 1) + (x - intradius))
        
        intupperbound = y + Int(Sqr(Abs(intradius ^ 2 - (intxcoord - x) ^ 2)))
        intlowerbound = y - Int(Sqr(Abs(intradius ^ 2 - (intxcoord - x) ^ 2)))
        intycoord = Int(((intupperbound - intlowerbound) * Rnd + 1) + intlowerbound)
        
        Picture1.PSet (intxcoord, intycoord)
        'Pause timetopause
        Loop
        
    
    
    End If
 
    
        If z = 5 Then
        If Button = 1 Then
    
        Picture1.DrawMode = 13
        Picture1.Line (startx, starty)-(x, y)
    
        startx = x
        starty = y
    
    
    End If
    
    
    
    If Button = 2 Then
        
        Picture1.DrawMode = 13
        Picture1.DrawWidth = Slider.Value
        Picture1.ForeColor = chosen2.BackColor
    
        Picture1.Line (startx, starty)-(x, y)
    
        startx = x
        starty = y
    
        Picture1.DrawWidth = Slider.Value
        Picture1.ForeColor = chosen.BackColor
    
    End If
    End If
    
    If z = 8 And Button = 1 Then
    
        Picture1.ForeColor = &H0&
        
        Picture1.Line (startx, starty)-(x, y)
        Picture1.Line (startx, starty)-(oldx, oldy)
        Label2.Caption = "Length: " & Round(Sqr((startx - x) ^ 2 + (starty - y) ^ 2))
    
        oldx = x
        oldy = y

     End If
    
    If z = 10 And Button = 1 Then
    
        Picture1.DrawMode = 13
        Picture1.DrawWidth = 10
        Picture1.ForeColor = chosen2.BackColor
    
        Picture1.Line (startx, starty)-(x, y)
    
        startx = x
        starty = y
    
        Picture1.DrawWidth = linebox.DrawWidth
        Picture1.ForeColor = chosen.BackColor
    
    End If
    
    Exit Sub
error:
    blnmouseisdown = False
    blnmouseisdown = True
    Me.Cls
    

End Sub



Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    somevar = 2
     blnmouseisdown = False
     Picture1.DrawWidth = Slider.Value
    Label2.Caption = ""
    If z <> 55 Then Picture1.FillStyle = ro
    If z = 55 Then Picture1.FillStyle = 0
    Picture1.DrawStyle = number
    Picture1.ForeColor = chosen.BackColor
    Picture1.DrawMode = 13
    If z = 27 Then
    Picture1.Line (startx, starty)-(x, y)
    End If
    xx = 2
    If z = 27 Then
    startx = x
    starty = y
    End If
    If z = 23 Then Picture1.DrawStyle = 0
    If z = 23 Then z = 24
    If z = 24 And Picture1.DrawStyle = 2 Then
    z = 23
    Picture1.DrawStyle = 0
    End If
    
    
   
    
    Picture1.DrawWidth = linebox.DrawWidth
    
    If z = 8 Then
    Picture1.Line (startx, starty)-(x, y)
    End If
    If z = 1 Then
    Picture1.Line (startx, starty)-(x, y), , B
    End If
    If z = 2 Then
    Picture1.DrawMode = 13
    
    Call Ellipse(Picture1.hdc, oldx1, oldy1, oldx2, oldy2)
    Picture1.Refresh
 
    
    End If
       

    
    

End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

If Button = 1 Then
    'build a random fill color
    'Randomize
    Picture2.FillColor = chosen.BackColor
    'call the dll using a temporary variable
    ExtFloodFill Picture2.hdc, x, y, Picture2.Point(x, y), 1
End If

End Sub

Private Sub play_Click()

    RealAudio.DoPlay


End Sub

Private Sub Preferences_Click()

    Preferencesmp3.Label2.Caption = "jkl;"
    Paint.Enabled = False
    
    Preferencesmp3.Visible = True

End Sub

Private Sub Print_Click()
On Error GoTo printerror
    frmPrint.Visible = True
    Exit Sub
printerror:
End Sub

Private Sub Progress_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
progressvar = 1
    
    Pause 100
    RealAudio.DoPause
progressvar = 1
    If Button = 1 And RealAudio.GetPlayState = 4 Then
    progressvar = 1
    Progress.Value = Progress.Max * (x / 2820)
    RealAudio.SetPosition (Progress.Max * (x / 2820))
    Progress.Value = Progress.Max * (x / 2820)
    
    End If


End Sub

Private Sub Progress_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    On Error GoTo error
    If Button = 1 And RealAudio.GetPlayState = 4 Then
    progressvar = 1
    RealAudio.SetPosition (Progress.Max * (x / 2820))
    Progress.Value = Progress.Max * (x / 2820)
    
    End If
error:
    

End Sub

Private Sub Progress_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    If RealAudio.GetPlayState = 4 Then
    RealAudio.DoPlay
    
    RealAudio.SetPosition (Progress.Max * (x / 2820))
    End If
    



End Sub

Private Sub Quit_Click()

    End

End Sub



Private Sub Save_Click()

On Error GoTo lblerror1
    CommonDialog.CancelError = True
    CommonDialog.ShowSave
    SavePicture Picture1.Image, CommonDialog.FileName
    CommonDialog.CancelError = False
    
lblerror1:

End Sub

Private Sub savea_Click()
On Error GoTo idontknow
   CommonDialog1.ShowSave
   Open CommonDialog1.FileName For Output As #1
    
        For i = 0 To List.ListCount - 1
        
        Print #1, List2.List(i)
           
        Next i
    
    Close (1)
idontknow:

End Sub

Private Sub Slider_Change()

    If z = 5 Then Slider.Value = 1
    Picture1.DrawWidth = Slider.Value
    linebox.DrawWidth = Slider.Value
    
    
    If linebox.DrawWidth <> 1 And Combo1.ListIndex <> 0 Then
    slidervar = Slider.Value
    Combo1.ListIndex = 0
    Picture1.DrawStyle = 0
    linebox.DrawStyle = 0
    Slider.Value = slidervar
    
    End If
    
    Picture1.DrawWidth = Slider.Value
    linebox.DrawWidth = Slider.Value
    


    linebox.Cls
    linebox.Line (0, 250)-(1455, 250)
    
    

End Sub



Private Sub stop_Click()
    
    Progress.Value = 0
    RealAudio.DoStop
    Progress.Value = 0

End Sub

Private Sub Text1_Change()

On Error GoTo error
    
    If Text1.Text <= 0 Then
    o = 2
    
    Command20.Visible = False
    Command12.Visible = True
    End If
    Exit Sub

error:
Text1.Text = "100"
Timer1.Interval = Text1.Text
End Sub

Private Sub Timer1_Timer()
    
On Error GoTo lblerror

    If o = 2 Then Frame.Caption = ""

    If o = 1 Then
    Timer1.Interval = Text1.Text
    If List2.ListCount = 0 Then
    o = 2
    Command20.Visible = False
    Command12.Visible = True
    End If
    If n = List2.ListCount - 1 And Option2 = True Then
    Command20.Visible = False
    Command12.Visible = True
    o = 2
    n = 0
    End If
    If n > (List2.ListCount - 1) And Option1 = True Then
    n = 0
    End If
    
    Frame.Caption = "Frame: " & n + 1
    mystr = List2.List(n)
    
    Picture1 = LoadPicture(mystr)
    n = n + 1
    End If
Exit Sub

lblerror:
MsgBox "ERROR, Check To See If Correct"
o = 2
    
    Command20.Visible = False
    Command12.Visible = True




End Sub




Private Sub Timer3_Timer()

    
    If z = 4 Then Picture1.DrawWidth = 1
    If z <> 4 Then Picture1.DrawWidth = Slider.Value
    

    If RealAudio.GetPlayState = 3 Or RealAudio.GetPlayState = 4 Then
    Progress.Max = RealAudio.GetLength
    Progress.Value = RealAudio.GetPosition
    Label3.Caption = RealAudio.GetAuthor
    Label4.Caption = RealAudio.GetTitle
    timevariable = Round(RealAudio.GetLength / 60000)
    If timevariable > (RealAudio.GetLength / 60000) Then
    timevariable = timevariable - 1
    End If
    timevar = Round(RealAudio.GetLength / 1000) - (timevariable * 60)
    If timevar >= 10 Then Label12.Caption = "Length: " & timevariable & ":" & timevar
    If timevar < 10 Then Label12.Caption = "Length: " & timevariable & ":" & "0" & timevar
    End If
    If RealAudio.GetPlayState = 0 Then
    Progress.Value = 0
    Label3.Caption = ""
    Label4.Caption = ""
    Label12.Caption = ""
    End If
    If RealAudio.GetPlayState = 0 Then
    Label13.Caption = "Stopped"
    End If
    If RealAudio.GetPlayState = 4 Then
    If progressvar = 2 Then
    Label13.Caption = "Paused"
    End If
    End If
    If RealAudio.GetPlayState = 3 Then
    Label13.Caption = "Playing"
    progressvar = 2
    End If
    If progressvar = 1 And RealAudio.GetPlayState <> 0 Then Label13.Caption = "Choosing Location"
    linebox.Line (0, 250)-(1455, 250)
    If RealAudio.GetPlayState = 3 Or RealAudio.GetPlayState = 4 Then
    timevariable = Round(RealAudio.GetPosition / 60000)
    If timevariable > (RealAudio.GetPosition / 60000) Then
    timevariable = timevariable - 1
    End If
    timevar = Round(RealAudio.GetPosition / 1000) - (timevariable * 60)
    If timevar >= 10 Then Label14.Caption = timevariable & ":" & timevar
    If timevar < 10 Then Label14.Caption = timevariable & ":" & "0" & timevar
    End If
    If RealAudio.GetPlayState = 0 Then Label14.Caption = ""
    
    
    

End Sub

Private Sub Undo_Click()

    Picturebox.Picture = Picture1.Image
    Picture1 = LoadPicture
    Picturebox.Height = Picture1.Height
    Picturebox.Width = Picture1.Width
    
    Picture1.Picture = picundo.Image
    picundo.Picture = Picturebox.Image
End Sub
