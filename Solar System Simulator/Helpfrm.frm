VERSION 5.00
Begin VB.Form Helpfrm 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Helpfrm"
   ClientHeight    =   6480
   ClientLeft      =   150
   ClientTop       =   150
   ClientWidth     =   8625
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   MouseIcon       =   "Helpfrm.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   6480
   ScaleWidth      =   8625
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Label L1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "How to delete files"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   7
      Left            =   5880
      TabIndex        =   9
      Top             =   1440
      Width           =   2535
   End
   Begin VB.Label helplbl 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   3975
      Left            =   3120
      TabIndex        =   8
      Top             =   1800
      Width           =   5295
      WordWrap        =   -1  'True
   End
   Begin VB.Label L1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "What do the statistics mean"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   3
      Left            =   3120
      TabIndex        =   7
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label L1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "How to edit planets"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   4
      Left            =   5760
      TabIndex        =   6
      Top             =   1080
      Width           =   2655
   End
   Begin VB.Label L1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "How to create a new solar system"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   5
      Left            =   3120
      TabIndex        =   5
      Top             =   1440
      Width           =   2655
   End
   Begin VB.Label L1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "How to open"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   2
      Left            =   6720
      TabIndex        =   4
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label L1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "How to save"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   1
      Left            =   5160
      TabIndex        =   3
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label L1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   0
      Left            =   3120
      TabIndex        =   2
      Top             =   720
      Width           =   1935
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   4
      X1              =   3000
      X2              =   8520
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   4
      X1              =   3000
      X2              =   3000
      Y1              =   600
      Y2              =   6360
   End
   Begin VB.Line Line3 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   4
      X1              =   8520
      X2              =   3000
      Y1              =   6360
      Y2              =   6360
   End
   Begin VB.Line Line4 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   4
      X1              =   8520
      X2              =   8520
      Y1              =   600
      Y2              =   6360
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EÞs¡lon Solar System Simulator"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   240
      Left            =   3000
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   240
      Width           =   5595
   End
   Begin VB.Label L1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Index           =   6
      Left            =   4680
      TabIndex        =   0
      Top             =   5880
      Width           =   2295
   End
End
Attribute VB_Name = "Helpfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'ESSS
'This is just a general help and information form
'To direct the wayward user or satisfy the curious user

Option Explicit

Private Sub Form_Activate()
    Randomize
    Dim n As Byte, xx As Integer, yy As Integer
    For n = 1 To 254
        xx = Int(Rnd * Me.Width + 1)
        yy = Int(Rnd * Me.Height + 1)
        If xx > Line4.X1 Or xx < Line2.X1 Then PSet (xx, yy), RGB(Int(Rnd * 11) + 245, Int(Rnd * 11) + 245, Int(Rnd * 11) + 245)
        If yy > Line3.Y1 Or yy < Line1.Y1 Then PSet (xx, yy), RGB(Int(Rnd * 11) + 245, Int(Rnd * 11) + 245, Int(Rnd * 11) + 245)
        xx = Int(Rnd * Me.Width + 1)
        yy = Int(Rnd * Me.Height + 1)
        If xx > Line4.X1 Or xx < Line2.X1 Then PSet (xx, yy), RGB(Int(Rnd * 11) + 245, Int(Rnd * 11) + 245, Int(Rnd * 11) + 245)
        If yy > Line3.Y1 Or yy < Line1.Y1 Then PSet (xx, yy), RGB(Int(Rnd * 11) + 245, Int(Rnd * 11) + 245, Int(Rnd * 11) + 245)
    Next n
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'shortcut keys:
    If KeyAscii = Asc("o") Then Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    TopToolBarfrm.Show
End Sub
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'Buton Click-----------------------------------------------
'//////////////////////////////////////////////////////////
Private Sub L1_Click(Index As Integer)
    Select Case Index
    Case 0:
        helplbl.Caption = "Epsilon Solar System Simulator" & vbCrLf & "Made by Brian Adriance" & vbCrLf & "The solar system simulator is meant to be some thing which you can 'play' almost like a game but has a little more of a realistic feel." & vbCrLf & "If you find any bugs or have any questions or comments, you can e-mail me at rba@valstar.net, make sure that your e-mail's subject is about vb, or regarding esss (otherwise it may be deleted)."
    Case 1:
        helplbl.Caption = "To save your solar system click the save button, then choose the file directory in which you want to save the file. Finnaly give your solar system a name and click save."
    Case 2:
        helplbl.Caption = "To open a solar system, go to the file directory in which you have a previously saved file, it will appear in the file list box. Select it then click open."
    Case 3:
        helplbl.Caption = "The statistics give you some way to 'keep score'." & vbCrLf & "Age:  The age tells how long your solar system has existed." & vbCrLf & "Life:   Sometimes, if you do it right, your solar system may develop life, it can either struggle, thrive, or be non-existant." & vbCrLf & "Orbital Stability:    The orbital stability all depends on the size of your star and number of planets." & vbCrLf & "Star Life Phase:   This simply tells you the phase your star is currently in." & vbCrLf & "Asteroid Threat:   This just tells you if you have asteroids floating around or not." & vbCrLf & "Nova Threat:   Tells you if your star is going to explode soon."
    Case 4:
        helplbl.Caption = "You first edit your planets by clicking on the planet you want to edit. Then find the 'Planet Edit' button on the tool bar and click it. You then will be able to change the type of planet it is, its chemical composition, and its orbital plane."
    Case 5:
        helplbl.Caption = "Creating a solar system is simple; to begin click the 'Create New' button. Then a set of choices will appear. Choose the size of the star, then the color, then the number of planets you want. After, you can choose if you want an asteoid belt and/or a comet in your solar system."
    Case 6:
        Unload Me
    Case 7:
        helplbl.Caption = "You can either delete it normally,  or use ESSS to delete it. To delete a file using ESSS, simply find the file you want to delete,select it and press 'Delete'. To find the 'Delete' button, click the 'Open' button from the main screen."
    End Select
End Sub
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'Highlight-------------------------------------------------
'//////////////////////////////////////////////////////////
Private Sub L1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UnhighlightH
    L1(Index).ForeColor = &H404040: L1(Index).BackColor = &HFF00&
End Sub

Private Sub helplbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UnhighlightH
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UnhighlightH
End Sub

Sub UnhighlightH()
    Dim n As Integer
    For n = 0 To 7
        L1(n).ForeColor = &HFF00&: L1(n).BackColor = &H404040:
    Next n
End Sub


