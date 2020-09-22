VERSION 5.00
Begin VB.Form Questionfrm 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "esssqfrm"
   ClientHeight    =   6570
   ClientLeft      =   585
   ClientTop       =   150
   ClientWidth     =   9690
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   MouseIcon       =   "Questionfrm.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   6570
   ScaleWidth      =   9690
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Label Ca 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "&Cancel"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   34
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label34 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   7680
      TabIndex        =   33
      Top             =   6240
      Width           =   1335
   End
   Begin VB.Label Label33 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   4800
      TabIndex        =   32
      Top             =   6240
      Width           =   1335
   End
   Begin VB.Label Label32 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   2040
      TabIndex        =   31
      Top             =   6240
      Width           =   1335
   End
   Begin VB.Label Label31 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "EÞs¡lon Solar System Simulator"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   495
      Left            =   2040
      TabIndex        =   30
      Top             =   600
      Width           =   7575
   End
   Begin VB.Label OC 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Orbiting Comet"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   6240
      TabIndex        =   29
      ToolTipText     =   "Choose to have or not have a comet in your solar system."
      Top             =   4560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label AB 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Asteroid Belt"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   4200
      TabIndex        =   28
      ToolTipText     =   "Choose to have or not have an asteroid belt in your solar system."
      Top             =   4560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label L1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "&OK"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   24
      Left            =   3960
      TabIndex        =   27
      Top             =   5040
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Label L1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "10"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   23
      Left            =   8760
      TabIndex        =   26
      Top             =   3840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label L1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "9"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   22
      Left            =   7920
      TabIndex        =   25
      Top             =   3840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label L1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "8"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   21
      Left            =   8760
      TabIndex        =   24
      Top             =   3360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label L1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "7"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   20
      Left            =   7920
      TabIndex        =   23
      Top             =   3360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label L1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "6"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   19
      Left            =   8760
      TabIndex        =   22
      Top             =   2880
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label L1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "5"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   18
      Left            =   7920
      TabIndex        =   21
      Top             =   2880
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label L1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "4"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   17
      Left            =   8760
      TabIndex        =   20
      Top             =   2400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label L1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "3"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   16
      Left            =   7920
      TabIndex        =   19
      Top             =   2400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label L1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "2"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   15
      Left            =   8760
      TabIndex        =   18
      Top             =   1920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label L1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "1"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   14
      Left            =   7920
      TabIndex        =   17
      Top             =   1920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Select the number of planets:"
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   7560
      TabIndex        =   16
      Top             =   1320
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Line Line12 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   4
      Visible         =   0   'False
      X1              =   7560
      X2              =   7560
      Y1              =   1200
      Y2              =   4320
   End
   Begin VB.Line Line11 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   4
      Visible         =   0   'False
      X1              =   9600
      X2              =   9600
      Y1              =   1200
      Y2              =   4320
   End
   Begin VB.Line Line10 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   4
      Visible         =   0   'False
      X1              =   7560
      X2              =   9600
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line Line9 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   4
      Visible         =   0   'False
      X1              =   7560
      X2              =   9600
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label L1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Brown "
      ForeColor       =   &H00000040&
      Height          =   255
      Index           =   13
      Left            =   5280
      TabIndex        =   15
      Top             =   3840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label L1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Dull Red"
      ForeColor       =   &H00000080&
      Height          =   255
      Index           =   12
      Left            =   5280
      TabIndex        =   14
      Top             =   3480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label L1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Bright Red"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   11
      Left            =   5280
      TabIndex        =   13
      Top             =   3120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label L1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Orange"
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   10
      Left            =   5280
      TabIndex        =   12
      Top             =   2760
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label L1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Yellow"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Index           =   9
      Left            =   5280
      TabIndex        =   11
      Top             =   2400
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label L1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Blue"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Index           =   8
      Left            =   5280
      TabIndex        =   10
      Top             =   2040
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label L1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "White"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   5280
      TabIndex        =   9
      Top             =   1680
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Select the color of the star:"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   4800
      TabIndex        =   8
      Top             =   1320
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Line Line8 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   4
      Visible         =   0   'False
      X1              =   4800
      X2              =   4800
      Y1              =   1200
      Y2              =   4320
   End
   Begin VB.Line Line7 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   4
      Visible         =   0   'False
      X1              =   6840
      X2              =   6840
      Y1              =   1200
      Y2              =   4320
   End
   Begin VB.Line Line6 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   4
      Visible         =   0   'False
      X1              =   4800
      X2              =   6840
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line Line5 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   4
      Visible         =   0   'False
      X1              =   4800
      X2              =   6840
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label L1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Neutron Star"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   6
      Left            =   2520
      TabIndex        =   7
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Label L1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Dwarf Star"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   5
      Left            =   2520
      TabIndex        =   6
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label L1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Small"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   4
      Left            =   2520
      TabIndex        =   5
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label L1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Medium"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   3
      Left            =   2520
      TabIndex        =   4
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label L1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Large"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   2
      Left            =   2520
      TabIndex        =   3
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label L1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Giant"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   1
      Left            =   2520
      TabIndex        =   2
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label L1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Super Giant"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   0
      Left            =   2520
      TabIndex        =   1
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Select the size of the star:"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   2040
      TabIndex        =   0
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Line Line4 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   4
      X1              =   2040
      X2              =   2040
      Y1              =   1200
      Y2              =   4320
   End
   Begin VB.Line Line3 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   4
      X1              =   4080
      X2              =   4080
      Y1              =   1200
      Y2              =   4320
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   4
      X1              =   2040
      X2              =   4080
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   4
      X1              =   2040
      X2              =   4080
      Y1              =   1200
      Y2              =   1200
   End
End
Attribute VB_Name = "Questionfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'ESSS

'This is where you build the foundation of your solar
'system, staged in three steps so there is little
'user confusion.

Option Explicit

Private Sub Form_Activate()
    Randomize
    Dim n As Byte, xx As Integer, yy As Integer
    For n = 1 To 254
        xx = Int(Rnd * Me.Width + 1)
        yy = Int(Rnd * Me.Height + 1)
        If (xx >= Line4.X1 And xx <= Line3.X1 And yy <= Line2.Y1 And yy >= Line1.Y1) Or (xx >= Line8.X1 And xx <= Line7.X1 And yy <= Line6.Y1 And yy >= Line5.Y1) Or (xx >= Line12.X1 And xx <= Line11.X1 And yy <= Line10.Y1 And yy >= Line9.Y1) Then
        Else: PSet (xx, yy), RGB(Int(Rnd * 11) + 245, Int(Rnd * 11) + 245, Int(Rnd * 11) + 245)
        End If
        xx = Int(Rnd * Me.Width + 1)
        yy = Int(Rnd * Me.Height + 1)
        If (xx >= Line4.X1 And xx <= Line3.X1 And yy <= Line2.Y1 And yy >= Line1.Y1) Or (xx >= Line8.X1 And xx <= Line7.X1 And yy <= Line6.Y1 And yy >= Line5.Y1) Or (xx >= Line12.X1 And xx <= Line11.X1 And yy <= Line10.Y1 And yy >= Line9.Y1) Then
        Else: PSet (xx, yy), RGB(Int(Rnd * 11) + 245, Int(Rnd * 11) + 245, Int(Rnd * 11) + 245)
        End If
    Next n
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("o") Then
        If L1(24).Visible = True Then
            StartS = 2
            Unload Me
        End If
    ElseIf KeyAscii = Asc("c") Then
        Unload Me
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    TopToolBarfrm.Show
End Sub
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'Button click---------------------------------------------
'/////////////////////////////////////////////////////////
Private Sub L1_Click(Index As Integer)
    If Index = 0 Then
        StarSize = 2400
        Label32.Caption = L1(Index).Caption
        Call Set1On
    ElseIf Index = 1 Then
        StarSize = 1200
        Label32.Caption = L1(Index).Caption
        Call Set1On
    ElseIf Index = 2 Then
        StarSize = 600
        Label32.Caption = L1(Index).Caption
        Call Set1On
    ElseIf Index = 3 Then
        StarSize = 300
        Label32.Caption = L1(Index).Caption
        Call Set1On
    ElseIf Index = 4 Then
        StarSize = 200
        Label32.Caption = L1(Index).Caption
        Call Set1On
    ElseIf Index = 5 Then
        StarSize = 150
        Label32.Caption = L1(Index).Caption
        Call Set1On
    ElseIf Index = 6 Then
        StarSize = 120
        Label32.Caption = L1(Index).Caption
        Call Set1On
    ElseIf Index > 6 And Index < 14 Then
        StarColor = L1(Index).BackColor
        Label33.Caption = L1(Index).Caption
        Call Set2On
    ElseIf Index = 14 Then
        PlanetNumber = L1(Index).Caption
        Label34.Caption = L1(Index).Caption & " planet"
        OC.Visible = True
        AB.Visible = True
        L1(24).Visible = True
    ElseIf Index > 14 And Index < 24 Then
        PlanetNumber = L1(Index).Caption
        Label34.Caption = L1(Index).Caption & " planets"
        OC.Visible = True
        AB.Visible = True
        L1(24).Visible = True
    ElseIf Index = 24 Then
        StartS = 2
        CanSave = 254
        Unload Me
    End If
End Sub

Private Sub Ca_Click()
    Unload Me
End Sub

Private Sub AB_Click()
    If AB.BorderStyle = 0 Then
        AB.BorderStyle = 1
    Else: AB.BorderStyle = 0
    End If
    If SSAsteroid = 0 Then
        SSAsteroid = 1
    Else: SSAsteroid = 0
    End If
End Sub

Private Sub OC_Click()
    If OC.BorderStyle = 0 Then
        OC.BorderStyle = 1
    Else: OC.BorderStyle = 0
    End If
    If SSComet = 0 Then
        SSComet = 1
    Else: SSComet = 0
    End If
End Sub
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'Highlight------------------------------------------------
'/////////////////////////////////////////////////////////
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Unhighlight2
End Sub

Private Sub L1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Unhighlight2
    If Index <= 6 Or Index >= 14 Then
        L1(Index).ForeColor = &H404040
        L1(Index).BackColor = &HFF00&
    End If
    Select Case Index
    Case 7:
        L1(Index).ForeColor = &H404040
        L1(Index).BackColor = &HFFFFFF
    Case 8:
        L1(Index).ForeColor = &H404040
        L1(Index).BackColor = &HFFFF00
    Case 9:
        L1(Index).ForeColor = &H404040
        L1(Index).BackColor = &HFFFF&
    Case 10:
        L1(Index).ForeColor = &H404040
        L1(Index).BackColor = &H80FF&
    Case 11:
        L1(Index).ForeColor = &H404040
        L1(Index).BackColor = &HFF&
    Case 12:
        L1(Index).ForeColor = &H404040
        L1(Index).BackColor = &H80&
    Case 13:
        L1(Index).ForeColor = &H404040
        L1(Index).BackColor = &H40&
    End Select
End Sub

Private Sub Ca_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Unhighlight2
    Ca.BackColor = &HFF00&
    Ca.ForeColor = &H404040
End Sub

Private Sub OC_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Unhighlight2
    OC.BackColor = &HFF00&
    OC.ForeColor = &H404040
End Sub

Private Sub AB_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Unhighlight2
    AB.BackColor = &HFF00&
    AB.ForeColor = &H404040
End Sub

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'Subroutines-------------------------------------------
'//////////////////////////////////////////////////////
Sub Set1On()
    Line5.Visible = True: Line6.Visible = True: Line7.Visible = True: Line8.Visible = True
    Dim n As Integer
    For n = 7 To 13
        L1(n).Visible = True
    Next n
    Label9.Visible = True
End Sub

Sub Set2On()
    Line9.Visible = True: Line10.Visible = True: Line11.Visible = True: Line12.Visible = True
    Dim n As Integer
    For n = 13 To 23
        L1(n).Visible = True
    Next n
    Label17.Visible = True
End Sub

Sub Unhighlight2()
    Dim n As Integer
    For n = 0 To 24
        If n <= 6 Or n >= 14 Then
            L1(n).ForeColor = &HFF00&
            L1(n).BackColor = &H404040
        End If
    Next n
    For n = 7 To 13
        Select Case n
        Case 7:
            L1(n).ForeColor = &HFFFFFF
            L1(n).BackColor = &H404040
        Case 8:
            L1(n).ForeColor = &HFFFF00
            L1(n).BackColor = &H404040
        Case 9:
            L1(n).ForeColor = &HFFFF&
            L1(n).BackColor = &H404040
        Case 10:
            L1(n).ForeColor = &H80FF&
            L1(n).BackColor = &H404040
        Case 11:
            L1(n).ForeColor = &HFF&
            L1(n).BackColor = &H404040
        Case 12:
            L1(n).ForeColor = &H80&
            L1(n).BackColor = &H404040
        Case 13:
            L1(n).ForeColor = &H40&
            L1(n).BackColor = &H404040
        End Select
    Next n
    AB.BackColor = &H404040
    AB.ForeColor = &HFF00&
    OC.BackColor = &H404040
    OC.ForeColor = &HFF00&
    Ca.BackColor = &H404040
    Ca.ForeColor = &HFF00&
End Sub

'ESSS
