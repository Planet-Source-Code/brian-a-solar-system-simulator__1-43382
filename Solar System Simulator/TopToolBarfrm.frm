VERSION 5.00
Begin VB.Form TopToolBarfrm 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "ESSS"
   ClientHeight    =   2940
   ClientLeft      =   -150
   ClientTop       =   -300
   ClientWidth     =   11385
   ForeColor       =   &H00C0C0C0&
   Icon            =   "TopToolBarfrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "TopToolBarfrm.frx":030A
   MousePointer    =   99  'Custom
   ScaleHeight     =   2940
   ScaleWidth      =   11385
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer3 
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer Timer2 
      Interval        =   300
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   60
      Left            =   0
      Top             =   0
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   480
      TabIndex        =   27
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label L2 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "&Run in Background"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   330
      Index           =   7
      Left            =   5280
      MousePointer    =   99  'Custom
      TabIndex        =   26
      ToolTipText     =   "Make a screen saver based on the solar system."
      Top             =   600
      Width           =   4095
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   5400
      MousePointer    =   99  'Custom
      TabIndex        =   25
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   5520
      MousePointer    =   99  'Custom
      TabIndex        =   24
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   5400
      MousePointer    =   99  'Custom
      TabIndex        =   23
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   2760
      MousePointer    =   99  'Custom
      TabIndex        =   22
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   2520
      MousePointer    =   99  'Custom
      TabIndex        =   21
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   2640
      MousePointer    =   99  'Custom
      TabIndex        =   20
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Comet Name:"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1680
      MousePointer    =   99  'Custom
      TabIndex        =   19
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Number of Planets:"
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   1680
      MousePointer    =   99  'Custom
      TabIndex        =   18
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label L2 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "&Edit  Planet"
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
      Height          =   615
      Index           =   6
      Left            =   7560
      MousePointer    =   99  'Custom
      TabIndex        =   17
      ToolTipText     =   "Edit the selected planet's properties."
      Top             =   1680
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Planet Type:"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   4320
      MousePointer    =   99  'Custom
      TabIndex        =   16
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Planet Number:"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   4320
      MousePointer    =   99  'Custom
      TabIndex        =   15
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Planet Name:"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   4320
      MousePointer    =   99  'Custom
      TabIndex        =   14
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Star Name:"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1680
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label L2 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Statistics"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   4
      Left            =   9840
      MousePointer    =   99  'Custom
      TabIndex        =   12
      ToolTipText     =   "View the statistics of the solar system."
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label L2 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Help"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   3
      Left            =   9840
      MousePointer    =   99  'Custom
      TabIndex        =   11
      ToolTipText     =   "Help"
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Line Line11 
      BorderColor     =   &H0000FF00&
      X1              =   360
      X2              =   0
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Time Flow"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Line Line10 
      BorderColor     =   &H0000FF00&
      X1              =   1560
      X2              =   1560
      Y1              =   1320
      Y2              =   2880
   End
   Begin VB.Label L2 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Fast"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   2
      Left            =   120
      MousePointer    =   99  'Custom
      TabIndex        =   9
      ToolTipText     =   "Time passes quickly."
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label L2 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Normal"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   1
      Left            =   120
      MousePointer    =   99  'Custom
      TabIndex        =   8
      ToolTipText     =   "Time passes normaly."
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label L2 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Slow"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   0
      Left            =   120
      MousePointer    =   99  'Custom
      TabIndex        =   7
      ToolTipText     =   "Time passes slowly."
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Line Line9 
      BorderColor     =   &H0000FF00&
      X1              =   9720
      X2              =   9720
      Y1              =   1320
      Y2              =   2520
   End
   Begin VB.Line Line6 
      BorderColor     =   &H0000FF00&
      X1              =   1080
      X2              =   5160
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line5 
      BorderColor     =   &H0000FF00&
      X1              =   14040
      X2              =   6480
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "System Information"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   5040
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "by ßrian Adriance"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   120
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   240
      Width           =   3015
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "EÞs¡lon Solar System Simulator"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   120
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   60
      Width           =   3015
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "&Quit ESSS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   325
      Left            =   9720
      MousePointer    =   99  'Custom
      TabIndex        =   3
      ToolTipText     =   "Quit Epsilon Solar System Simulator."
      Top             =   90
      Width           =   1935
   End
   Begin VB.Label SaveC 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "&Save Current System"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   325
      Left            =   7440
      MousePointer    =   99  'Custom
      TabIndex        =   2
      ToolTipText     =   "Save your current Solar System."
      Top             =   90
      Width           =   1935
   End
   Begin VB.Label OpenS 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "&Open Solar System"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   325
      Left            =   5280
      MousePointer    =   99  'Custom
      TabIndex        =   1
      ToolTipText     =   "Open a saved Solar System."
      Top             =   90
      Width           =   1815
   End
   Begin VB.Label CreateN 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "&Create New"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   325
      Left            =   3240
      MousePointer    =   99  'Custom
      TabIndex        =   0
      ToolTipText     =   "Create a blank Solar System."
      Top             =   90
      Width           =   1695
   End
   Begin VB.Line Line4 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   4
      X1              =   120
      X2              =   960
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line3 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   4
      X1              =   120
      X2              =   960
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   4
      X1              =   0
      X2              =   960
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   4
      X1              =   120
      X2              =   1080
      Y1              =   0
      Y2              =   0
   End
End
Attribute VB_Name = "TopToolBarfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'EÞs¡lon Solar System Simulator
'Made By ßrian Adriance

'This is the form where the user spends most of his/her
'time, the solar system is shown here, and this is the
'loading form. From here the user can access every thing
'else. Here the user can select one of his/her planet(s)
'to be able to access the edit,and can watch the solar
'system progress.


Option Explicit
Private TopBarState As Boolean  'Tells if top bar is large: true, or small: false
Private A As Integer, B As Integer 'Mouse X and Y location
Private PlanetX(1 To 10) As Integer, PlanetY(1 To 10) As Single   'location of planet
Private PlanetRadius(1 To 10) As Integer 'Radius of a planet, used for click events
Private PlanetDir(1 To 10) As Byte 'Which way the planet appears to be moving
Private PlanetPlacement As Integer 'Make planets appear in a row, not stacked
Private SX As Integer, SY As Integer 'Location of star
Private Orbital(1 To 10) As Integer ' the orbital location of your planet
Private StarX(0 To 400) As Integer, StarY(0 To 400) As Integer 'Starscape
Private RadiusAugment(1 To 10) As Integer 'radial change by type of planet
Const PI As Double = 3.141592654 'PI rounded to the 9th place

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Dim quit As Integer
        quit = MsgBox("Make sure you save your solar system. Are you sure you want to exit the Epsilon Solar System Simulator?", vbYesNo, "Quit Verification")
        If quit = vbYes Then End
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("o") Then
        With OpenSavefrm
            .Show
            .Label2.Visible = True
            .Label1.Visible = False
            .File1.Visible = True
            .Label4.Visible = False
            .Text1.Visible = False
        End With
    End If
    If KeyAscii = Asc("s") Then
        With OpenSavefrm
            .Show
            .Label2.Visible = False
            .File1.Visible = False
            .Label1.Visible = True
            .Label4.Visible = True
            .Text1.Visible = True
        End With
    End If
    If KeyAscii = Asc("c") Then Questionfrm.Show
    If L2(6).Visible = True Then
        If KeyAscii = Asc("e") Then Planetfrm.Show
    End If
    If KeyAscii = Asc("r") Then Me.WindowState = 1
    Dim quit As Integer
    If KeyAscii = Asc("q") Then
        quit = MsgBox("Make sure you save your solar system. Are you sure you want to exit the Epsilon Solar System Simulator?", vbYesNo, "Quit Verification")
        If quit = vbYes Then End
    End If
End Sub

Private Sub Form_Activate()
    Me.Cls
    If StartS = 2 Then
        Call CreateSolarSystem
        StartS = 1
    End If
    Dim n As Byte
    If StartS = 3 Then
        SX = Me.Width / 2
        SY = Me.Height / 2.3
        FillColor = StarColor
        FillStyle = vbFSSolid
        Circle (SX, SY), StarSize, StarColor
        For n = 1 To PlanetNumber
            PlanetX(n) = Px(n)
            PlanetY(n) = Py(n)
            PlanetRadius(n) = PRad(n)
            PlanetDir(n) = PD(n)
            Orbital(n) = PO(n)
            OrbitalPlane(n) = POA(n)
        Next n
        Timer1.Enabled = True
        StartS = 1
    End If
    Dim m As Integer
    If CanSave <> 254 Then
        For m = 0 To 400
            StarX(m) = Int(Rnd * Me.Width + 1)
            StarY(m) = Int(Rnd * (Me.Height - Line3.Y1)) + Line3.Y1
        Next m
    End If
    For m = 0 To 400
        PSet (StarX(m), StarY(m)), RGB(Int(Rnd * 21) + 235, Int(Rnd * 21) + 235, Int(Rnd * 21) + 235)
    Next m
    Me.Cls
End Sub

Private Sub Form_Load()
    Dim n As Integer
    CanSave = 1
    Me.Top = 0
    Me.Left = 0
    Me.Width = Screen.Width
    Me.Height = Screen.Height
    StartS = 1
    Randomize
    Call MBorder
    Dim m As Integer
    For m = 0 To 400
        StarX(m) = Int(Rnd * Me.Width + 1)
        StarY(m) = Int(Rnd * (Me.Height - Line3.Y1)) + Line3.Y1
    Next m
End Sub

Private Sub Form_Resize()
    Call MBorder
End Sub
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'Planet Click----------------------------------------------
'//////////////////////////////////////////////////////////
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim n As Byte
    For n = 1 To PlanetNumber
        If Sqr(((X - PlanetX(n)) ^ 2) + ((Y - PlanetY(n)) ^ 2)) <= PlanetRadius(n) + RadiusAugment(n) + MassAugment(n) Then
            If SelectedPlanet <> 0 Then
                If PlanetY(SelectedPlanet) > Line3.Y1 + 15 Then
                    FillStyle = vbFSTransparent
                    Circle (PlanetX(SelectedPlanet), PlanetY(SelectedPlanet)), PlanetRadius(SelectedPlanet) + 240, Me.BackColor
                    FillStyle = vbFSSolid
                End If
            End If
            Label20.Caption = n
            SelectedPlanet = n
            Label21.Caption = PlanetType(SelectedPlanet)
            Label19.Caption = PlanetName(SelectedPlanet)
        End If
    Next n
End Sub
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'Button Click----------------------------------------------
'//////////////////////////////////////////////////////////

Private Sub CreateN_Click()
    'create new solar system
    Questionfrm.Show
End Sub

Private Sub L2_Click(Index As Integer)
    'other buttons
    If Index < 3 Then 'time flow
        L2(0).BorderStyle = 0
        L2(1).BorderStyle = 0
        L2(2).BorderStyle = 0
        L2(Index).BorderStyle = 1
    End If
    If Index = 0 Then   'slow
        Timer2.Interval = 70
        Timer1.Interval = 100
    ElseIf Index = 1 Then 'normal
        Timer2.Interval = 60
        Timer1.Interval = 50
    ElseIf Index = 2 Then 'fast
        Timer2.Interval = 50
        Timer1.Interval = 25
    ElseIf Index = 3 Then 'help
        Me.Cls
        Helpfrm.Show
    ElseIf Index = 4 Then 'statistics
        Me.Cls
        Statsfrm.Show
    ElseIf Index = 6 Then 'edit planet
        Me.Cls
        Planetfrm.Show
    ElseIf Index = 7 Then 'run in background
        Me.WindowState = 1
    End If
End Sub

Private Sub OpenS_Click()
    'open a solar system
    With OpenSavefrm
        .Show
        .Label2.Visible = True
        .Label7.Visible = True
        .Label1.Visible = False
        .File1.Visible = True
        .Label4.Visible = False
        .Text1.Visible = False
    End With
End Sub

Private Sub SaveC_Click()
    'Save your solar system
    If CanSave <> 254 Then
        MsgBox "You must create a solar system before you can save it!", vbExclamation, "No System"
    ElseIf CanSave = 254 Then
        With OpenSavefrm
            .Show
            .Label2.Visible = False
            .File1.Visible = False
            .Label1.Visible = True
            .Label4.Visible = True
            .Text1.Visible = True
        End With
    End If
End Sub

Private Sub Label4_Click()
    'QUIT ESSS
    Dim quit As Integer
    quit = MsgBox("Make sure you save your solar system. Are you sure you want to exit the Epsilon Solar System Simulator?", vbYesNo, "Quit Verification")
    If quit = vbYes Then End
End Sub
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'Button Highlight------------------------------------------
'//////////////////////////////////////////////////////////
Private Sub CreateN_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Unhighlight
    CreateN.BackColor = &HFF00&: CreateN.ForeColor = &H404040
    A = X: B = Y
    Call SetSize
End Sub

Private Sub OpenS_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Unhighlight
    OpenS.BackColor = &HFF00&: OpenS.ForeColor = &H404040
    A = X: B = Y
    Call SetSize
End Sub

Private Sub SaveC_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Unhighlight
    SaveC.BackColor = &HFF00&: SaveC.ForeColor = &H404040
    A = X: B = Y
    Call SetSize
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Unhighlight
    Label4.BackColor = &HFF00&: Label4.ForeColor = &H404040
    A = X: B = Y
    Call SetSize
End Sub

Private Sub L2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Unhighlight
    L2(Index).BackColor = &HFF00&: L2(Index).ForeColor = &H404040
    A = X: B = Y
    Call SetSize
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Unhighlight
    A = X: B = Y
    Call SetSize
    If Me.Height > Screen.Height Then
        If Y >= Screen.Height - (Screen.Height / 20) Then
            Me.Top = Me.Top - 120
        End If
    End If
    If Me.Height < 0 Then
        If Y <= Screen.Height / 20 Then
            Me.Top = Me.Top + 120
        End If
    End If
    If Me.Width > Screen.Width Then
        If X >= Screen.Width - (Screen.Width / 20) Then
            Me.Left = Me.Left - 120
        End If
    End If
    If Me.Left < 0 Then
        If X <= Screen.Width / 20 Then
            Me.Left = Me.Left + 120
        End If
    End If
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Unhighlight
    A = X: B = Y
    Call SetSize
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Unhighlight
    A = X: B = Y
    Call SetSize
End Sub
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'Refresh Solar System-------------------------------------
'/////////////////////////////////////////////////////////
Private Sub Timer1_Timer()
    'main timer controls planet speeds
    Label15.Caption = StarName
    If SelectedPlanet <> 0 Then
        Label21.Caption = PlanetType(SelectedPlanet)
        Label19.Caption = PlanetName(SelectedPlanet)
    End If
    If SSComet = 1 Then Label17.Caption = CometName
    Call AnimateSS
End Sub
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'Subroutines----------------------------------------------
'/////////////////////////////////////////////////////////
Sub CreateSolarSystem()
    'Make and show solar system
    Dim n As Byte
    Label16.Caption = PlanetNumber
    If SSComet = 1 Then
        CometX = -4000
        CometY = -240
        CometRadius = 100
        CometName = "C" & (Int(Rnd * 9) + 1)
    End If
    SX = Me.Width / 2
    SY = Me.Height / 2.3
    Age = 0
    PlanetPlacement = 100
    FillColor = StarColor
    FillStyle = vbFSSolid
    Circle (SX, SY), StarSize, StarColor
    For n = PlanetNumber To 10
        PlanetX(n) = 0
        PlanetRadius(n) = 0
    Next n
    For n = 1 To PlanetNumber
        PlanetDir(n) = 1
        MassAugment(n) = 1
        MagLevel(n) = 1
        PlanetRadius(n) = 220
        PlanetPlacement = PlanetPlacement + 800
        OrbitalPlane(n) = 0
        PlanetColor(n) = vbGreen
        PlanetAtmosphereColor(n) = vbCyan
        PlanetX(n) = SX - (StarSize + PlanetPlacement)
        PlanetY(n) = SY
        Orbital(n) = SX - PlanetX(n)
        PlanetColor(n) = vbGreen
        PlanetName(n) = "Planet" & n
        PlanetAtmosphereColor(n) = vbCyan
        PlanetType(n) = "Proto-Planet"
    Next n
    StarName = "S" & (Int(Rnd * 9) + 1)
    Label15.Caption = StarName
    Timer1.Enabled = True
End Sub

Sub AnimateSS()
    'Animate the solar system
    Dim n As Byte
    FillColor = StarColor
    FillStyle = vbFSSolid
    Circle (SX, SY), StarSize, StarColor
    For n = 1 To PlanetNumber
        If PlanetType(n) = "Proto-Planet" Then
            RadiusAugment(n) = 0
        ElseIf PlanetType(n) = "Terrestrial" Then
            RadiusAugment(n) = -30
        ElseIf PlanetType(n) = "Gas Giant" Then
            RadiusAugment(n) = 75
        ElseIf PlanetType(n) = "Frozen" Then
            RadiusAugment(n) = -60
        ElseIf PlanetType(n) = "Liquid" Then
            RadiusAugment(n) = 15
        End If
        Px(n) = PlanetX(n)
        PRad(n) = PlanetRadius(n)
        PD(n) = PlanetDir(n)
        PO(n) = Orbital(n)
        POA(n) = OrbitalPlane(n)
        Py(n) = PlanetY(n)
        Call Overlap(1, n)
        If PlanetDir(n) = 1 Then
            PlanetX(n) = PlanetX(n) + 60
        ElseIf PlanetDir(n) = 2 Then PlanetX(n) = PlanetX(n) - 60
        End If
        If PlanetX(n) < SX Then
            PlanetRadius(n) = PlanetRadius(n) + 1
        ElseIf PlanetX(n) > SX Then
            PlanetRadius(n) = PlanetRadius(n) - 1
        End If
        If PlanetX(n) = SX - Orbital(n) Then PlanetRadius(n) = 220
        If PlanetX(n) = SX + Orbital(n) Then PlanetRadius(n) = 220
        Call Overlap(2, n)
        If PlanetX(n) >= SX + Orbital(n) Then
            PlanetDir(n) = 2
        ElseIf PlanetX(n) <= SX - Orbital(n) Then
            PlanetDir(n) = 1
        End If
    Next n
    If SSComet = 1 Then Call ShowComet
End Sub

Sub MBorder()
    'Make Border and show button
    Dim n As Integer
    Line1.X1 = 30
    Line1.X2 = 30
    Line1.Y1 = 0
    Line4.X1 = Me.Width - 15
    Line4.X2 = Me.Width - 15
    Line4.Y1 = 0
    Line2.X1 = 0
    Line2.X2 = Me.Width
    Line2.Y1 = 0
    Line2.Y2 = 0
    Line3.X1 = 0
    Line3.X2 = Me.Width
    If TopBarState = False Then
        Line3.Y1 = (Screen.Height / 20)
        Line3.Y2 = (Screen.Height / 20)
        Line4.Y2 = (Screen.Height / 20)
        Line1.Y2 = (Screen.Height / 20)
        Line9.Y2 = (Screen.Height / 3)
        Line10.Y2 = (Screen.Height / 3)
        Line5.Visible = False: Line6.Visible = False: Line9.Visible = False: Line10.Visible = False: Line11.Visible = False
        Label1.Visible = False: Label2.Visible = False: Label3.Visible = False: Label7.Visible = False: Label9.Visible = False: Label10.Visible = False: Label11.Visible = False: Label12.Visible = False: Label13.Visible = False: Label15.Visible = False: Label16.Visible = False: Label17.Visible = False: Label19.Visible = False: Label20.Visible = False: Label21.Visible = False
        For n = 0 To 7
            If n <> 5 Then L2(n).Visible = False
        Next n
    Else
        Line3.Y1 = (Screen.Height / 3)
        Line3.Y2 = (Screen.Height / 3)
        Line4.Y2 = Screen.Height / 3
        Line1.Y2 = Screen.Height / 3
        Line5.Visible = True: Line6.Visible = True: Line9.Visible = True: Line10.Visible = True: Line11.Visible = True
        Label1.Visible = True: Label2.Visible = True: Label3.Visible = True: Label7.Visible = True: Label9.Visible = True: Label10.Visible = True: Label11.Visible = True: Label12.Visible = True: Label13.Visible = True: Label15.Visible = True: Label16.Visible = True: Label17.Visible = True: Label19.Visible = True: Label20.Visible = True: Label21.Visible = True
        For n = 0 To 7
            If n <> 5 Then L2(n).Visible = True
            If Label20.Caption <> "" Then L2(6).Visible = True Else: L2(6).Visible = False
        Next n
        If SelectedPlanet = 0 Then L2(6).Visible = False
    End If
End Sub

Sub SetSize()
    Dim d As Integer
    'change top bar size
    If TopBarState = False Then
        If B < Screen.Height / 20 Then
            TopBarState = True
            d = 1
            Call MBorder
        End If
    ElseIf TopBarState = True Then
        If B > Screen.Height / 3 Then
            TopBarState = False
            d = 1
            Call MBorder
        End If
    End If
    'move objects accordingly
    Dim f As Byte
    Dim m As Integer
    If Line3.Y1 = Screen.Height / 3 Then
        If SY = 3913 Then
            SY = SY + (Screen.Height / 4)
            For f = 1 To PlanetNumber
                PlanetY(f) = PlanetY(f) + (Screen.Height / 4)
            Next f
            Me.Cls
            CometY = CometY + (Screen.Height / 4)
        End If
        If d = 1 Then
            For m = 0 To 400
                StarY(m) = StarY(m) + (Screen.Height / 4)
            Next m
            d = 2
        End If
    ElseIf Line3.Y1 = Screen.Height / 20 Then
        If SY = 3913 + (Screen.Height / 4) Then
            SY = SY - (Screen.Height / 4)
            For f = 1 To PlanetNumber
                PlanetY(f) = PlanetY(f) - (Screen.Height / 4)
            Next f
            Me.Cls
            CometY = CometY - (Screen.Height / 4)
        End If
        If d = 1 Then
            For m = 0 To 400
                StarY(m) = StarY(m) - (Screen.Height / 4)
            Next m
            d = 2
        End If
    End If
End Sub

Sub Unhighlight()
    'reverts buttons to idle state
    Dim n As Integer
    For n = 0 To 7
        If n <> 5 Then
            L2(n).BackColor = &H404040: L2(n).ForeColor = &HFF00&
        End If
    Next n
    CreateN.BackColor = &H404040: CreateN.ForeColor = &HFF00&
    OpenS.BackColor = &H404040: OpenS.ForeColor = &HFF00&
    SaveC.BackColor = &H404040: SaveC.ForeColor = &HFF00&
    Label4.BackColor = &H404040: Label4.ForeColor = &HFF00&
End Sub

Sub Overlap(ByRef s As Byte, ByRef n As Byte)
    'Handle planetary/solar overlap
    Dim f As Byte
    Call PointIntercept(n)
    If PlanetY(n) > Line3.Y1 Then
        Select Case s
        Case 1:
            If Sqr(((PlanetX(n) - SX) ^ 2) + ((PlanetY(n) - SY) ^ 2)) >= StarSize - PlanetRadius(n) - MassAugment(n) - RadiusAugment(n) And Sqr(((PlanetX(n) - SX) ^ 2) + ((PlanetY(n) - SY) ^ 2)) <= StarSize + PlanetRadius(n) + MassAugment(n) + RadiusAugment(n) Then
                FillColor = StarColor
                FillStyle = vbFSSolid
                Circle (SX, SY), StarSize, StarColor
            End If
            If SelectedPlanet <> 0 Then
                If PlanetY(SelectedPlanet) > Line3.Y1 + 240 Then
                    FillStyle = vbFSTransparent
                    Circle (PlanetX(SelectedPlanet), PlanetY(SelectedPlanet)), PlanetRadius(SelectedPlanet) + 240, Me.BackColor
                    FillStyle = vbFSSolid
                End If
                If PlanetY(SelectedPlanet) - 240 <= Line3.Y1 Then Line3.Refresh
            End If
            If PlanetDir(n) = 1 Or Sqr(((PlanetX(n) - SX) ^ 2) + ((PlanetY(n) - SY) ^ 2)) > StarSize - PlanetRadius(n) - MassAugment(n) - RadiusAugment(n) Then
                FillColor = Me.BackColor
                If PlanetRadius(n) + RadiusAugment(n) + MassAugment(n) > 0 Then Circle (PlanetX(n), PlanetY(n)), PlanetRadius(n) + RadiusAugment(n) + MassAugment(n), Me.BackColor
            End If
            If PlanetDir(n) = 1 And Sqr(((PlanetX(n) - SX) ^ 2) + ((PlanetY(n) - SY) ^ 2)) < StarSize - PlanetRadius(n) - MassAugment(n) - RadiusAugment(n) Then
                FillColor = StarColor
                If PlanetRadius(n) + RadiusAugment(n) + MassAugment(n) > 0 Then Circle (PlanetX(n), PlanetY(n)), PlanetRadius(n) + RadiusAugment(n) + MassAugment(n), StarColor
            End If
        Case 2:
            If Sqr(((PlanetX(n) - SX) ^ 2) + ((PlanetY(n) - SY) ^ 2)) >= StarSize - PlanetRadius(n) - MassAugment(n) - RadiusAugment(n) And Sqr(((PlanetX(n) - SX) ^ 2) + ((PlanetY(n) - SY) ^ 2)) <= StarSize + PlanetRadius(n) + MassAugment(n) + RadiusAugment(n) Then
                FillColor = StarColor
                FillStyle = vbFSSolid
                Circle (SX, SY), StarSize, StarColor
            End If
            For f = 1 To PlanetNumber
                If PlanetY(f) > Line3.Y1 + 15 Then
                    If PlanetDir(f) = 1 Then
                        FillColor = PlanetColor(f)
                        If PlanetRadius(f) + RadiusAugment(f) + MassAugment(f) > 0 Then Circle (PlanetX(f), PlanetY(f)), PlanetRadius(f) + RadiusAugment(f) + MassAugment(f), PlanetAtmosphereColor(f)
                    End If
                End If
            Next f
            If PlanetDir(n) = 1 Or Sqr(((PlanetX(n) - SX) ^ 2) + ((PlanetY(n) - SY) ^ 2)) > StarSize - PlanetRadius(n) - MassAugment(n) - RadiusAugment(n) Or (PlanetDir(n) = 2 And Sqr(((PlanetX(n) - SX) ^ 2) + ((PlanetY(n) - SY) ^ 2)) > StarSize - MassAugment(n) - RadiusAugment(n)) Then
                FillColor = PlanetColor(n)
                If PlanetRadius(n) + RadiusAugment(n) + MassAugment(n) > 0 Then Circle (PlanetX(n), PlanetY(n)), PlanetRadius(n) + RadiusAugment(n) + MassAugment(n), PlanetAtmosphereColor(n)
                For f = 1 To PlanetNumber
                    If n <> f Then
                        If Sqr(((PlanetX(n) - PlanetX(f)) ^ 2) + ((PlanetY(n) - PlanetY(f)) ^ 2)) < PlanetRadius(n) * 2 Then
                            If PlanetRadius(n) > PlanetRadius(f) Then
                                FillColor = PlanetColor(n)
                                If PlanetRadius(n) + RadiusAugment(n) + MassAugment(n) > 0 Then Circle (PlanetX(n), PlanetY(n)), PlanetRadius(n) + RadiusAugment(n) + MassAugment(n), PlanetAtmosphereColor(n)
                            ElseIf PlanetRadius(f) > PlanetRadius(n) Then
                                FillColor = PlanetColor(f)
                                If PlanetRadius(f) + RadiusAugment(f) + MassAugment(f) > 0 Then Circle (PlanetX(f), PlanetY(f)), PlanetRadius(f) + RadiusAugment(f) + MassAugment(f), PlanetAtmosphereColor(f)
                            End If
                        End If
                    End If
                Next f
            End If
            If PlanetDir(n) = 2 Then
                If Sqr(((PlanetX(n) - SX) ^ 2) + ((PlanetY(n) - SY) ^ 2)) < StarSize + PlanetRadius(n) + MassAugment(n) + RadiusAugment(n) And Sqr(((PlanetX(n) - SX) ^ 2) + ((PlanetY(n) - SY) ^ 2)) > StarSize - PlanetRadius(n) - MassAugment(n) - RadiusAugment(n) Then
                    FillColor = StarColor
                    Circle (SX, SY), StarSize, StarColor
                End If
            End If
            If SelectedPlanet <> 0 Then
                If PlanetY(SelectedPlanet) > Line3.Y1 + 240 Then
                    FillStyle = vbFSTransparent
                    Circle (PlanetX(SelectedPlanet), PlanetY(SelectedPlanet)), PlanetRadius(SelectedPlanet) + 240, RGB(210, 210, 210)
                    FillStyle = vbFSSolid
                End If
                If PlanetY(SelectedPlanet) - 240 <= Line3.Y1 Then Line3.Refresh
            End If
        End Select
    ElseIf PlanetY(n) = Line3.Y1 Then Me.Cls
    End If
    If PlanetY(n) - PlanetRadius(n) - MassAugment(n) - RadiusAugment(n) <= Line3.Y1 Then Line3.Refresh
End Sub

Private Sub Timer2_Timer()
    'solar system age
    Age = Age + 1
    'clock
    Label3.Caption = Format$(Time$, "hh:mm:ss AM/PM")
    If SelectedPlanet = 0 Then L2(6).Visible = False
End Sub

Private Sub Timer3_Timer()
    'Starscape
    Dim m As Integer
    For m = 0 To 400
        If Sqr(((SX - StarX(m)) ^ 2) + ((SY - StarY(m)) ^ 2)) > StarSize And StarY(m) > Line3.Y1 Then
            PSet (StarX(m), StarY(m)), RGB(Int(Rnd * 21) + 235, Int(Rnd * 21) + 235, Int(Rnd * 21) + 235)
        End If
    Next m
End Sub

Sub ShowComet()
    'All comet information
    If CometX >= Me.Width + 4000 Then CometDir = 2
    If CometX <= -4000 Then CometDir = 1
    If (Sqr(((SX - CometX) ^ 2) + ((SY - CometY) ^ 2)) > StarSize - (CometRadius * 2) And CometDir = 2) Or (CometDir = 1) Then
        If CometRadius >= 0 Then
            If CometDir = 1 Then
                FillColor = Me.BackColor
                If CometRadius - 30 >= 0 Then Circle (CometX - (CometRadius / 2), (Tan(Rad(12)) * ((CometX - (CometRadius / 2)) - SX)) + SY), CometRadius - 30, Me.BackColor
                If CometRadius - 45 >= 0 Then Circle (CometX - (CometRadius / 1.5), (Tan(Rad(12)) * ((CometX - (CometRadius / 1.5)) - SX)) + SY), CometRadius - 45, Me.BackColor
                If CometRadius - 60 >= 0 Then Circle (CometX - (CometRadius), (Tan(Rad(12)) * ((CometX - (CometRadius)) - SX)) + SY), CometRadius - 60, Me.BackColor
                If CometRadius - 75 >= 0 Then Circle (CometX - (CometRadius * 1.5), (Tan(Rad(12)) * ((CometX - (CometRadius * 1.5)) - SX)) + SY), CometRadius - 75, Me.BackColor
                If CometRadius - 90 >= 0 Then Circle (CometX - (CometRadius * 2), (Tan(Rad(12)) * ((CometX - (CometRadius * 2)) - SX)) + SY), CometRadius - 90, Me.BackColor
                If CometRadius - 105 >= 0 Then Circle (CometX - (CometRadius * 2.5), (Tan(Rad(12)) * ((CometX - (CometRadius * 2.5)) - SX)) + SY), CometRadius - 105, Me.BackColor
            ElseIf CometDir = 2 Then
                FillColor = Me.BackColor
                If CometRadius - 30 >= 0 Then Circle (CometX + (CometRadius / 2), (Tan(Rad(12)) * ((CometX + (CometRadius / 2)) - SX)) + SY), CometRadius - 30, Me.BackColor
                If CometRadius - 45 >= 0 Then Circle (CometX + (CometRadius / 1.5), (Tan(Rad(12)) * ((CometX + (CometRadius / 1.5)) - SX)) + SY), CometRadius - 45, Me.BackColor
                If CometRadius - 60 >= 0 Then Circle (CometX + (CometRadius), (Tan(Rad(12)) * ((CometX + (CometRadius)) - SX)) + SY), CometRadius - 60, Me.BackColor
                If CometRadius - 75 >= 0 Then Circle (CometX + (CometRadius * 1.5), (Tan(Rad(12)) * ((CometX + (CometRadius * 1.5)) - SX)) + SY), CometRadius - 75, Me.BackColor
                If CometRadius - 90 >= 0 Then Circle (CometX + (CometRadius * 2), (Tan(Rad(12)) * ((CometX + (CometRadius * 2)) - SX)) + SY), CometRadius - 90, Me.BackColor
                If CometRadius - 105 >= 0 Then Circle (CometX + (CometRadius * 2.5), (Tan(Rad(12)) * ((CometX + (CometRadius * 2.5)) - SX)) + SY), CometRadius - 105, Me.BackColor
            End If
        End If
        FillColor = Me.BackColor
        If CometRadius >= 0 Then Circle (CometX, CometY), CometRadius, Me.BackColor
    End If
    If CometDir = 1 Then CometX = CometX + 60
    If CometDir = 2 Then CometX = CometX - 60
    CometY = (Tan(Rad(12)) * (CometX - SX)) + SY
    If CometX <= SX Then CometRadius = CometRadius + 1
    If CometX >= SX Then CometRadius = CometRadius - 1
    If (Sqr(((SX - CometX) ^ 2) + ((SY - CometY) ^ 2)) > StarSize + CometRadius And CometDir = 2) Or (CometDir = 1) Then
        If CometDir = 1 Then
            FillColor = RGB(210, 210, 240)
            If CometRadius - 30 >= 0 Then Circle (CometX - (CometRadius / 2), (Tan(Rad(12)) * ((CometX - (CometRadius / 2)) - SX)) + SY), CometRadius - 30, RGB(210, 210, 240)
            If CometRadius - 45 >= 0 Then Circle (CometX - (CometRadius / 1.5), (Tan(Rad(12)) * ((CometX - (CometRadius / 1.5)) - SX)) + SY), CometRadius - 45, RGB(210, 210, 240)
            If CometRadius - 60 >= 0 Then Circle (CometX - (CometRadius), (Tan(Rad(12)) * ((CometX - (CometRadius)) - SX)) + SY), CometRadius - 60, RGB(210, 210, 240)
            If CometRadius - 75 >= 0 Then Circle (CometX - (CometRadius * 1.5), (Tan(Rad(12)) * ((CometX - (CometRadius * 1.5)) - SX)) + SY), CometRadius - 75, RGB(210, 210, 240)
            If CometRadius - 90 >= 0 Then Circle (CometX - (CometRadius * 2), (Tan(Rad(12)) * ((CometX - (CometRadius * 2)) - SX)) + SY), CometRadius - 90, RGB(210, 210, 240)
            If CometRadius - 105 >= 0 Then Circle (CometX - (CometRadius * 2.5), (Tan(Rad(12)) * ((CometX - (CometRadius * 2.5)) - SX)) + SY), CometRadius - 105, RGB(210, 210, 240)
        ElseIf CometDir = 2 Then
            FillColor = RGB(210, 210, 240)
            If CometRadius - 30 >= 0 Then Circle (CometX + (CometRadius / 2), (Tan(Rad(12)) * ((CometX + (CometRadius / 2)) - SX)) + SY), CometRadius - 30, RGB(210, 210, 240)
            If CometRadius - 45 >= 0 Then Circle (CometX + (CometRadius / 1.5), (Tan(Rad(12)) * ((CometX + (CometRadius / 1.5)) - SX)) + SY), CometRadius - 45, RGB(210, 210, 240)
            If CometRadius - 60 >= 0 Then Circle (CometX + (CometRadius), (Tan(Rad(12)) * ((CometX + (CometRadius)) - SX)) + SY), CometRadius - 60, RGB(210, 210, 240)
            If CometRadius - 75 >= 0 Then Circle (CometX + (CometRadius * 1.5), (Tan(Rad(12)) * ((CometX + (CometRadius * 1.5)) - SX)) + SY), CometRadius - 75, RGB(210, 210, 240)
            If CometRadius - 90 >= 0 Then Circle (CometX + (CometRadius * 2), (Tan(Rad(12)) * ((CometX + (CometRadius * 2)) - SX)) + SY), CometRadius - 90, RGB(210, 210, 240)
            If CometRadius - 105 >= 0 Then Circle (CometX + (CometRadius * 2.5), (Tan(Rad(12)) * ((CometX + (CometRadius * 2.5)) - SX)) + SY), CometRadius - 105, RGB(210, 210, 240)
        End If
        FillColor = RGB(170, 170, 190)
        If CometRadius >= 0 Then Circle (CometX, CometY), CometRadius, RGB(170, 170, 190)
    End If
End Sub

Sub PointIntercept(ByRef n As Byte)
    'get y value from x value and orbital angle
    PlanetY(n) = (Tan(Rad(-OrbitalPlane(n))) * (PlanetX(n) - SX)) + SY
End Sub

Function Rad(Degrees As Integer) As Double
    'degrees to radians
    Rad = (Degrees / 180) * PI
End Function

'ESSS
