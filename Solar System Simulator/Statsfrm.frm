VERSION 5.00
Begin VB.Form Statsfrm 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Statistical Survey"
   ClientHeight    =   4830
   ClientLeft      =   -150
   ClientTop       =   -150
   ClientWidth     =   7680
   FillColor       =   &H00C0C0C0&
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   MouseIcon       =   "Statsfrm.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   4830
   ScaleWidth      =   7680
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Label Label9 
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
      Left            =   4440
      TabIndex        =   14
      Top             =   4200
      Width           =   2655
   End
   Begin VB.Label Year 
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
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   6000
      TabIndex        =   13
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Age of Solar System:"
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
      Height          =   240
      Left            =   4080
      TabIndex        =   12
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label NT 
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
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   5280
      TabIndex        =   11
      Top             =   3840
      Width           =   1815
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nova Threat:"
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
      Height          =   240
      Left            =   4080
      TabIndex        =   10
      Top             =   3840
      Width           =   1170
   End
   Begin VB.Label AT 
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
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   5520
      TabIndex        =   9
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Asteroid Threat:"
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
      Height          =   240
      Left            =   4080
      TabIndex        =   8
      Top             =   3480
      Width           =   1425
   End
   Begin VB.Label slp 
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
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   5160
      TabIndex        =   7
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Star Phase:"
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
      Height          =   240
      Left            =   4080
      TabIndex        =   6
      Top             =   3120
      Width           =   1035
   End
   Begin VB.Label Orbstab 
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
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   5520
      TabIndex        =   5
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Orbital Stability:"
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
      Height          =   240
      Left            =   4080
      TabIndex        =   4
      Top             =   2760
      Width           =   1380
   End
   Begin VB.Label life 
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
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   4560
      TabIndex        =   3
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Line Line6 
      BorderColor     =   &H0000FF00&
      X1              =   6840
      X2              =   7440
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line5 
      BorderColor     =   &H0000FF00&
      X1              =   4680
      X2              =   4080
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Life:"
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
      Height          =   240
      Left            =   4080
      TabIndex        =   2
      Top             =   2400
      Width           =   360
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
      Left            =   3960
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   960
      Width           =   3675
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Solar System Statistics"
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
      Height          =   240
      Left            =   3960
      TabIndex        =   0
      Top             =   1440
      Width           =   3615
   End
   Begin VB.Line Line4 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   4
      X1              =   7560
      X2              =   7560
      Y1              =   1320
      Y2              =   4680
   End
   Begin VB.Line Line3 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   4
      X1              =   7560
      X2              =   3960
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   4
      X1              =   3960
      X2              =   3960
      Y1              =   1320
      Y2              =   4680
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   4
      X1              =   3960
      X2              =   7560
      Y1              =   1320
      Y2              =   1320
   End
End
Attribute VB_Name = "Statsfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'ESSS

'This form gives the user a way to 'keep score',
'by rating various aspects of their solar system

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
    If KeyAscii = Asc("o") Then Unload Me
End Sub

Private Sub Form_Load()
    'Condition of life in the solar system
    Dim n As Byte
    For n = 1 To PlanetNumber
        Call SearchForLife(n)
    Next n
    If LifeS = 0 Then
        life.Caption = "None"
    ElseIf LifeS = 1 Then
        life.Caption = "Struggling"
    ElseIf LifeS = 2 Then
        life.Caption = "Thriving"
    ElseIf LifeS = 3 Then
        life.Caption = "Expanding"
    End If
    'As time progresses the solar system will age
    Year.Caption = Age
    'The more planets you have, and the older the system
    'is, the the orbital stability will be lower
    For n = 1 To PlanetNumber
        If OrbitalPlane(n) = 0 Then OrbitS = 2
    Next n
    For n = 1 To PlanetNumber
        If n > 1 Then
            If OrbitalPlane(n) = OrbitalPlane(n - 1) Then
                OrbitS = 2
            Else:
                OrbitS = 1
                Exit For
            End If
        Else:
            If OrbitalPlane(n) = OrbitalPlane(PlanetNumber) Then
                OrbitS = 2
            Else:
                OrbitS = 1
                Exit For
            End If
        End If
    Next n
    If OrbitS = 1 Then
        For n = 1 To PlanetNumber
            If n > 1 Then
                If Abs(OrbitalPlane(n) - OrbitalPlane(n - 1)) > 30 Then
                    OrbitS = 0
                    Exit For
                Else:
                    OrbitS = 1
                End If
            Else:
                If Abs(OrbitalPlane(n) - OrbitalPlane(PlanetNumber)) > 30 Then
                    OrbitS = 0
                    Exit For
                Else:
                    OrbitS = 1
                End If
            End If
        Next n
    End If
    If Age > 3500 Then OrbitS = 0
    If OrbitS = 0 Then
        Orbstab.Caption = "Dangerous"
    ElseIf OrbitS = 1 Then
        Orbstab.Caption = "Unsafe"
    ElseIf OrbitS = 2 Then
        Orbstab.Caption = "Safe"
    End If
    If StarSize = 2400 Then
        slp.Caption = "Super Giant"
    ElseIf StarSize = 1200 Then
        slp.Caption = "Giant"
    ElseIf StarSize = 600 Then
        slp.Caption = "Large"
    ElseIf StarSize = 300 Then
        slp.Caption = "Medium"
    ElseIf StarSize = 200 Then
        slp.Caption = "Small"
    ElseIf StarSize = 150 Then
        slp.Caption = "Dwarf Star"
    ElseIf StarSize = 120 Then
        slp.Caption = "Neutron Star"
    End If
    'If you have an asteroid belt, you risk asteroid
    'collisions with the planets
    If SSAsteroid = 1 Then
        AT.Caption = "High"
    Else: AT.Caption = "None"
    End If
    'The bigger, older, the star the more danger of a nova
    'occuring
    If StarSize = 2400 Then
        NT.Caption = "High"
    ElseIf StarSize = 1200 Then
        NT.Caption = "Some"
    Else: NT.Caption = "None"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    TopToolBarfrm.Show
End Sub
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'Button click-------------------------------------------
'///////////////////////////////////////////////////////
Private Sub Label9_Click()
    Unload Me
End Sub
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'Highlight----------------------------------------------
'///////////////////////////////////////////////////////
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label9.BackColor = &H404040: Label9.ForeColor = &HFF00&
End Sub

Private Sub Label9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label9.BackColor = &HFF00&: Label9.ForeColor = &H404040
End Sub
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'Sub-Procedures--------------------------------------------
'//////////////////////////////////////////////////////////
Sub SearchForLife(ByVal n As Byte)
    If Abs(OrbitalPlane(n)) <= 35 Then
        If PlanetTheme(n) = "Stormy" Or PlanetTheme(n) = "Dynamically Electric" Or PlanetTheme(n) = "Cratered" Then
            If GroundElement1(n) = "Carbon" Or GroundElement2(n) = "Carbon" Then
                If (GroundElement1(n) = "Carbon" Or GroundElement2(n) = "Carbon") And (GroundElement1(n) = "Silicon" Or GroundElement2(n) = "Silicon") Then
                    If GroundElement3(n) = "Lithium" Or GroundElement3(n) = "Boron" Then
                        If AirElement1(n) = "Nitrogen" Or AirElement2(n) = "Nitrogen" Then
                            If (AirElement1(n) = "Nitrogen" Or AirElement2(n) = "Nitrogen") And (GroundElement1(n) = "Oxygen" Or GroundElement2(n) = "Oxygen") Then
                                If AirElement3(n) = "Helium" Or AirElement3(n) = "Helium" Then
                                    LifeS = 1
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
        If PlanetTheme(n) = "Mountainous" Or PlanetTheme(n) = "Volcanic" Or PlanetTheme(n) = "Atmosphereic" Or PlanetTheme(n) = "Terrestrial" Then
            If GroundElement1(n) = "Carbon" Or GroundElement2(n) = "Carbon" Then
                If (GroundElement1(n) = "Carbon" Or GroundElement2(n) = "Carbon") And (GroundElement1(n) = "Sulfur" Or GroundElement2(n) = "Sulfur") Then
                    If GroundElement3(n) = "Thallium" Or GroundElement3(n) = "Phosphorus" Then
                        If AirElement1(n) = "Nitrogen" Or AirElement2(n) = "Nitrogen" Then
                            If (AirElement1(n) = "Nitrogen" Or AirElement2(n) = "Nitrogen") And (GroundElement1(n) = "Flourine" Or GroundElement2(n) = "Flourine") Then
                                If AirElement3(n) = "Argon" Or AirElement3(n) = "Argon" Then
                                    LifeS = 1
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
        If (PlanetTheme(n) = "Placid" Or PlanetTheme(n) = "Radioactive") And MagLevel(n) >= 15 Then
            If GroundElement1(n) = "Carbon" Or GroundElement2(n) = "Carbon" Then
                If (GroundElement1(n) = "Carbon" Or GroundElement2(n) = "Carbon") And (GroundElement1(n) = "Oxygen" Or GroundElement2(n) = "Oxygen") Then
                    If GroundElement3(n) = "Sulfur" Or GroundElement3(n) = "Phosphorus" Then
                        If AirElement1(n) = "Nitrogen" Or AirElement2(n) = "Nitrogen" Then
                            If (AirElement1(n) = "Nitrogen" Or AirElement2(n) = "Nitrogen") And (GroundElement1(n) = "Oxygen" Or GroundElement2(n) = "Oxygen") Then
                                If AirElement3(n) = "Hydrogen" Or AirElement3(n) = "Hydrogen" Then
                                    LifeS = 2
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
                                                                                                                                                                                        If PlanetName(n) = "Epsilon" And MagLevel(n) >= 98 And PlanetTheme(n) = "Radioactive" And GroundElement1(n) = "Lithium" And AirElement1(n) = "Flourine" And Abs(OrbitalPlane(n)) <= 23 Then LifeS = 3
End Sub


'ESSS
