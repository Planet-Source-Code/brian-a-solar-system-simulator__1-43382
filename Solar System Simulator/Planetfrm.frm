VERSION 5.00
Begin VB.Form Planetfrm 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Planet Editer"
   ClientHeight    =   7755
   ClientLeft      =   -450
   ClientTop       =   0
   ClientWidth     =   9930
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   MouseIcon       =   "Planetfrm.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   7755
   ScaleWidth      =   9930
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text6 
      BackColor       =   &H00000000&
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   7680
      TabIndex        =   52
      Text            =   "C"
      Top             =   5520
      Width           =   2055
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   360
      Top             =   4560
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00000000&
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   7680
      TabIndex        =   50
      Text            =   "S"
      Top             =   4800
      Width           =   2055
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00000000&
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   7680
      TabIndex        =   33
      Text            =   "Planet"
      Top             =   4080
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00000000&
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   7680
      TabIndex        =   31
      Text            =   "1"
      Top             =   3360
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00000000&
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   7680
      TabIndex        =   29
      Text            =   "100"
      Top             =   2640
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   7680
      TabIndex        =   27
      Text            =   "0"
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Comet Name:"
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
      Left            =   7560
      TabIndex        =   53
      Top             =   5160
      Width           =   2235
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Star Name:"
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
      Left            =   7560
      TabIndex        =   51
      Top             =   4440
      Width           =   2235
   End
   Begin VB.Label L1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Cancel"
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
      Index           =   37
      Left            =   7680
      TabIndex        =   49
      Top             =   6600
      Width           =   2055
   End
   Begin VB.Label L1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Reset  Defaults"
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
      Height          =   375
      Index           =   36
      Left            =   7680
      TabIndex        =   48
      Top             =   6120
      Width           =   2055
   End
   Begin VB.Label L1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Index           =   35
      Left            =   7680
      TabIndex        =   47
      Top             =   7080
      Width           =   2055
   End
   Begin VB.Label L1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Argon"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   34
      Left            =   4320
      TabIndex        =   46
      Top             =   7080
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label L1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Helium"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   33
      Left            =   4320
      TabIndex        =   45
      Top             =   6720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label L1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Hydrogen"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   32
      Left            =   4320
      TabIndex        =   44
      Top             =   6360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label L1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Flourine"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   31
      Left            =   4320
      TabIndex        =   43
      Top             =   6000
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label L1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Nitrogen"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   30
      Left            =   4320
      TabIndex        =   42
      Top             =   5640
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label L1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Oxygen"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   29
      Left            =   4320
      TabIndex        =   41
      Top             =   5280
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label L1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Lithium"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   28
      Left            =   4320
      TabIndex        =   40
      Top             =   4200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label L1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Sulfur"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   27
      Left            =   4320
      TabIndex        =   39
      Top             =   3840
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label L1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Thallium"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   26
      Left            =   4320
      TabIndex        =   38
      Top             =   3480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label L1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Phosphorus"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   24
      Left            =   4320
      TabIndex        =   37
      Top             =   3120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label L1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Carbon"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   23
      Left            =   4320
      TabIndex        =   36
      Top             =   2760
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label L1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Boron"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   22
      Left            =   4320
      TabIndex        =   35
      Top             =   2400
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label L1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Silicon"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   21
      Left            =   4320
      TabIndex        =   34
      Top             =   2040
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Planet Name:"
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
      Left            =   7560
      TabIndex        =   32
      Top             =   3720
      Width           =   2235
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Magnetic Field: (1-100)"
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
      Left            =   7560
      TabIndex        =   30
      Top             =   3000
      Width           =   2235
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Mass: (100-10000)"
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
      Left            =   7560
      TabIndex        =   28
      Top             =   2280
      Width           =   2235
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Planet Orbit Angle:"
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
      Left            =   7560
      TabIndex        =   26
      Top             =   1560
      Width           =   2235
   End
   Begin VB.Label L1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Cratered"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   20
      Left            =   6000
      TabIndex        =   25
      Top             =   7200
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label L1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Mountainous"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   19
      Left            =   6000
      TabIndex        =   24
      Top             =   6840
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label L1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Terrestrial"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   18
      Left            =   6000
      TabIndex        =   23
      Top             =   6480
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label L1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Atmospheric"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   17
      Left            =   6000
      TabIndex        =   22
      Top             =   6120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label L1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Placid"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   16
      Left            =   6000
      TabIndex        =   21
      Top             =   5760
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label L1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Dynamically Electric"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   15
      Left            =   6000
      TabIndex        =   20
      Top             =   5400
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label L1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Radioactive"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   14
      Left            =   6000
      TabIndex        =   19
      Top             =   5040
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label L1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Stormy "
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   13
      Left            =   6000
      TabIndex        =   18
      Top             =   4680
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label L1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Volcanic"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   12
      Left            =   6000
      TabIndex        =   17
      Top             =   4320
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Planet Theme:"
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
      Left            =   6000
      TabIndex        =   16
      Top             =   3840
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Label L1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Liquid"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   11
      Left            =   6000
      TabIndex        =   15
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label L1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Frozen"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   10
      Left            =   6000
      TabIndex        =   14
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Label L1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Terrestrial"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   9
      Left            =   6000
      TabIndex        =   13
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label L1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Gas Giant"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   8
      Left            =   6000
      TabIndex        =   12
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Planet Type:"
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
      Left            =   6000
      TabIndex        =   11
      Top             =   1680
      Width           =   1395
   End
   Begin VB.Line Line7 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   2
      X1              =   7560
      X2              =   7560
      Y1              =   1440
      Y2              =   7680
   End
   Begin VB.Line Line6 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   2
      X1              =   2520
      X2              =   9840
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line5 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   2
      X1              =   5880
      X2              =   5880
      Y1              =   1440
      Y2              =   7680
   End
   Begin VB.Label L1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Third Abundant Element"
      ForeColor       =   &H0000FF00&
      Height          =   615
      Index           =   6
      Left            =   2880
      TabIndex        =   10
      Top             =   6720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label L1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Second Abundant Element"
      ForeColor       =   &H0000FF00&
      Height          =   615
      Index           =   5
      Left            =   2880
      TabIndex        =   9
      Top             =   6000
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label L1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Most Abundant Element"
      ForeColor       =   &H0000FF00&
      Height          =   615
      Index           =   4
      Left            =   2880
      TabIndex        =   8
      Top             =   5280
      Width           =   1335
   End
   Begin VB.Label L1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Third Abundant  Element"
      ForeColor       =   &H0000FF00&
      Height          =   615
      Index           =   2
      Left            =   2880
      TabIndex        =   7
      Top             =   3480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label L1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Second Abundant  Element"
      ForeColor       =   &H0000FF00&
      Height          =   615
      Index           =   1
      Left            =   2880
      TabIndex        =   6
      Top             =   2760
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label L1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Most Abundant Element"
      ForeColor       =   &H0000FF00&
      Height          =   615
      Index           =   0
      Left            =   2880
      TabIndex        =   5
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Planet Atmosphere:"
      ForeColor       =   &H0000C000&
      Height          =   195
      Left            =   2760
      TabIndex        =   4
      Top             =   4800
      Width           =   1395
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Planet Ground:"
      ForeColor       =   &H0000C000&
      Height          =   195
      Left            =   2760
      TabIndex        =   3
      Top             =   1800
      Width           =   1155
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Main Elements:"
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
      Left            =   2640
      TabIndex        =   2
      Top             =   1440
      Width           =   1395
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Planet Edit"
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
      Height          =   300
      Left            =   5505
      TabIndex        =   1
      Top             =   1080
      Width           =   1365
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
      Left            =   2520
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   600
      Width           =   7275
   End
   Begin VB.Line Line4 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   4
      X1              =   2520
      X2              =   9840
      Y1              =   7680
      Y2              =   7680
   End
   Begin VB.Line Line3 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   4
      X1              =   2520
      X2              =   9840
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   4
      X1              =   9840
      X2              =   9840
      Y1              =   960
      Y2              =   7680
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   4
      X1              =   2520
      X2              =   2520
      Y1              =   960
      Y2              =   7680
   End
End
Attribute VB_Name = "Planetfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'ESSS

'This is where the use can edit many of the selected
'planets properties. The mass and planet type will
'alter the planets radius. The theme, elements, and
'magnetic field effect life. Here the user can also
'change planetary orbital angles and the name of the
'planet, star, and comet if there is one.

Option Explicit
Private Gdown As Byte, Adown As Byte
Private ElementNumber As Integer

Private Sub Form_Activate()
    Randomize
    Dim n As Byte, xx As Integer, yy As Integer
    For n = 1 To 254
        xx = Int(Rnd * Me.Width + 1)
        yy = Int(Rnd * Me.Height + 1)
        If xx > Line2.X1 Or xx < Line3.X1 Then PSet (xx, yy), RGB(Int(Rnd * 11) + 245, Int(Rnd * 11) + 245, Int(Rnd * 11) + 245)
        If yy > Line4.Y1 Or yy < Line1.Y1 Then PSet (xx, yy), RGB(Int(Rnd * 11) + 245, Int(Rnd * 11) + 245, Int(Rnd * 11) + 245)
        xx = Int(Rnd * Me.Width + 1)
        yy = Int(Rnd * Me.Height + 1)
        If xx > Line2.X1 Or xx < Line3.X1 Then PSet (xx, yy), RGB(Int(Rnd * 11) + 245, Int(Rnd * 11) + 245, Int(Rnd * 11) + 245)
        If yy > Line4.Y1 Or yy < Line1.Y1 Then PSet (xx, yy), RGB(Int(Rnd * 11) + 245, Int(Rnd * 11) + 245, Int(Rnd * 11) + 245)
    Next n
End Sub

Private Sub Form_Load()
    If PlanetName(SelectedPlanet) = "" Then Text4.Text = "Planet " & (SelectedPlanet) Else: Text4.Text = PlanetName(SelectedPlanet)
    If StarName = "" Then Text5.Text = "S" & (Int(Rnd * 9) + 1) Else: Text5.Text = StarName
    Text1.Text = OrbitalPlane(SelectedPlanet)
    Text2.Text = MassAugment(SelectedPlanet) * 100
    Text3.Text = MagLevel(SelectedPlanet)
    If SSComet = 1 Then
        Text6.Visible = True
        Label13.Visible = True
        Text6.Text = CometName
    Else
        Text6.Visible = False
        Label13.Visible = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    TopToolBarfrm.Show
    TopToolBarfrm.Cls
End Sub
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'Button Click----------------------------------------------
'//////////////////////////////////////////////////////////
Private Sub L1_Click(Index As Integer)
    If Index = 35 Then 'Ok
        If Text1.Text = "-" Or Text1.Text = "" Then Text1.Text = "0"
        PlanetName(SelectedPlanet) = Text4.Text
        OrbitalPlane(SelectedPlanet) = Text1.Text
        StarName = Text5.Text
        MassAugment(SelectedPlanet) = Int(Text2.Text) / 100
        MagLevel(SelectedPlanet) = Text3.Text
        CometName = Text6.Text
        Unload Me
    End If
    If Index = 36 Then 'Reset defaults
        Text4.Text = "Planet " & SelectedPlanet
        Text1.Text = "0"
        Text5.Text = "S" & Int(Rnd * 9) + 1
        Text2.Text = "100"
        Text3.Text = "1"
        Text6.Text = "C" & (Int(Rnd * 9) + 1)
    End If
    If Index >= 12 And Index <= 20 Then
        PlanetTheme(SelectedPlanet) = L1(Index).Caption
    End If
    If Index = 8 Then 'Gas Giant
        PlanetType(SelectedPlanet) = "Gas Giant"
        PlanetColor(SelectedPlanet) = RGB(220, 208, 45)
        PlanetAtmosphereColor(SelectedPlanet) = PlanetColor(SelectedPlanet)
    ElseIf Index = 9 Then 'Terrestrial
        PlanetType(SelectedPlanet) = "Terrestrial"
        PlanetColor(SelectedPlanet) = RGB(134, 123, 50)
        PlanetAtmosphereColor(SelectedPlanet) = RGB(55, 245, 245)
    ElseIf Index = 10 Then 'Frozen
        PlanetType(SelectedPlanet) = "Frozen"
        PlanetColor(SelectedPlanet) = RGB(205, 235, 255)
        PlanetAtmosphereColor(SelectedPlanet) = RGB(240, 240, 240)
    ElseIf Index = 11 Then 'Liquid
        PlanetType(SelectedPlanet) = "Liquid"
        PlanetColor(SelectedPlanet) = RGB(46, 230, 210)
        PlanetAtmosphereColor(SelectedPlanet) = RGB(0, 210, 250)
    End If
    If Index = 37 Then Unload Me 'cancel
End Sub

Private Sub L1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ElementNumber = Index
    If Index >= 0 And Index <= 3 Then Gdown = 1
    If Index >= 4 And Index <= 7 Then Adown = 1
End Sub

Private Sub L1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Gdown = 2
    Adown = 2
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Gdown = 3
    Adown = 3
End Sub
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'Highlight-------------------------------------------------
'//////////////////////////////////////////////////////////
Private Sub L1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UnhighlightP
    L1(Index).ForeColor = &H404040: L1(Index).BackColor = &HFF00&
    If Gdown = 2 Or Adown = 2 Then
        If Index = 21 Or Index = 22 Or Index = 23 Or Index = 24 Or Index = 26 Or Index = 27 Or Index = 28 Or Index = 29 Or Index = 30 Or Index = 31 Or Index = 32 Or Index = 33 Or Index = 34 Then
            If ElementNumber = 0 Then
                L1(ElementNumber).Caption = "Most Abundant Element:" & vbCrLf & L1(Index).Caption
                GroundElement1(SelectedPlanet) = L1(Index).Caption
                If GroundElement1(SelectedPlanet) = GroundElement2(SelectedPlanet) Then
                    GroundElement2(SelectedPlanet) = ""
                    L1(1).Caption = "Second Abundant Element:"
                ElseIf GroundElement1(SelectedPlanet) = GroundElement3(SelectedPlanet) Then
                    GroundElement3(SelectedPlanet) = ""
                    L1(2).Caption = "Third Abundant Element:"
                End If
            ElseIf ElementNumber = 1 Then
                L1(ElementNumber).Caption = "Second Abundant Element:" & vbCrLf & L1(Index).Caption
                GroundElement2(SelectedPlanet) = L1(Index).Caption
                If GroundElement2(SelectedPlanet) = GroundElement1(SelectedPlanet) Then
                    GroundElement1(SelectedPlanet) = ""
                    L1(0).Caption = "Most Abundant Element:"
                ElseIf GroundElement2(SelectedPlanet) = GroundElement3(SelectedPlanet) Then
                    GroundElement3(SelectedPlanet) = ""
                    L1(2).Caption = "Third Abundant Element:"
                End If
            ElseIf ElementNumber = 2 Then
                L1(ElementNumber).Caption = "Third Abundant Element:" & vbCrLf & L1(Index).Caption
                GroundElement3(SelectedPlanet) = L1(Index).Caption
                If GroundElement3(SelectedPlanet) = GroundElement1(SelectedPlanet) Then
                    GroundElement1(SelectedPlanet) = ""
                    L1(0).Caption = "Most Abundant Element:"
                ElseIf GroundElement3(SelectedPlanet) = GroundElement2(SelectedPlanet) Then
                    GroundElement2(SelectedPlanet) = ""
                    L1(1).Caption = "Second Abundant Element:"
                End If
            ElseIf ElementNumber = 4 Then
                L1(ElementNumber).Caption = "Most Abundant Element:" & vbCrLf & L1(Index).Caption
                AirElement1(SelectedPlanet) = L1(Index).Caption
                If AirElement1(SelectedPlanet) = AirElement2(SelectedPlanet) Then
                    AirElement2(SelectedPlanet) = ""
                    L1(5).Caption = "Second Abundant Element:"
                ElseIf AirElement1(SelectedPlanet) = AirElement3(SelectedPlanet) Then
                    AirElement3(SelectedPlanet) = ""
                    L1(6).Caption = "Third Abundant Element:"
                End If
            ElseIf ElementNumber = 5 Then
                L1(ElementNumber).Caption = "Second Abundant Element:" & vbCrLf & L1(Index).Caption
                AirElement2(SelectedPlanet) = L1(Index).Caption
                If AirElement2(SelectedPlanet) = AirElement1(SelectedPlanet) Then
                    AirElement1(SelectedPlanet) = ""
                    L1(4).Caption = "Most Abundant Element:"
                ElseIf AirElement2(SelectedPlanet) = AirElement3(SelectedPlanet) Then
                    AirElement3(SelectedPlanet) = ""
                    L1(6).Caption = "Third Abundant Element:"
                End If
            ElseIf ElementNumber = 6 Then
                L1(ElementNumber).Caption = "Third Abundant Element:" & vbCrLf & L1(Index).Caption
                AirElement3(SelectedPlanet) = L1(Index).Caption
                If AirElement3(SelectedPlanet) = AirElement1(SelectedPlanet) Then
                    AirElement1(SelectedPlanet) = ""
                    L1(4).Caption = "Most Abundant Element:"
                ElseIf AirElement3(SelectedPlanet) = AirElement2(SelectedPlanet) Then
                    AirElement2(SelectedPlanet) = ""
                    L1(5).Caption = "Second Abundant Element:"
                End If
            End If
            Gdown = 3
            Adown = 3
        End If
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UnhighlightP
End Sub

Sub UnhighlightP()
    Dim n As Integer
    For n = 0 To 37
        If n <> 3 And n <> 7 And n <> 25 Then
            L1(n).BackColor = &H404040: L1(n).ForeColor = &HFF00&
        End If
    Next n
End Sub
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'Text Limits-----------------------------------------------
'//////////////////////////////////////////////////////////
Private Sub Text1_Change()
    On Error Resume Next
    If Text1.Text = "-" Or Text1.Text = "" Then
    Else:
        Dim n As Byte
        For n = 1 To Len(Text1.Text)
            If Mid(Text1.Text, n, 1) <> "1" And Mid(Text1.Text, n, 1) <> "2" And Mid(Text1.Text, n, 1) <> "3" And Mid(Text1.Text, n, 1) <> "4" And Mid(Text1.Text, n, 1) <> "5" And Mid(Text1.Text, n, 1) <> "6" And Mid(Text1.Text, n, 1) <> "7" And Mid(Text1.Text, n, 1) <> "8" And Mid(Text1.Text, n, 1) <> "0" And Mid(Text1.Text, n, 1) <> "9" And Mid(Text1.Text, 1, 1) <> "-" Then Text1.Text = "0"
        Next n
        If Int(Text1.Text) > 45 Then
            Text1.Text = "45"
            Text1.SelStart = Len(Text1.Text)
        End If
        If Int(Text1.Text) < -45 Then
            Text1.Text = "-45"
            Text1.SelStart = Len(Text1.Text)
        End If
    End If
End Sub

Private Sub Text2_Change()
    On Error Resume Next
    Dim n As Byte
    For n = 1 To Len(Text2.Text)
        If Mid(Text2.Text, n, 1) <> "1" And Mid(Text2.Text, n, 1) <> "2" And Mid(Text2.Text, n, 1) <> "3" And Mid(Text2.Text, n, 1) <> "4" And Mid(Text2.Text, n, 1) <> "5" And Mid(Text2.Text, n, 1) <> "6" And Mid(Text2.Text, n, 1) <> "7" And Mid(Text2.Text, n, 1) <> "8" And Mid(Text2.Text, n, 1) <> "0" And Mid(Text2.Text, n, 1) <> "9" Then Text2.Text = "100"
    Next n
    If Int(Text2.Text) > 10000 Then
        Text2.Text = "10000"
        Text2.SelStart = Len(Text2.Text)
    End If
End Sub

Private Sub Text2_LostFocus()
    If Int(Text2.Text) < 100 Then Text2.Text = "100"
End Sub

Private Sub Text3_Change()
    On Error Resume Next
    Dim n As Byte
    For n = 1 To Len(Text3.Text)
        If Mid(Text3.Text, n, 1) <> "1" And Mid(Text3.Text, n, 1) <> "2" And Mid(Text3.Text, n, 1) <> "3" And Mid(Text3.Text, n, 1) <> "4" And Mid(Text3.Text, n, 1) <> "5" And Mid(Text3.Text, n, 1) <> "6" And Mid(Text3.Text, n, 1) <> "7" And Mid(Text3.Text, n, 1) <> "8" And Mid(Text3.Text, n, 1) <> "0" And Mid(Text3.Text, n, 1) <> "9" Then Text3.Text = "1"
    Next n
    If Int(Text3.Text) > 100 Then
        Text3.Text = "100"
        Text3.SelStart = Len(Text3.Text)
    End If
    If Int(Text3.Text) < 1 Then Text3.Text = "1"
End Sub

Private Sub Text4_Change()
    If Len(Text4.Text) > 18 Then
        Text4.Text = Mid(Text4.Text, 1, 18)
        Beep
        Text4.SelStart = Len(Text4.Text)
    End If
End Sub

Private Sub Text5_Change()
    If Len(Text5.Text) > 18 Then
        Text5.Text = Mid(Text5.Text, 1, 18)
        Beep
        Text5.SelStart = Len(Text5.Text)
    End If
End Sub

Private Sub Text6_Change()
    If Len(Text6.Text) > 18 Then
        Text6.Text = Mid(Text6.Text, 1, 18)
        Beep
        Text6.SelStart = Len(Text6.Text)
    End If
End Sub
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'Pop-up Buttons--------------------------------------------
'//////////////////////////////////////////////////////////
Private Sub Timer1_Timer()
    Dim n As Integer
    If Adown = 1 Then
        For n = 29 To 34
            L1(n).Visible = True
        Next n
    End If
    If Gdown = 1 Then
        For n = 21 To 28
            If n <> 25 Then L1(n).Visible = True
        Next n
    End If
    If Gdown <> 1 Then
        For n = 21 To 28
            If n <> 25 Then L1(n).Visible = False
        Next n
    End If
    If Adown <> 1 Then
        For n = 29 To 34
            L1(n).Visible = False
        Next n
    End If
    For n = 12 To 20
        L1(n).Visible = False
    Next n
    If PlanetType(SelectedPlanet) = "Proto-Planet" Then
        Label7.Visible = False
        For n = 12 To 20
            L1(n).Visible = False
        Next n
    ElseIf PlanetType(SelectedPlanet) = "Gas Giant" Then
        Label7.Visible = True
        For n = 13 To 16
            L1(n).Visible = True
        Next n
    ElseIf PlanetType(SelectedPlanet) = "Terrestrial" Then
        Label7.Visible = True
        L1(12).Visible = True
        L1(14).Visible = True
        For n = 16 To 20
            L1(n).Visible = True
        Next n
    ElseIf PlanetType(SelectedPlanet) = "Frozen" Then
        Label7.Visible = True
        For n = 16 To 20
            L1(n).Visible = True
        Next n
    ElseIf PlanetType(SelectedPlanet) = "Liquid" Then
        Label7.Visible = True
        L1(13).Visible = True
        L1(15).Visible = True
        L1(16).Visible = True
    End If
    If GroundElement1(SelectedPlanet) <> "" Then
        L1(1).Visible = True
    Else
        L1(1).Visible = False
        L1(2).Visible = False
    End If
    If GroundElement2(SelectedPlanet) <> "" Then
        L1(2).Visible = True
    Else
        L1(2).Visible = False
    End If
    If AirElement1(SelectedPlanet) <> "" Then
        L1(5).Visible = True
    Else
        L1(5).Visible = False
        L1(6).Visible = False
    End If
    If AirElement2(SelectedPlanet) <> "" Then
        L1(6).Visible = True
    Else
        L1(6).Visible = False
    End If
End Sub

'ESSS
