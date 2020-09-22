VERSION 5.00
Begin VB.Form OpenSavefrm 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Open/Save"
   ClientHeight    =   6480
   ClientLeft      =   150
   ClientTop       =   150
   ClientWidth     =   10185
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   MouseIcon       =   "OpenSavefrm.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   6480
   ScaleWidth      =   10185
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      BackColor       =   &H00404040&
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   5160
      TabIndex        =   7
      Top             =   5640
      Width           =   3375
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   315
      Left            =   3360
      TabIndex        =   2
      Top             =   1680
      Width           =   2535
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H00808080&
      ForeColor       =   &H00000080&
      Height          =   3405
      Left            =   6000
      TabIndex        =   1
      Top             =   2040
      Width           =   2535
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H00808080&
      ForeColor       =   &H00000080&
      Height          =   3465
      Left            =   3360
      TabIndex        =   0
      Top             =   2040
      Width           =   2535
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Delete"
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
      Left            =   8640
      TabIndex        =   10
      ToolTipText     =   "Delete the selected file."
      Top             =   3480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "EÞs¡lon Solar System Simulator"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   2880
      TabIndex        =   9
      Top             =   1200
      Width           =   7215
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "by ßrian Adriance"
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   2880
      TabIndex        =   8
      Top             =   1365
      Width           =   7215
   End
   Begin VB.Line Line4 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   4
      X1              =   10080
      X2              =   10080
      Y1              =   1080
      Y2              =   6360
   End
   Begin VB.Line Line3 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   4
      X1              =   2880
      X2              =   2880
      Y1              =   1080
      Y2              =   6360
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   4
      X1              =   2880
      X2              =   10080
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   4
      X1              =   2880
      X2              =   10080
      Y1              =   6360
      Y2              =   6360
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Enter a name for the file:"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   3360
      TabIndex        =   6
      Top             =   5640
      Width           =   1725
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Cancel"
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
      Left            =   8640
      TabIndex        =   5
      ToolTipText     =   "Quit back to main screen."
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Open"
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
      Left            =   8640
      TabIndex        =   4
      ToolTipText     =   "Open the selected solar system."
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Save"
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
      Left            =   8640
      TabIndex        =   3
      ToolTipText     =   "Save your solar system."
      Top             =   2040
      Width           =   1215
   End
End
Attribute VB_Name = "OpenSavefrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'ESSS

'This form is for the saving or loading of files
'using FreeFile

Option Explicit
Private OpnFile As String, SavFile As String 'Path and name of file opening, path of file saving

Private Sub Form_Activate()
    Randomize
    Label3.Caption = "Cancel"
    Dim n As Byte, xx As Integer, yy As Integer
    For n = 1 To 254
        xx = Int(Rnd * Me.Width + 1)
        yy = Int(Rnd * Me.Height + 1)
        If xx > Line4.X1 Or xx < Line3.X1 Then PSet (xx, yy), RGB(Int(Rnd * 11) + 245, Int(Rnd * 11) + 245, Int(Rnd * 11) + 245)
        If yy > Line1.Y1 Or yy < Line2.Y1 Then PSet (xx, yy), RGB(Int(Rnd * 11) + 245, Int(Rnd * 11) + 245, Int(Rnd * 11) + 245)
        xx = Int(Rnd * Me.Width + 1)
        yy = Int(Rnd * Me.Height + 1)
        If xx > Line4.X1 Or xx < Line3.X1 Then PSet (xx, yy), RGB(Int(Rnd * 11) + 245, Int(Rnd * 11) + 245, Int(Rnd * 11) + 245)
        If yy > Line1.Y1 Or yy < Line2.Y1 Then PSet (xx, yy), RGB(Int(Rnd * 11) + 245, Int(Rnd * 11) + 245, Int(Rnd * 11) + 245)
    Next n
End Sub

Private Sub Form_Load()
    On Error GoTo NextDrive1
    Dir1.Path = App.Path & "\"
    SavFile = Dir1.Path
    Exit Sub
NextDrive1:
    On Error GoTo NextDrive2
    Dir1.Path = "c:\"
    SavFile = Dir1.Path
    Exit Sub
NextDrive2:
    On Error GoTo NextDrive3
    Dir1.Path = "h:\"
    SavFile = Dir1.Path
    Exit Sub
NextDrive3:
    On Error GoTo NextDrive4
    Dir1.Path = "x:\"
    SavFile = Dir1.Path
    Exit Sub
NextDrive4:
    On Error GoTo NextDrive5
    Dir1.Path = "g:\"
    SavFile = Dir1.Path
    Exit Sub
NextDrive5:
    On Error GoTo NextDrive6
    Dir1.Path = "e:\"
    SavFile = Dir1.Path
    Exit Sub
NextDrive6:
    SavFile = InputBox("ESSS cannot find a drive on your computer. Please enter a directory to save to or load from.", "Cannot Locate Drive.")
    OpnFile = SavFile
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    TopToolBarfrm.Show
End Sub
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'SaverTools-------------------------------------------
'/////////////////////////////////////////////////////
Private Sub Dir1_Change()
  File1.Path = Dir1.Path
End Sub

Private Sub Dir1_Click()
  Dir1.Path = Dir1.List(Dir1.ListIndex)
  SavFile = Dir1.Path
End Sub

Private Sub Drive1_Change()
  Dir1.Path = Drive1.Drive
  File1.Path = Dir1.Path
End Sub

Private Sub File1_Click()
    OpnFile = FormatPath(Dir1.Path) & File1.List(File1.ListIndex)
End Sub
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'Highlight---------------------------------------------
'//////////////////////////////////////////////////////
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Unhighlight1
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Unhighlight1
    Label1.BackColor = &HFF00&: Label1.ForeColor = &H404040
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Unhighlight1
    Label2.BackColor = &HFF00&: Label2.ForeColor = &H404040
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Unhighlight1
    Label3.BackColor = &HFF00&: Label3.ForeColor = &H404040
End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Unhighlight1
    Label7.BackColor = &HFF00&: Label7.ForeColor = &H404040
End Sub
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'Button Click-------------------------------------------
'///////////////////////////////////////////////////////
Private Sub Label1_Click()
    'save
    On Error GoTo Problem1
    Dim n As Byte
    For n = 1 To Len(Text1.Text)
        If Mid(Text1.Text, n, 1) = "/" Or Mid(Text1.Text, n, 1) = "\" Or Mid(Text1.Text, n, 1) = "?" Or Mid(Text1.Text, n, 1) = Chr(34) Or Mid(Text1.Text, n, 1) = ":" Or Mid(Text1.Text, n, 1) = "*" Or Mid(Text1.Text, n, 1) = "<" Or Mid(Text1.Text, n, 1) = ">" Or Mid(Text1.Text, n, 1) = "|" Then
            Beep
            Text1.SelStart = n - 1
            Text1.SelLength = 1
            MsgBox "There is an invalid character in the solar system name.", vbCritical, "Invalid Character"
            Exit Sub
        End If
    Next n
    If Text1.Text = "" Then
        MsgBox "You must enter a name for your solar system file.", vbExclamation, "File Name Required!"
    Else: Call SaveSettings
    End If
    Exit Sub
Problem1:
    MsgBox "ESSS could not save your solar system.", vbExclamation, "Error"
    Exit Sub
End Sub

Private Sub Label2_Click()
    'open
    On Error GoTo Problem2
    If OpnFile = "" Then
        MsgBox "You must select a file to load!", vbExclamation, "No File Selected!"
    Else
        If LCase(Mid(OpnFile, Len(OpnFile) - 2, 3)) = "sss" Then
            Call GetSettings
            StartS = 3
            CanSave = 254
            Unload Me
        Else: MsgBox "That is not a valid Epsilon Solar System Simulator file.", vbExclamation, "Incorrect File Type!"
        End If
    End If
    Exit Sub
Problem2:
    MsgBox "ESSS can not open that solar system.", vbExclamation, "Error"
    Exit Sub
End Sub

Private Sub Label7_Click()
    'Delete File
    On Error GoTo Problem3
    If OpnFile = "" Then
        MsgBox "You must select a file to delete!", vbExclamation, "No File Selected!"
    Else
        If LCase(Mid(OpnFile, Len(OpnFile) - 2, 3)) = "sss" Then
            Dim Sure1 As Integer
            Sure1 = MsgBox("Are you sure you want to delete '" & OpnFile & "'?", vbYesNo, "Delete File?")
            If Sure1 = vbYes Then Kill OpnFile
        Else: MsgBox "That is not a valid Epsilon Solar System Simulator file.", vbExclamation, "Incorrect File Type!"
        End If
    End If
    Exit Sub
Problem3:
    MsgBox "ESSS can not delete that file.", vbExclamation, "Error"
    Exit Sub
End Sub

Private Sub Label3_Click()
    Unload Me
End Sub
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'Subroutines----------------------------------------------
'/////////////////////////////////////////////////////////
Sub GetSettings()
    ' StarSize, StarColor, PlanetNumber As Integer
    ' FileName, PlanetName(1 To 12) As String
    Dim Free As Byte, n As Byte
    Free = FreeFile
    Open OpnFile For Input As Free
        Input #Free, StarSize
        Input #Free, StarColor
        Input #Free, PlanetNumber
        Input #Free, Age
        Input #Free, LifeS
        Input #Free, OrbitS
        Input #Free, SelectedPlanet
        Input #Free, StarName
        Input #Free, SSComet
        Input #Free, SSAsteroid
        Input #Free, CometX
        Input #Free, CometY
        Input #Free, CometRadius
        Input #Free, CometName
        Input #Free, CometDir
        For n = 1 To 10
            Input #Free, Px(n)
            Input #Free, PRad(n)
            Input #Free, PD(n)
            Input #Free, PO(n)
            Input #Free, Py(n)
            Input #Free, POA(n)
            Input #Free, PlanetColor(n)
            Input #Free, PlanetAtmosphereColor(n)
            Input #Free, PlanetName(n)
            Input #Free, PlanetType(n)
            Input #Free, MassAugment(n)
            Input #Free, MagLevel(n)
            Input #Free, GroundElement1(n)
            Input #Free, GroundElement2(n)
            Input #Free, GroundElement3(n)
            Input #Free, AirElement1(n)
            Input #Free, AirElement2(n)
            Input #Free, AirElement3(n)
            Input #Free, PlanetTheme(n)
        Next n
    Close #Free
End Sub

Sub SaveSettings()
    'asteroids, and comet still need to be included
    FileName = Text1.Text & ".sss"
    Dim Free As Byte, n As Byte
    Free = FreeFile
    Open SavFile & "\" & FileName For Output As Free
        Print #Free, StarSize
        Print #Free, StarColor
        Print #Free, PlanetNumber
        Print #Free, Age
        Print #Free, LifeS
        Print #Free, OrbitS
        Print #Free, SelectedPlanet
        Print #Free, StarName
        Print #Free, SSComet
        Print #Free, SSAsteroid
        Print #Free, CometX
        Print #Free, CometY
        Print #Free, CometRadius
        Print #Free, CometName
        Print #Free, CometDir
        For n = 1 To 10
            Print #Free, Px(n)
            Print #Free, PRad(n)
            Print #Free, PD(n)
            Print #Free, PO(n)
            Print #Free, Py(n)
            Print #Free, POA(n)
            Print #Free, PlanetColor(n)
            Print #Free, PlanetAtmosphereColor(n)
            Print #Free, PlanetName(n)
            Print #Free, PlanetType(n)
            Print #Free, MassAugment(n)
            Print #Free, MagLevel(n)
            Print #Free, GroundElement1(n)
            Print #Free, GroundElement2(n)
            Print #Free, GroundElement3(n)
            Print #Free, AirElement1(n)
            Print #Free, AirElement2(n)
            Print #Free, AirElement3(n)
            Print #Free, PlanetTheme(n)
        Next n
    Close #Free
    MsgBox "Save complete!", vbOKOnly, "Save Confirmation"
    OpenSavefrm.Label3.Caption = "OK"
End Sub

Sub Unhighlight1()
    Label1.BackColor = &H404040: Label1.ForeColor = &HFF00&
    Label2.BackColor = &H404040: Label2.ForeColor = &HFF00&
    Label3.BackColor = &H404040: Label3.ForeColor = &HFF00&
    Label7.BackColor = &H404040: Label7.ForeColor = &HFF00&
End Sub

Private Sub Text1_Change()
    If Len(Text1.Text) > 15 Then
        Text1.Text = Mid(Text1.Text, 1, 15)
        Beep
        Text1.SelStart = Len(Text1.Text)
    End If
    If Text1.SelStart > 1 Then
        If Mid(Text1.Text, Text1.SelStart, 1) = "/" Or Mid(Text1.Text, Text1.SelStart, 1) = "\" Or Mid(Text1.Text, Text1.SelStart, 1) = "?" Or Mid(Text1.Text, Text1.SelStart, 1) = Chr(34) Or Mid(Text1.Text, Text1.SelStart, 1) = ":" Or Mid(Text1.Text, Text1.SelStart, 1) = "*" Or Mid(Text1.Text, Text1.SelStart, 1) = "<" Or Mid(Text1.Text, Text1.SelStart, 1) = ">" Or Mid(Text1.Text, Text1.SelStart, 1) = "|" Then
            Beep
            Text1.Text = Mid(Text1.Text, 1, Len(Text1.Text) - 1)
            Text1.SelStart = Len(Text1.Text)
        End If
    ElseIf Text1.Text = "/" Or Text1.Text = "\" Or Text1.Text = "?" Or Text1.Text = Chr(34) Or Text1.Text = ":" Or Text1.Text = "*" Or Text1.Text = "<" Or Text1.Text = ">" Or Text1.Text = "|" Then
        Beep
        Text1.Text = ""
    End If
End Sub

'ESSS
