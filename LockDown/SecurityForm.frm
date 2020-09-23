VERSION 5.00
Begin VB.Form SecurityForm 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   2940
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   4575
   ControlBox      =   0   'False
   Icon            =   "SecurityForm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2940
   ScaleWidth      =   4575
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame PassFrame 
      Caption         =   "Jero-Secure 1.0"
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Tag             =   "Jero-Secure 1.0"
      Top             =   0
      Width           =   4335
      Begin VB.TextBox PassBox 
         BeginProperty Font 
            Name            =   "Copperplate Gothic Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   120
         MaxLength       =   30
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   960
         Width           =   2655
      End
      Begin VB.CommandButton CheckBut 
         Caption         =   "Submit"
         BeginProperty Font 
            Name            =   "Copperplate Gothic Bold"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   960
         Width           =   1290
      End
      Begin VB.ListBox ErrorList 
         BeginProperty Font 
            Name            =   "Copperplate Gothic Bold"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1035
         ItemData        =   "SecurityForm.frx":030A
         Left            =   120
         List            =   "SecurityForm.frx":030C
         TabIndex        =   1
         Top             =   1680
         Width           =   4095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Attempt Log:"
         BeginProperty Font 
            Name            =   "Copperplate Gothic Bold"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Width           =   2100
      End
      Begin VB.Label Logo 
         BackStyle       =   0  'Transparent
         Caption         =   "Jero-Secure 1.0"
         BeginProperty Font 
            Name            =   "Copperplate Gothic Bold"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   4095
      End
   End
End
Attribute VB_Name = "SecurityForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CountNum As Integer
Dim SLevel As Integer
Dim PassWord As String

Private Sub CheckBut_Click()
CheckPassword PassBox.Text
End Sub

Public Function LoadSecurity()
On Error Resume Next
'Set Password
PassWord = "PASSWORD"
'Reset to Normal
PassBox.Text = ""
ErrorList.Clear
'Hide all Icons\SysBar
Call WHide
'Reset Security and Chance Levels
CountNum = 0
SLevel = 0
'Show Form
Me.Show
'Constrict Mouse to Form
ConstrictMouse
'Set Focus to Password Box
PassBox.SetFocus
End Function

Public Function CheckPassword(Try As String)
On Error Resume Next
If UCase(Try) <> PassWord Then
    'Normal Procedure
    PassBox.Text = ""
    PassBox.SetFocus
    CountNum = CountNum + 1
    ErrorList.AddItem "Password Attempt: " & Chr(34) & Try & Chr(34) & " at " & Time, 0
    DoEvents
    'Extra Security Levels
    Select Case SLevel
        Case 1
            Call SLevel1
        Case 2
            Call SLevel2
        Case 3
            Call SLevel3
    End Select
    'Check To Further UpGrade
    If CountNum = 3 Then
        If SLevel >= 3 Then
            PassBox.Text = ""
            PassBox.SetFocus
            CountNum = 0
            Exit Function
            End If
        SLevel = SLevel + 1
        ErrorList.AddItem "Security UpGrade: Level " & SLevel, 0
        PassBox.Text = ""
        PassBox.SetFocus
        CountNum = 0
        End If
    PassBox.SetFocus
Else
    WShow
    Me.Hide
    OptionForm.Show
End If
AddScroll ErrorList
End Function

Private Sub CheckBut_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call DisButtons
End Sub

Private Sub ErrorList_KeyDown(KeyCode As Integer, Shift As Integer)
Call DisButtons
End Sub

Private Sub ErrorList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call DisButtons
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Call DisButtons
End Sub

Private Sub Form_LostFocus()
Me.Show
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call DisButtons
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call DisButtons
End Sub

Private Sub Logo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call DisButtons
End Sub

Private Sub PassBox_KeyDown(KeyCode As Integer, Shift As Integer)
Call DisButtons
If KeyCode = vbKeyReturn Then
    CheckPassword PassBox.Text
    End If
End Sub

Public Function ConstrictMouse()
Dim Client As RECT
Dim UpperLeft As POINT
GetClientRect Me.hWnd, Client
UpperLeft.X = Client.Left
UpperLeft.Y = Client.Top
ClientToScreen Me.hWnd, UpperLeft
OffsetRect Client, UpperLeft.X, UpperLeft.Y
ClipCursor Client
End Function

Public Function SLevel1()
'Feature for Security Level 1
PassFrame.Caption = "Waiting 5 Seconds"
DoEvents
BlockInput True
Sleep 5000
BlockInput False
DoEvents
PassFrame.Caption = PassFrame.Tag
End Function

Public Function SLevel2()
'Feature for Security Level 2
PassFrame.Caption = "Waiting 10 Seconds"
DoEvents
BlockInput True
Sleep 10000
BlockInput False
DoEvents
PassFrame.Caption = PassFrame.Tag
End Function

Public Function SLevel3()
'Feature for Security Level 3
PassFrame.Caption = "Waiting 15 Seconds"
DoEvents
BlockInput True
Sleep 15000
BlockInput False
DoEvents
PassFrame.Caption = PassFrame.Tag
End Function

Public Sub AddScroll(List As ListBox)
Dim lngGreatestWidth As Long
lngGreatestWidth = 500
SendMessage List.hWnd, LB_SETHORIZONTALEXTENT, lngGreatestWidth, 0
End Sub

Private Sub PassBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call DisButtons
End Sub

Private Sub PassFrame_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Disable Ctrl-Alt-Delete incase of screensaver
Call DisButtons
End Sub
