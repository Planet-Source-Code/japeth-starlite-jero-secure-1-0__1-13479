VERSION 5.00
Begin VB.Form OptionForm 
   Caption         =   "Jero-Secure 1.0"
   ClientHeight    =   2130
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4290
   Icon            =   "OptionForm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2130
   ScaleWidth      =   4290
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Main"
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   4095
      Begin VB.CommandButton ExitBut 
         Height          =   615
         Left            =   2280
         Picture         =   "OptionForm.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton EnableBut 
         Height          =   615
         Left            =   360
         Picture         =   "OptionForm.frx":17EC
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   360
         Width           =   1455
      End
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
      TabIndex        =   1
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "OptionForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub EnableBut_Click()
Me.Hide
Call SecurityForm.LoadSecurity
SecurityForm.Show
End Sub

Private Sub ExitBut_Click()
Call WShow
Unload SecurityForm
Unload Me
End
End Sub
