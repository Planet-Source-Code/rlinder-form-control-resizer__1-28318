VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin Project1.ctlResize ctlResize1 
      Height          =   2775
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   4895
      orgWidth        =   4215
      orgHeight       =   2775
      BackColor       =   -2147483633
      Begin VB.PictureBox Picture1 
         Height          =   855
         Left            =   240
         Picture         =   "Form1.frx":0000
         ScaleHeight     =   795
         ScaleMode       =   0  'User
         ScaleWidth      =   3795
         TabIndex        =   4
         Top             =   120
         Width           =   3855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   375
         Left            =   2040
         TabIndex        =   3
         Top             =   2040
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   375
         Left            =   480
         TabIndex        =   2
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   480
         TabIndex        =   1
         Text            =   "This Text Box Will Resize With The Form."
         Top             =   1200
         Width           =   3135
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub Form_Resize()
    On Error Resume Next
       ctlResize1.Move 0, 0, ScaleWidth, ScaleHeight

End Sub
