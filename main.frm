VERSION 5.00
Begin VB.Form main 
   Caption         =   "Form4"
   ClientHeight    =   9825
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form4"
   Picture         =   "main.frx":0000
   ScaleHeight     =   9825
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command7 
      Caption         =   "QUIT"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13920
      TabIndex        =   6
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      Caption         =   "REPORTS"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10920
      TabIndex        =   5
      Top             =   2040
      Width           =   3015
   End
   Begin VB.CommandButton Command5 
      Caption         =   "ACCOUNT TYPE"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8400
      TabIndex        =   4
      Top             =   2040
      Width           =   2535
   End
   Begin VB.CommandButton Command4 
      Caption         =   "TRANSACTIONS"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6240
      TabIndex        =   3
      Top             =   2040
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "WITHDRAWALS"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3960
      TabIndex        =   2
      Top             =   2040
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "DEPOSITS"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1920
      TabIndex        =   1
      Top             =   2040
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CUSTOMERS"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "BANK MANAGEMENT "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1575
      Left            =   0
      TabIndex        =   7
      Top             =   480
      Width           =   15255
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form2.Show

End Sub

Private Sub Command2_Click()
Form3.Show

End Sub

Private Sub Command3_Click()
Form4.Show
End Sub

Private Sub Command4_Click()
Form5.Show
End Sub

Private Sub Command5_Click()
Form6.Show
Me.Hide
End Sub

Private Sub Command6_Click()
Form8.Show

End Sub

Private Sub Command7_Click()
End
End Sub
