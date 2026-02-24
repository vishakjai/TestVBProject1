VERSION 5.00
Begin VB.Form Form8 
   Caption         =   "Form8"
   ClientHeight    =   8415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11400
   LinkTopic       =   "Form8"
   Picture         =   "reports.frx":0000
   ScaleHeight     =   8415
   ScaleWidth      =   11400
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Transaction Reports"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4080
      TabIndex        =   3
      Top             =   6240
      Width           =   3495
   End
   Begin VB.CommandButton custreport 
      BackColor       =   &H8000000C&
      Caption         =   "Customer Report"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4080
      TabIndex        =   2
      Top             =   2040
      Width           =   3615
   End
   Begin VB.CommandButton withdaw 
      Caption         =   "Withdrawal Report"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4080
      TabIndex        =   1
      Top             =   5040
      Width           =   3615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Deposit Report"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4080
      TabIndex        =   0
      Top             =   3480
      Width           =   3615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "REPORT"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4080
      TabIndex        =   4
      Top             =   360
      Width           =   3735
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
DataReport4.Show
End Sub

Private Sub Command3_Click()
DataReport2.Show

End Sub

Private Sub custreport_Click()
DataReport1.Show
End Sub

Private Sub withdaw_Click()
DataReport3.Show
End Sub
