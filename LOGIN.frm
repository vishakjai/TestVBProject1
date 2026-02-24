VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000008&
   Caption         =   "Form1"
   ClientHeight    =   3720
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9090
   LinkTopic       =   "Form1"
   PaletteMode     =   2  'Custom
   Picture         =   "LOGIN.frx":0000
   ScaleHeight     =   3720
   ScaleWidth      =   9090
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   7680
      TabIndex        =   6
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   5760
      TabIndex        =   5
      Top             =   3120
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   5760
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   2520
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   5760
      TabIndex        =   2
      Top             =   1920
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Change Password"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   1560
      TabIndex        =   7
      Top             =   3240
      Width           =   3975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "PASSWORD"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   2640
      TabIndex        =   3
      Top             =   2280
      Width           =   2895
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "USER NAME"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2640
      TabIndex        =   1
      Top             =   1680
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   2400
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






Private Sub Command1_Click()
frm

 If RS.State = 1 Then RS.Close
 RS.Open "SELECT * FROM logi WHERE user1='" & Text1.Text & "' AND pass='" & Text2.Text & "'", CON
 If RS.EOF = False Then
   Form9.Show
    Me.Hide
  Else
    MsgBox "plz enter correct username and passwod"
    Text1.Text = ""
    Text2.Text = ""
    End If
    
End Sub

Private Sub Label3_Click(Index As Integer)
Form7.Show

End Sub
