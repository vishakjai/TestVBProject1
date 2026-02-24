VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form9 
   Caption         =   "Form9"
   ClientHeight    =   8025
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11565
   LinkTopic       =   "Form9"
   Picture         =   "splash.frx":0000
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   4320
      TabIndex        =   0
      Top             =   6840
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Timer Timer1 
      Left            =   600
      Top             =   3840
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "PLEASE WAIT.."
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   5640
      TabIndex        =   3
      Top             =   5760
      Width           =   3495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "LOADING MENU.........."
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   855
      Left            =   5640
      TabIndex        =   2
      Top             =   4560
      Width           =   5055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "BANKING MANAGEMENT"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   975
      Left            =   4320
      TabIndex        =   1
      Top             =   2640
      Width           =   6615
   End
   Begin VB.Image Image1 
      Height          =   11130
      Left            =   0
      Picture         =   "splash.frx":90C1
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15255
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Timer1.Interval = 10

End Sub

Private Sub Timer1_Timer()
If ProgressBar1.Value >= ProgressBar1.Max Then
main.Show
Timer1.Interval = 0
ProgressBar1.Value = 0
Unload Me
Else
ProgressBar1.Value = ProgressBar1.Value + 1
End If

End Sub
