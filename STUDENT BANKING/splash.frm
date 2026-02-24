VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form8 
   Caption         =   "Form8"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form8"
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   4200
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Timer Timer1 
      Left            =   960
      Top             =   1800
   End
End
Attribute VB_Name = "Form8"
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
