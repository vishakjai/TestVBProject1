VERSION 5.00
Begin VB.Form Form7 
   Caption         =   "Form7"
   ClientHeight    =   7680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8730
   LinkTopic       =   "Form7"
   ScaleHeight     =   7680
   ScaleWidth      =   8730
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "CANCEL"
      Height          =   615
      Left            =   5520
      TabIndex        =   9
      Top             =   6240
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CHANGE"
      Height          =   615
      Left            =   2520
      TabIndex        =   8
      Top             =   6240
      Width           =   1935
   End
   Begin VB.TextBox txtpass1 
      Height          =   615
      Left            =   5280
      TabIndex        =   7
      Top             =   4560
      Width           =   1695
   End
   Begin VB.TextBox txtname1 
      Height          =   495
      Left            =   5280
      TabIndex        =   6
      Top             =   3480
      Width           =   1575
   End
   Begin VB.TextBox txtpass 
      Height          =   495
      Left            =   5280
      TabIndex        =   3
      Top             =   2280
      Width           =   1575
   End
   Begin VB.TextBox txtname 
      Height          =   495
      Left            =   5280
      TabIndex        =   2
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "NEW PASSWORD"
      Height          =   735
      Left            =   1800
      TabIndex        =   5
      Top             =   4440
      Width           =   2655
   End
   Begin VB.Label Label3 
      Caption         =   "NEW USER NAME"
      Height          =   615
      Left            =   1800
      TabIndex        =   4
      Top             =   3480
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "OLD PASSWORD"
      Height          =   735
      Left            =   1680
      TabIndex        =   1
      Top             =   2040
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "OLD USER NAME"
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   1080
      Width           =   2415
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    frm
If RS.State = 1 Then RS.Close
RS.Open "SELECT * FROM LOGIN WHERE USERNAME <> '" & txtname.Text & "' AND PASSWORD <> " & txtpass.Text & "", CON
 If RS.EOF = False Then
 MsgBox "username and password not valid"
 Else
 
 If RS.State = 1 Then RS.Close
 RS.Open "SELECT * FROM LOGIN WHERE USERNAME='" & txtname.Text & "' AND PASSWORD=" & txtpass.Text & "", CON
 If RS.EOF = False Then
 If txtname1.Text = "" And txtpass1.Text = "" Then
 MsgBox "PLease enter new username and password"
  Else
  
  'frm
'CON.Execute "update LOGIN set username='" & txtname1.Text & "',password=" & txtpass1.Text & " where username='" & txtname.Text & "'"
CON.Execute "update LOGIN  set USERNAME='" & txtname1.Text & "',PASSWORD=" & txtpass1.Text & " where USERNAME='" & txtname.Text & "'"
 MsgBox "Username & Password Changed"
 End If
 
 Else
 MsgBox "Access Denied"
End If
End If
End Sub
