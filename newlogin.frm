VERSION 5.00
Begin VB.Form Form7 
   BackColor       =   &H0080C0FF&
   Caption         =   "Form7"
   ClientHeight    =   6600
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9945
   LinkTopic       =   "Form7"
   ScaleHeight     =   6600
   ScaleWidth      =   9945
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtname 
      Height          =   372
      Left            =   4680
      TabIndex        =   5
      Top             =   1440
      Width           =   1695
   End
   Begin VB.TextBox txtpass 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   4680
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1920
      Width           =   1695
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H0000FF00&
      Caption         =   "Change"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4200
      Width           =   1695
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H000000FF&
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4200
      Width           =   1575
   End
   Begin VB.TextBox txtname1 
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   2640
      Width           =   1695
   End
   Begin VB.TextBox txtpass1 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   4680
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   " Old User Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   9
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   " Old Password"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   8
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "New User Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   7
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "New Password"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      Top             =   3360
      Width           =   1695
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcancel_Click()
Form1.Show
Me.Hide

End Sub

Private Sub cmdok_Click()
frm
If RS.State = 1 Then RS.Close
 RS.Open "SELECT * FROM logi WHERE user1 <> '" & txtname.Text & "' AND pass <> '" & txtpass.Text & "'", CON
 If RS.EOF = False Then
 MsgBox "username and password not valid"
 Else
 
 If RS.State = 1 Then RS.Close
 RS.Open "SELECT * FROM logi WHERE user1='" & txtname.Text & "' AND pass='" & txtpass.Text & "'", CON
 If RS.EOF = False Then
 If txtname1.Text = "" And txtpass1.Text = "" Then
 MsgBox "PLease enter new username and password"
  Else
 CON.Execute "update logi set user1='" & txtname1.Text & "',pass='" & txtpass1.Text & "' where user1='" & txtname.Text & "'"

 MsgBox "Username & Password Changed"
 End If
 
 Else
 MsgBox "Access Denied"
End If
End If

End Sub

