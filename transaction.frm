VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form5 
   Caption         =   "Form5"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13620
   LinkTopic       =   "Form5"
   ScaleHeight     =   11115
   ScaleWidth      =   13620
   StartUpPosition =   3  'Windows Default
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   3360
      TabIndex        =   29
      Top             =   6000
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   19791873
      CurrentDate     =   39827
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1080
      TabIndex        =   28
      Top             =   6000
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   19791873
      CurrentDate     =   39827
   End
   Begin VB.CommandButton cmddelete 
      Caption         =   "Delete"
      DownPicture     =   "transaction.frx":0000
      Height          =   375
      Left            =   6960
      Picture         =   "transaction.frx":1E53
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   5040
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   30
      Left            =   4920
      TabIndex        =   26
      Top             =   7920
      Width           =   135
      _ExtentX        =   238
      _ExtentY        =   53
      _Version        =   393216
   End
   Begin VB.TextBox txtMode 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   15
      Top             =   4920
      Width           =   2055
   End
   Begin VB.TextBox txtBalance 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7320
      TabIndex        =   14
      Top             =   4080
      Width           =   1695
   End
   Begin VB.TextBox txtCredit 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7320
      TabIndex        =   13
      Top             =   3600
      Width           =   1695
   End
   Begin VB.TextBox txtDebit 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7320
      TabIndex        =   12
      Top             =   3000
      Width           =   1695
   End
   Begin VB.TextBox txtCheckNo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   11
      Top             =   4320
      Width           =   2655
   End
   Begin VB.TextBox txtNarration 
      Appearance      =   0  'Flat
      Height          =   645
      Left            =   1440
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   3480
      Width           =   2895
   End
   Begin VB.TextBox txtAccountNo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   9
      Top             =   3000
      Width           =   2775
   End
   Begin VB.TextBox txtCustomerID 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   8
      Top             =   2400
      Width           =   2775
   End
   Begin MSFlexGridLib.MSFlexGrid msfgroom 
      Height          =   3495
      Left            =   360
      TabIndex        =   7
      Top             =   7440
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   6165
      _Version        =   393216
      Rows            =   10
      Cols            =   10
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   480
      TabIndex        =   5
      Top             =   5520
      Width           =   10215
      Begin VB.CommandButton Command1 
         Caption         =   "Proceed"
         Height          =   375
         Left            =   4800
         TabIndex        =   30
         Top             =   480
         Width           =   1575
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         DownPicture     =   "transaction.frx":895D
         Height          =   375
         Left            =   7080
         Picture         =   "transaction.frx":A7B0
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Choose the View Mode"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4560
      TabIndex        =   2
      Top             =   7440
      Width           =   5775
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Custom"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "View All"
         Height          =   255
         Left            =   1800
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "&Quit"
      DownPicture     =   "transaction.frx":112BA
      Height          =   375
      Left            =   12360
      Picture         =   "transaction.frx":1310D
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9840
      Width           =   735
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      DownPicture     =   "transaction.frx":19C17
      Height          =   375
      Left            =   4560
      Picture         =   "transaction.frx":1BA6A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5040
      Width           =   1575
   End
   Begin MSComCtl2.DTPicker txtDated 
      Height          =   375
      Left            =   7320
      TabIndex        =   16
      Top             =   2400
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      Format          =   19791873
      CurrentDate     =   38293
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mode:"
      Height          =   195
      Index           =   8
      Left            =   600
      TabIndex        =   25
      Top             =   4920
      Width           =   450
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Balance:"
      Height          =   195
      Index           =   7
      Left            =   6480
      TabIndex        =   24
      Top             =   4200
      Width           =   630
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Credit:"
      Height          =   195
      Index           =   6
      Left            =   6600
      TabIndex        =   23
      Top             =   3720
      Width           =   450
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Debit:"
      Height          =   195
      Index           =   5
      Left            =   6600
      TabIndex        =   22
      Top             =   3120
      Width           =   420
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dated:"
      Height          =   195
      Index           =   4
      Left            =   6480
      TabIndex        =   21
      Top             =   2520
      Width           =   480
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CheckNo:"
      Height          =   195
      Index           =   3
      Left            =   360
      TabIndex        =   20
      Top             =   4440
      Width           =   720
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Narration:"
      Height          =   195
      Index           =   2
      Left            =   360
      TabIndex        =   19
      Top             =   3720
      Width           =   690
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AccountNo:"
      Height          =   195
      Index           =   1
      Left            =   480
      TabIndex        =   18
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CustomerID:"
      Height          =   195
      Index           =   0
      Left            =   480
      TabIndex        =   17
      Top             =   2400
      Width           =   870
   End
   Begin VB.Image Image1 
      Height          =   1500
      Left            =   0
      Picture         =   "transaction.frx":22574
      Top             =   0
      Width           =   12000
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub DD()
txtCustomerID.Enabled = False
txtAccountNo.Enabled = False
txtNarration.Enabled = False
txtCheckNo.Enabled = False
txtDebit.Enabled = False
txtCredit.Enabled = False
txtBalance.Enabled = False
txtMode.Enabled = False
End Sub

Public Sub DD1()
txtCustomerID.Enabled = True
txtAccountNo.Enabled = True
txtNarration.Enabled = True
txtCheckNo.Enabled = True
txtDebit.Enabled = True
txtCredit.Enabled = True
txtBalance.Enabled = True
txtMode.Enabled = True
End Sub

Private Sub cmddelete_Click()
If txtCustomerID.Text = "" Then
   MsgBox "please enter all details"
Else
If CON.State = 1 Then CON.Close
 CON.Open "PROVIDER=MICROSOFT.JET.OLEDB.4.0;DATA SOURCE=" & App.Path & "/BANK1.MDB"
CON.Execute "delete from transactions where CustomerID=" & txtCustomerID & ""
MsgBox "record deleted"
End If

End Sub

Private Sub cmdEdit_Click()
'frm
If txtCustomerID.Text = "" Then
   MsgBox "please enter all details"
Else
If CON.State = 1 Then CON.Close
 CON.Open "PROVIDER=MICROSOFT.JET.OLEDB.4.0;DATA SOURCE=" & App.Path & "/BANK1.MDB"
CON.Execute "update transactions set CustomerID=" & txtCustomerID.Text & ",AccountNo=" & txtAccountNo.Text & ",Narration='" & txtNarration.Text & "',CheckNo='" & txtCheckNo.Text & "',Dated='" & txtDated.Value & "',Debit=" & txtDebit.Text & ",Credit=" & txtCredit.Text & ",Balance=" & txtBalance.Text & ",Mode='" & txtMode.Text & "' WHERE CustomerID=" & txtCustomerID.Text & ""
MsgBox "record updated"
display
End If
End Sub

Private Sub cmdok_Click()


End Sub

Private Sub cmdQuit_Click()
main.Show
End Sub

Private Sub cmdRefresh_Click()
display

End Sub

Private Sub Command1_Click()
msfgroom.Clear
'msfgroom.ColWidth(0) = 2000
'msfgroom.ColWidth(1) = 2000
'msfgroom.ColWidth(2) = 2000
'msfgroom.ColWidth(3) = 2000
'msfgroom.ColWidth(4) = 2000
'msfgroom.ColWidth(5) = 2000
'msfgroom.ColWidth(6) = 2000
'msfgroom.ColWidth(7) = 2000
'msfgroom.ColWidth(8) = 2000
msfgroom.TextMatrix(0, 0) = "CustomerId"
msfgroom.TextMatrix(0, 1) = "Account No"
msfgroom.TextMatrix(0, 2) = "Narration"
msfgroom.TextMatrix(0, 3) = "ChequeNo"
msfgroom.TextMatrix(0, 4) = "Date"
msfgroom.TextMatrix(0, 5) = "Debit"
msfgroom.TextMatrix(0, 6) = "Credit"
msfgroom.TextMatrix(0, 7) = "Balance"
msfgroom.TextMatrix(0, 8) = "Mode"
'If RS.State = 1 Then RS.Close
'RS.Open "select count(*) from transactions", CON
'msfgroom.Rows = IIf(IsNull(RS(0)), 0, RS(0)) + 25
'MsgBox CDate(DTPicker1.Value)
'MsgBox CDate(CDate(DTPicker2.Value))
If RS.State = 1 Then RS.Close
RS.Open "select * from transactions where Dated between #" & CDate(DTPicker1.Value) & "# and #" & CDate(DTPicker2.Value) & "#", CON
RNO = 1
Do While Not RS.EOF
msfgroom.TextMatrix(RNO, 0) = RS(0)
msfgroom.TextMatrix(RNO, 1) = RS(1)
msfgroom.TextMatrix(RNO, 2) = RS(2)
msfgroom.TextMatrix(RNO, 3) = RS(3)
msfgroom.TextMatrix(RNO, 4) = CDate(RS.Fields(4).Value)
msfgroom.TextMatrix(RNO, 5) = RS(5)
msfgroom.TextMatrix(RNO, 6) = RS(6)
msfgroom.TextMatrix(RNO, 7) = RS(7)
msfgroom.TextMatrix(RNO, 8) = RS(8)
RNO = RNO + 1
RS.MoveNext
Loop
End Sub

Private Sub Form_Load()
If CON.State = 1 Then CON.Close
 CON.Open "PROVIDER=MICROSOFT.JET.OLEDB.4.0;DATA SOURCE=" & App.Path & "/BANk1.MDB"
display

DD
'If RS.State = 1 Then RS.Close
'RS.Open "select "
End Sub
Sub display()
msfgroom.Clear
msfgroom.ColWidth(0) = 2000
msfgroom.ColWidth(1) = 2000
msfgroom.ColWidth(2) = 2000
msfgroom.ColWidth(3) = 2000
msfgroom.ColWidth(4) = 2000
msfgroom.ColWidth(5) = 2000
msfgroom.ColWidth(6) = 2000
msfgroom.ColWidth(7) = 2000
msfgroom.ColWidth(8) = 2000
msfgroom.TextMatrix(0, 0) = "CustomerId"
msfgroom.TextMatrix(0, 1) = "Account No"
msfgroom.TextMatrix(0, 2) = "Narration"
msfgroom.TextMatrix(0, 3) = "ChequeNo"
msfgroom.TextMatrix(0, 4) = "Date"
msfgroom.TextMatrix(0, 5) = "Debit"
msfgroom.TextMatrix(0, 6) = "Credit"
msfgroom.TextMatrix(0, 7) = "Balance"
msfgroom.TextMatrix(0, 8) = "Mode"
If RS.State = 1 Then RS.Close
RS.Open "select count(*) from transactions", CON
msfgroom.Rows = IIf(IsNull(RS(0)), 0, RS(0)) + 25
If RS.State = 1 Then RS.Close
RS.Open "select * from transactions", CON
RNO = 1
Do While Not RS.EOF
msfgroom.TextMatrix(RNO, 0) = RS(0)
msfgroom.TextMatrix(RNO, 1) = RS(1)
msfgroom.TextMatrix(RNO, 2) = RS(2)
msfgroom.TextMatrix(RNO, 3) = RS(3)
msfgroom.TextMatrix(RNO, 4) = RS(4)
msfgroom.TextMatrix(RNO, 5) = RS(5)
msfgroom.TextMatrix(RNO, 6) = RS(6)
msfgroom.TextMatrix(RNO, 7) = RS(7)
msfgroom.TextMatrix(RNO, 8) = RS(8)
RNO = RNO + 1
RS.MoveNext
Loop
End Sub

Private Sub msfgroom_Click()
If Not msfgroom.TextMatrix(msfgroom.Row, 0) = "" Then
z = msfgroom.TextMatrix(msfgroom.Row, 0)
txtCustomerID.Text = msfgroom.TextMatrix(msfgroom.Row, 0)
txtAccountNo.Text = msfgroom.TextMatrix(msfgroom.Row, 1)
txtNarration.Text = msfgroom.TextMatrix(msfgroom.Row, 2)
txtCheckNo.Text = msfgroom.TextMatrix(msfgroom.Row, 3)
txtDated.Value = msfgroom.TextMatrix(msfgroom.Row, 4)
txtDebit.Text = msfgroom.TextMatrix(msfgroom.Row, 5)
txtCredit.Text = msfgroom.TextMatrix(msfgroom.Row, 6)
txtBalance.Text = msfgroom.TextMatrix(msfgroom.Row, 7)
txtMode.Text = msfgroom.TextMatrix(msfgroom.Row, 8)
End If
'If Not msfgroom.TextMatrix(msfgroom.Row, 0) = "" Then
'z = msfgroom.TextMatrix(msfgroom.Row, 0)
'TXTSlNO.Text = msfgroom.TextMatrix(msfgroom.Row, 0)
'TXTRNO.Text = msfgroom.TextMatrix(msfgroom.Row, 1)
'CMBRM_TYPE.Text = msfgroom.TextMatrix(msfgroom.Row, 2)
'CMBRM_DESC.Text = msfgroom.TextMatrix(msfgroom.Row, 3)
'TXTBED.Text = msfgroom.TextMatrix(msfgroom.Row, 4)
'End If
DD1


End Sub

Private Sub msfgroom_DblClick()
'Form7.Show
End Sub

