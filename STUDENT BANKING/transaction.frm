VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form6 
   Caption         =   "Form6"
   ClientHeight    =   8400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10605
   LinkTopic       =   "Form6"
   ScaleHeight     =   8400
   ScaleWidth      =   10605
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid msfgroom 
      Height          =   2775
      Left            =   960
      TabIndex        =   19
      Top             =   4680
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   4895
      _Version        =   393216
      Rows            =   25
      Cols            =   10
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      DownPicture     =   "transaction.frx":0000
      Height          =   375
      Left            =   2400
      Picture         =   "transaction.frx":1E53
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3360
      Width           =   1575
   End
   Begin VB.TextBox txtCustomerID 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1080
      TabIndex        =   7
      Top             =   120
      Width           =   2775
   End
   Begin VB.TextBox txtAccountNo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1080
      TabIndex        =   6
      Top             =   600
      Width           =   2775
   End
   Begin VB.TextBox txtNarration 
      Appearance      =   0  'Flat
      Height          =   645
      Left            =   1080
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   1080
      Width           =   2895
   End
   Begin VB.TextBox txtCheckNo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1080
      TabIndex        =   4
      Top             =   1920
      Width           =   2655
   End
   Begin VB.TextBox txtDebit 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6960
      TabIndex        =   3
      Top             =   600
      Width           =   1695
   End
   Begin VB.TextBox txtCredit 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6960
      TabIndex        =   2
      Top             =   1200
      Width           =   1695
   End
   Begin VB.TextBox txtBalance 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6960
      TabIndex        =   1
      Top             =   1680
      Width           =   1695
   End
   Begin VB.TextBox txtMode 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Top             =   2520
      Width           =   2055
   End
   Begin MSComCtl2.DTPicker txtDated 
      Height          =   375
      Left            =   6960
      TabIndex        =   9
      Top             =   0
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   19726337
      CurrentDate     =   38293
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CustomerID:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   870
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AccountNo:"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   17
      Top             =   600
      Width           =   855
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Narration:"
      Height          =   195
      Index           =   2
      Left            =   0
      TabIndex        =   16
      Top             =   1320
      Width           =   690
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CheckNo:"
      Height          =   195
      Index           =   3
      Left            =   0
      TabIndex        =   15
      Top             =   2040
      Width           =   720
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dated:"
      Height          =   195
      Index           =   4
      Left            =   6120
      TabIndex        =   14
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Debit:"
      Height          =   195
      Index           =   5
      Left            =   6240
      TabIndex        =   13
      Top             =   720
      Width           =   420
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Credit:"
      Height          =   195
      Index           =   6
      Left            =   6240
      TabIndex        =   12
      Top             =   1320
      Width           =   450
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Balance:"
      Height          =   195
      Index           =   7
      Left            =   6120
      TabIndex        =   11
      Top             =   1800
      Width           =   630
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mode:"
      Height          =   195
      Index           =   8
      Left            =   240
      TabIndex        =   10
      Top             =   2520
      Width           =   450
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
  frm
display

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
RS.Open "select count(*) from transctions", CON '
msfgroom.Rows = IIf(IsNull(RS(0)), 0, RS(0)) + 25
If RS.State = 1 Then RS.Close
RS.Open "select * from transctions", CON
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
