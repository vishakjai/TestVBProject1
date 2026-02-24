VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form3 
   ClientHeight    =   8370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9495
   LinkTopic       =   "Form3"
   ScaleHeight     =   8370
   ScaleWidth      =   9495
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtaccountno 
      Height          =   285
      Left            =   4440
      TabIndex        =   22
      Top             =   3240
      Width           =   1815
   End
   Begin VB.TextBox txtTransactionID 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4440
      TabIndex        =   12
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox txtNarration 
      Appearance      =   0  'Flat
      Height          =   645
      Left            =   4335
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   3540
      Width           =   3375
   End
   Begin VB.TextBox txtAmountDeposited 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4335
      TabIndex        =   10
      Top             =   4275
      Width           =   1575
   End
   Begin VB.TextBox txtCheckNo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4335
      TabIndex        =   9
      Top             =   5400
      Width           =   3375
   End
   Begin VB.OptionButton optCash 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cash"
      Height          =   195
      Left            =   4350
      TabIndex        =   8
      Top             =   4710
      Width           =   975
   End
   Begin VB.OptionButton optCheque 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cheque"
      Height          =   195
      Left            =   5430
      TabIndex        =   7
      Top             =   4710
      Width           =   1095
   End
   Begin VB.ComboBox cboCustomerNo 
      Height          =   315
      Left            =   4350
      TabIndex        =   6
      Top             =   2760
      Width           =   2055
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      DownPicture     =   "deposit.frx":0000
      Height          =   375
      Left            =   4080
      Picture         =   "deposit.frx":1E53
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6480
      Width           =   975
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "&Quit"
      DownPicture     =   "deposit.frx":895D
      Height          =   375
      Left            =   7800
      Picture         =   "deposit.frx":A7B0
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      DownPicture     =   "deposit.frx":112BA
      Height          =   375
      Left            =   6480
      Picture         =   "deposit.frx":1310D
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      DownPicture     =   "deposit.frx":19C17
      Height          =   375
      Left            =   5160
      Picture         =   "deposit.frx":1BA6A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton cmdDeposit 
      Caption         =   "&Deposit"
      DownPicture     =   "deposit.frx":22574
      Height          =   375
      Left            =   2880
      Picture         =   "deposit.frx":243C7
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6480
      Width           =   1095
   End
   Begin VB.TextBox balance 
      Height          =   285
      Left            =   7320
      TabIndex        =   0
      Top             =   3000
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker txtDated 
      Height          =   375
      Left            =   6960
      TabIndex        =   13
      Top             =   2400
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      Format          =   63176705
      CurrentDate     =   38293
   End
   Begin VB.Image Image2 
      Height          =   6180
      Left            =   0
      Picture         =   "deposit.frx":2AED1
      Top             =   1440
      Width           =   2400
   End
   Begin VB.Image Image1 
      Height          =   1500
      Left            =   0
      Picture         =   "deposit.frx":376B1
      Top             =   0
      Width           =   12000
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TransactionID:"
      Height          =   195
      Index           =   0
      Left            =   3255
      TabIndex        =   21
      Top             =   2445
      Width           =   1050
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CustomerNo:"
      Height          =   195
      Index           =   1
      Left            =   3390
      TabIndex        =   20
      Top             =   2820
      Width           =   915
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AccountNo:"
      Height          =   195
      Index           =   2
      Left            =   3450
      TabIndex        =   19
      Top             =   3210
      Width           =   855
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Narration:"
      Height          =   195
      Index           =   3
      Left            =   3615
      TabIndex        =   18
      Top             =   3585
      Width           =   690
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AmountDeposited:"
      Height          =   195
      Index           =   4
      Left            =   3000
      TabIndex        =   17
      Top             =   4320
      Width           =   1305
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mode:"
      Height          =   195
      Index           =   5
      Left            =   3855
      TabIndex        =   16
      Top             =   4710
      Width           =   450
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CheckNo:"
      Height          =   195
      Index           =   6
      Left            =   3585
      TabIndex        =   15
      Top             =   5445
      Width           =   720
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dated:"
      Height          =   195
      Index           =   7
      Left            =   6345
      TabIndex        =   14
      Top             =   2460
      Width           =   480
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public mode As String
Public Sub fill()
   frm
   If RS.State = 1 Then RS.Close
    RS.Open "select customerid from customer", CON
    Do While Not RS.EOF
      cboCustomerno.AddItem (RS.Fields(0).Value)
      RS.MoveNext
    Loop
End Sub

Private Sub cboCustomerNo_Click()
    If RS.State = 1 Then RS.Close
RS.Open "select * from customer where customerid= " & cboCustomerno.Text & "", CON
If RS.EOF = False Then
txtaccountno.Text = RS.Fields(7).Value
balance.Text = RS.Fields(8).Value

End If
End Sub


Private Sub cmdDeposit_Click()
    frm
    If RS.State = 1 Then RS.Close
    RS.Open "select max(transactionid) from deposit", CON
   txtTransactionID.Text = RS.Fields(0) + 1
    Exit Sub
End Sub

Private Sub cmdQuit_Click()
Me.Hide
main.Show
End Sub

Private Sub cmdSave_Click()
  frm
  If optCash.Value = True Then
 
   CON.Execute ("insert into deposit values(" & txtTransactionID.Text & "," & cboCustomerno.Text & "," & txtaccountno.Text & ",'" & txtNarration.Text & "'," & txtAmountDeposited.Text & ",'" & mode & "','" & 0 & "','" & txtDated.Value & "')")
   MsgBox " record saved"
  ElseIf optCheque.Value = True Then
          CON.Execute ("insert into deposit values(" & txtTransactionID.Text & "," & cboCustomerno.Text & "," & txtaccountno.Text & ",'" & txtNarration.Text & "'," & txtAmountDeposited.Text & ",'" & mode & "'," & txtCheckNo.Text & ",'" & txtDated.Value & "')")
          MsgBox " record saved"
  End If
  
  
 'If txtTransactionID.Text = "" Or txtDated.Value = "" Or cboCustomerNo.Text = "" Or Text1.Text = "" Or txtNarration.Text = "" Or txtAmountDeposited.Text = "" Or optCash.Value = "" Or optCheque.Value = "" Or optOthers.Value = "" Or txtchequeno.Text = "" Then
   'MsgBox "plz ente'r all details"
 'Else
' CON.Execute ("insert into deposit values(" & txtTransactionID.Text & "," & cboCustomerNo.Text & "," & text1.Text & ",'" & txtNarration.Text & "'," & txtAmountDeposited.Text & ",'" & txtMode.Text & "'," & txtCheckNo.Text & ",'" & txtDated.Value & "')")
     'CON.Execute ("insert into deposit values(" & txtTransactionID.Text & "," & cboCustomerNo.Text & "," & txtAccountNo.Text & ",'" & txtNarration.Text & "'," & txtAmountDeposited.Text & ",'" & txtMode.Text & "'," & txtchequeno.Text & ",'" & txtDated.Value & "')")
   '  MsgBox ("SAVE RECORD")
  'End If
  
 
End Sub

Private Sub Text1_Change()

End Sub

Private Sub Form_Load()
fill
'txtChequeNo.Enabled = False

End Sub

Private Sub optCash_Click()
mode = "cash"
 txtCheckNo.Enabled = False
End Sub

Private Sub optCheque_Click()
    mode = "cheque"
    
End Sub

