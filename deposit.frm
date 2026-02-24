VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   7860
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11520
   LinkTopic       =   "Form3"
   ScaleHeight     =   7860
   ScaleWidth      =   11520
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      Picture         =   "deposit.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   6840
      Width           =   1335
   End
   Begin VB.TextBox balance 
      Enabled         =   0   'False
      Height          =   285
      Left            =   9120
      TabIndex        =   20
      Top             =   3000
      Width           =   1695
   End
   Begin VB.CommandButton cmdDeposit 
      Caption         =   "&Deposit"
      DownPicture     =   "deposit.frx":6B0A
      Height          =   375
      Left            =   3120
      Picture         =   "deposit.frx":895D
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6840
      Width           =   1095
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "&Quit"
      DownPicture     =   "deposit.frx":F467
      Height          =   375
      Left            =   7920
      Picture         =   "deposit.frx":112BA
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      DownPicture     =   "deposit.frx":17DC4
      Height          =   375
      Left            =   4680
      Picture         =   "deposit.frx":19C17
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6840
      Width           =   975
   End
   Begin VB.ComboBox cboCustomerNo 
      Height          =   315
      Left            =   4350
      TabIndex        =   8
      Top             =   2760
      Width           =   2055
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
   Begin VB.OptionButton optCash 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cash"
      Height          =   195
      Left            =   4350
      TabIndex        =   6
      Top             =   4710
      Width           =   975
   End
   Begin VB.TextBox txtCheckNo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4335
      TabIndex        =   5
      Top             =   5400
      Width           =   3375
   End
   Begin VB.TextBox txtAmountDeposited 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4335
      TabIndex        =   4
      Top             =   4275
      Width           =   1575
   End
   Begin VB.TextBox txtNarration 
      Appearance      =   0  'Flat
      Height          =   645
      Left            =   4335
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   3540
      Width           =   3375
   End
   Begin VB.TextBox txtAccountNo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4335
      TabIndex        =   2
      Top             =   3165
      Width           =   2295
   End
   Begin VB.TextBox txtTransactionID 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4320
      TabIndex        =   1
      Top             =   2400
      Width           =   1935
   End
   Begin MSComCtl2.DTPicker txtDated 
      Height          =   375
      Left            =   9120
      TabIndex        =   0
      Top             =   2280
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      Format          =   69271553
      CurrentDate     =   38293
   End
   Begin VB.Label Label1 
      Caption         =   "Balance"
      Height          =   255
      Left            =   7560
      TabIndex        =   22
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dated:"
      Height          =   195
      Index           =   7
      Left            =   8040
      TabIndex        =   19
      Top             =   2400
      Width           =   480
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CheckNo:"
      Height          =   195
      Index           =   6
      Left            =   3585
      TabIndex        =   18
      Top             =   5445
      Width           =   720
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mode:"
      Height          =   195
      Index           =   5
      Left            =   3855
      TabIndex        =   17
      Top             =   4710
      Width           =   450
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AmountDeposited:"
      Height          =   195
      Index           =   4
      Left            =   3000
      TabIndex        =   16
      Top             =   4320
      Width           =   1305
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Narration:"
      Height          =   195
      Index           =   3
      Left            =   3615
      TabIndex        =   15
      Top             =   3585
      Width           =   690
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AccountNo:"
      Height          =   195
      Index           =   2
      Left            =   3450
      TabIndex        =   14
      Top             =   3210
      Width           =   855
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CustomerNo:"
      Height          =   195
      Index           =   1
      Left            =   3390
      TabIndex        =   13
      Top             =   2820
      Width           =   915
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TransactionID:"
      Height          =   195
      Index           =   0
      Left            =   3255
      TabIndex        =   12
      Top             =   2445
      Width           =   1050
   End
   Begin VB.Image Image1 
      Height          =   1500
      Left            =   0
      Picture         =   "deposit.frx":20721
      Top             =   0
      Width           =   12000
   End
   Begin VB.Image Image2 
      Height          =   6180
      Left            =   0
      Picture         =   "deposit.frx":2BD9E
      Top             =   1440
      Width           =   2400
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
RS.Open "select  customerid from customer", CON
Do While Not RS.EOF
cboCustomerNo.AddItem ((RS.Fields(0).Value))
RS.MoveNext
Loop
End Sub

Private Sub cboCustomerNo_Click()
    
Dim y
If RS.State = 1 Then RS.Close
RS.Open "select * from customer where customerid= " & cboCustomerNo.Text & "", CON
If RS.EOF = False Then
txtAccountNo.Text = RS.Fields(7).Value
y = RS.Fields(7).Value
'balance.Text = RS.Fields(8).Value

End If
If RS.State = 1 Then RS.Close
RS.Open "select balance from balancedt where accno= " & Val(y) & "", CON
If RS.EOF = False Then
balance.Text = RS.Fields(0).Value
End If
'balance.Text = RS.Fields(8).Value




'frm


End Sub

Private Sub cboCustomerNo_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii < 48) And (KeyAscii > 57) Then
MsgBox "PLEASE ENTER ONLY Numbers"
KeyAscii = 0
End If
End Sub

Private Sub cmdDeposit_Click()
cmdDeposit.Enabled = False
cmdSave.Enabled = True

Command4.Enabled = True
frm



  If RS.State = 1 Then RS.Close
   RS.Open "select max(TransactionID) from  deposit", CON
  If RS.EOF = False Then
txtTransactionID.Text = RS.Fields(0) + 1

  End If
 
End Sub

Private Sub cmdQuit_Click()
Me.Hide
main.Show

End Sub

'Public mode
Private Sub cmdSave_Click()
cmdDeposit.Enabled = True

 Dim a, b, c
frm
If txtTransactionID.Text = "" Or txtDated.Value = "" Or cboCustomerNo.Text = "" Or txtAccountNo.Text = "" Or txtNarration.Text = "" Or txtAmountDeposited.Text = "" Then

'Or txtCheckNo.Text = "" Then
   MsgBox "please enter all details"
  Else
  If optCash.Value = True Then
 
    
    a = balance.Text
    b = txtAmountDeposited.Text
    c = Val(a) + Val(b)
  
  
   CON.Execute ("insert into deposit values(" & txtTransactionID.Text & "," & cboCustomerNo.Text & "," & txtAccountNo.Text & ",'" & txtNarration.Text & "'," & txtAmountDeposited.Text & ",'" & mode & "'," & 0 & ",'" & CDate(txtDated.Value) & "')")
    CON.Execute ("insert into transactions values(" & cboCustomerNo.Text & "," & txtAccountNo.Text & ",'" & txtNarration.Text & "','00','" & CDate(txtDated.Value) & "'," & txtAmountDeposited.Text & ",'0'," & c & ",'cash')")
    CON.Execute ("update balancedt set balance=" & c & " where accno=" & Val(txtAccountNo.Text) & "")
    MsgBox txtDated.Value
   MsgBox (" SAVE RECORD")
   cboCustomerNo.Clear

txtTransactionID.Text = ""
 'txtDated.Value = ""
 cboCustomerNo.Text = ""
 txtCheckNo.Text = ""
  txtAccountNo.Text = ""
  txtNarration.Text = ""
  txtAmountDeposited.Text = ""
balance.Text = ""
fill
   ElseIf optCheque.Value = True Then
 If txtTransactionID.Text = "" Or txtDated.Value = "" Or cboCustomerNo.Text = "" Or txtAccountNo.Text = "" Or txtNarration.Text = "" Or txtAmountDeposited.Text = "" Or txtCheckNo.Text = "" Then
 
   MsgBox "please enter all details"
 Else
    
     a = balance.Text
    b = txtAmountDeposited.Text
    c = Val(a) + Val(b)
   
  
   CON.Execute ("insert into deposit values(" & txtTransactionID.Text & "," & cboCustomerNo.Text & "," & txtAccountNo.Text & ",'" & txtNarration.Text & "'," & txtAmountDeposited.Text & ",'" & mode & "'," & txtCheckNo.Text & ",'" & txtDated.Value & "')")
   CON.Execute ("insert into transactions values(" & cboCustomerNo.Text & "," & txtAccountNo.Text & ",'" & txtNarration.Text & "'," & txtCheckNo & ",'" & txtDated.Value & "'," & txtAmountDeposited.Text & ",'0'," & c & ",'cheque')")
    CON.Execute ("update balancedt set balance=" & c & " where accno=" & Val(txtAccountNo.Text) & "")
    
   MsgBox (" SAVE RECORD")
  


cboCustomerNo.Clear

txtTransactionID.Text = ""
 'txtDated.Value = ""
 cboCustomerNo.Text = ""
 txtCheckNo.Text = ""
  txtAccountNo.Text = ""
  txtNarration.Text = ""
  txtAmountDeposited.Text = ""
balance.Text = ""
'optCash.Value = ""
End If
End If
fill
End If

End Sub

Private Sub Command4_Click()
cmdDeposit.Enabled = True

cboCustomerNo.Clear

txtTransactionID.Text = ""
 'txtDated.Value = ""
 cboCustomerNo.Text = ""
 txtCheckNo.Text = ""
  txtAccountNo.Text = ""
  txtNarration.Text = ""
  txtAmountDeposited.Text = ""
balance.Text = ""
fill
End Sub

Private Sub Form_Load()
fill
txtTransactionID.Enabled = False
cmdSave.Enabled = False
Command4.Enabled = False
'cboCustomerNo.Enabled = False

'cmdclear.
End Sub

Private Sub optCash_Click()
mode = "cash"
txtCheckNo.Enabled = False
End Sub

Private Sub optCheque_Click()
mode = "cheque"
txtCheckNo.Enabled = True

End Sub



Private Sub txtAccountNo_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii < 48) And (KeyAscii > 57) Then
       MsgBox "PLEASE ENTER ONLY Numbers"
       KeyAscii = 0
     End If
End Sub

Private Sub txtAmountDeposited_KeyPress(KeyAscii As Integer)
     If Not (KeyAscii < 48) And (KeyAscii > 57) Then
       MsgBox "PLEASE ENTER ONLY Numbers"
       KeyAscii = 0
     End If
End Sub

Private Sub txtCheckNo_KeyPress(KeyAscii As Integer)
     If Not (KeyAscii < 48) And (KeyAscii > 57) Then
       MsgBox "PLEASE ENTER ONLY Numbers"
       KeyAscii = 0
     End If
End Sub
