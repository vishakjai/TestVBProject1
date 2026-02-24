VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   7590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11475
   LinkTopic       =   "Form4"
   ScaleHeight     =   7590
   ScaleWidth      =   11475
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtBalance 
      Enabled         =   0   'False
      Height          =   285
      Left            =   7320
      TabIndex        =   16
      Top             =   3600
      Width           =   1575
   End
   Begin VB.TextBox txtTransactionID 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4440
      TabIndex        =   9
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox txtAccountNo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4455
      TabIndex        =   8
      Top             =   3840
      Width           =   2295
   End
   Begin VB.TextBox txtNarration 
      Appearance      =   0  'Flat
      Height          =   765
      Left            =   4455
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   4320
      Width           =   3375
   End
   Begin VB.TextBox txtAmountWithdrawn 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4455
      TabIndex        =   6
      Top             =   5400
      Width           =   1920
   End
   Begin VB.ComboBox cboCustomerNo 
      Height          =   315
      Left            =   4440
      TabIndex        =   5
      Top             =   3270
      Width           =   2055
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      DownPicture     =   "withdrawal.frx":0000
      Height          =   375
      Left            =   5400
      Picture         =   "withdrawal.frx":1E53
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6120
      Width           =   975
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "&Quit"
      DownPicture     =   "withdrawal.frx":895D
      Height          =   375
      Left            =   8520
      Picture         =   "withdrawal.frx":A7B0
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Clear"
      DownPicture     =   "withdrawal.frx":112BA
      Height          =   375
      Left            =   6960
      Picture         =   "withdrawal.frx":1310D
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton cmdWithdraw 
      Caption         =   "&Withdraw"
      DownPicture     =   "withdrawal.frx":19C17
      Height          =   375
      Left            =   3720
      Picture         =   "withdrawal.frx":1BA6A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6120
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker txtDated 
      Height          =   375
      Left            =   6960
      TabIndex        =   0
      Top             =   2760
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      _Version        =   393216
      Format          =   69337089
      CurrentDate     =   38293
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TransactionID:"
      Height          =   195
      Index           =   0
      Left            =   3375
      TabIndex        =   15
      Top             =   2760
      Width           =   1050
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CustomerNo:"
      Height          =   195
      Index           =   1
      Left            =   3510
      TabIndex        =   14
      Top             =   3300
      Width           =   915
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AccountNo:"
      Height          =   195
      Index           =   2
      Left            =   3570
      TabIndex        =   13
      Top             =   3840
      Width           =   855
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Narration:"
      Height          =   195
      Index           =   3
      Left            =   3720
      TabIndex        =   12
      Top             =   4560
      Width           =   690
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AmountWithdrawn:"
      Height          =   195
      Index           =   4
      Left            =   3075
      TabIndex        =   11
      Top             =   5400
      Width           =   1350
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dated:"
      Height          =   195
      Index           =   7
      Left            =   6465
      TabIndex        =   10
      Top             =   2760
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   1500
      Left            =   0
      Picture         =   "withdrawal.frx":22574
      Top             =   0
      Width           =   12000
   End
   Begin VB.Image Image2 
      Height          =   6180
      Left            =   0
      Picture         =   "withdrawal.frx":2DE97
      Top             =   1440
      Width           =   2400
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub fill()


 frm
If RS.State = 1 Then RS.Close
RS.Open "select  customerid from customer", CON
Do While Not RS.EOF
cboCustomerNo.AddItem ((RS.Fields(0).Value))
RS.MoveNext
Loop
End Sub
Public Sub clr()
txtTransactionID.Text = ""
cboCustomerNo.Text = ""
txtAccountNo.Text = ""
txtNarration.Text = ""
txtAmountWithdrawn.Text = ""
txtBalance.Text = ""

End Sub
Private Sub cboCustomerNo_Click()
Dim y
If RS.State = 1 Then RS.Close
RS.Open "select * from customer where customerid= " & cboCustomerNo.Text & "", CON
If RS.EOF = False Then
txtAccountNo.Text = RS.Fields(7).Value
y = txtAccountNo.Text
'txtBalance.Text = RS.Fields(8).Value

End If
If RS.State = 1 Then RS.Close
RS.Open "select balance from balancedt where accno= " & Val(y) & "", CON
If RS.EOF = False Then
txtBalance.Text = RS.Fields(0).Value
End If

End Sub

Private Sub cboCustomerNo_KeyPress(KeyAscii As Integer)
     If Not (KeyAscii < 48) And (KeyAscii > 57) Then
       MsgBox "PLEASE ENTER ONLY Numbers"
       KeyAscii = 0
     End If
End Sub

Private Sub cmdcancel_Click()
clr
End Sub

Private Sub cmdQuit_Click()
Me.Hide
main.Show

End Sub

Private Sub cmdSave_Click()
cmdWithdraw.Enabled = True
 frm
    If txtTransactionID.Text = "" Or cboCustomerNo.Text = "" Or txtAccountNo.Text = "" Or txtNarration.Text = "" Or txtAmountWithdrawn.Text = "" Or txtDated.Value = "" Then
       MsgBox "plz enter all details"
    Else
    Dim a, b, c
    
    a = txtBalance.Text
    b = txtAmountWithdrawn.Text
    c = Val(a) - Val(b)
    
    CON.Execute ("insert into withdrawal values(" & txtTransactionID.Text & "," & cboCustomerNo.Text & "," & txtAccountNo.Text & ",'" & txtNarration.Text & "'," & txtAmountWithdrawn.Text & ",'" & txtDated.Value & "')")
    CON.Execute ("insert into transactions values(" & cboCustomerNo.Text & "," & txtAccountNo.Text & ",'" & txtNarration.Text & "','00','" & txtDated.Value & "','0'," & txtAmountWithdrawn.Text & "," & c & ",'Cash')")
    CON.Execute ("update balancedt set balance=" & c & " where accno=" & Val(txtAccountNo.Text) & "")
    MsgBox " save record"
    txtBalance.Text = c
   clr
   fill
   End If
End Sub

Private Sub cmdWithdraw_Click()

cmdSave.Enabled = True
cmdcancel.Enabled = True

 frm
    If RS.State = 1 Then RS.Close
    RS.Open "select max(transactionid) from withdrawal", CON
    txtTransactionID.Text = RS.Fields(0) + 1
    Exit Sub
        
End Sub

Private Sub Form_Load()
cmdSave.Enabled = False
cmdcancel.Enabled = False

fill



End Sub


Private Sub txtAccountNo_Click()
    If Not (KeyAscii < 48) And (KeyAscii > 57) Then
MsgBox "PLEASE ENTER ONLY Numbers"
KeyAscii = 0
End If
End Sub

Private Sub txtAccountNo_KeyPress(KeyAscii As Integer)
     If Not (KeyAscii < 48) And (KeyAscii > 57) Then
       MsgBox "PLEASE ENTER ONLY Numbers"
       KeyAscii = 0
     End If
End Sub

Private Sub txtAmountWithdrawn_Change()
Dim dig$, i, digi$, digits$
If txtAmountWithdrawn.Text <> "" Then
    dig$ = Mid(txtAmountWithdrawn.Text, Len(txtAmountWithdrawn.Text), 1)
    If Asc(dig$) < 46 Or Asc(dig$) > 57 Then
        For i = 1 To Len(txtAmountWithdrawn.Text) - 1
            digi$ = Mid(txtAmountWithdrawn.Text, i, 1)
            digits$ = digits$ & digi$
        Next i
     txtAmountWithdrawn.Text = digits$
       txtAmountWithdrawn.SelStart = Len(txtAmountWithdrawn.Text)
    End If
End If
End Sub

Private Sub txtAmountWithdrawn_KeyPress(KeyAscii As Integer)
     If Not (KeyAscii < 48) And (KeyAscii > 57) Then
       MsgBox "PLEASE ENTER ONLY Numbers"
       KeyAscii = 0
     End If
End Sub
