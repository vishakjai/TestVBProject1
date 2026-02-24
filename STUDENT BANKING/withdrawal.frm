VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   9870
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10875
   LinkTopic       =   "Form4"
   ScaleHeight     =   9870
   ScaleWidth      =   10875
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox balance 
      Height          =   375
      Left            =   7320
      TabIndex        =   17
      Top             =   3840
      Width           =   2055
   End
   Begin VB.TextBox txtTransactionID 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4440
      TabIndex        =   10
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox txtaccountno 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4455
      TabIndex        =   9
      Top             =   3840
      Width           =   2295
   End
   Begin VB.TextBox txtNarration 
      Appearance      =   0  'Flat
      Height          =   765
      Left            =   4455
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   4320
      Width           =   3375
   End
   Begin VB.TextBox txtAmountWithdrawn 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4455
      TabIndex        =   7
      Top             =   5400
      Width           =   1920
   End
   Begin VB.ComboBox cboCustomerno 
      Height          =   315
      Left            =   4470
      TabIndex        =   6
      Top             =   3270
      Width           =   2055
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      DownPicture     =   "withdrawal.frx":0000
      Height          =   375
      Left            =   4200
      Picture         =   "withdrawal.frx":1E53
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6000
      Width           =   975
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "&Quit"
      DownPicture     =   "withdrawal.frx":895D
      Height          =   375
      Left            =   7920
      Picture         =   "withdrawal.frx":A7B0
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      DownPicture     =   "withdrawal.frx":112BA
      Height          =   375
      Left            =   6600
      Picture         =   "withdrawal.frx":1310D
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      DownPicture     =   "withdrawal.frx":19C17
      Height          =   375
      Left            =   5280
      Picture         =   "withdrawal.frx":1BA6A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton cmdWithdraw 
      Caption         =   "&Withdraw"
      DownPicture     =   "withdrawal.frx":22574
      Height          =   375
      Left            =   3000
      Picture         =   "withdrawal.frx":243C7
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6000
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
      Format          =   63176705
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
      TabIndex        =   16
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
      TabIndex        =   15
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
      TabIndex        =   14
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
      TabIndex        =   13
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
      TabIndex        =   12
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
      TabIndex        =   11
      Top             =   2760
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   1500
      Left            =   0
      Picture         =   "withdrawal.frx":2AED1
      Top             =   0
      Width           =   12000
   End
   Begin VB.Image Image2 
      Height          =   6180
      Left            =   0
      Picture         =   "withdrawal.frx":367F4
      Top             =   1200
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
    RS.Open "select customerid from customer", CON
    Do While Not RS.EOF
      cboCustomerno.AddItem (RS.Fields(0).Value)
      RS.MoveNext
      Loop
End Sub
Private Sub cboCustomerno_Click()
     If RS.State = 1 Then RS.Close
RS.Open "select * from customer where customerid= " & cboCustomerno.Text & "", CON
If RS.EOF = False Then
txtaccountno.Text = RS.Fields(7).Value
balance.Text = RS.Fields(8).Value

End If
End Sub

Private Sub cmdQuit_Click()
 Me.Hide
 main.Show
End Sub

Private Sub cmdSave_Click()
    frm
   ' If txtTransactionID.Text = "" Or cboCustomerno.Text = "" Or text1.Text = "" Or txtNarration.Text = "" Or txtAmountWithdrawn.Text = "" Or txtDated.Value = "" Then
     '  MsgBox "plz enter all details"
   ' Else
    CON.Execute ("insert into withdrawal values(" & txtTransactionID.Text & "," & cboCustomerno.Text & "," & txtaccountno.Text & ",'" & txtNarration.Text & "'," & txtAmountWithdrawn.Text & "," & txtDated.Value & ")")
    MsgBox " save record"
   'End If
    txtTransactionID.Text = ""
    cboCustomerno.Text = ""
    txtaccountno.Text = ""
    txtNarration.Text = ""
    txtAmountWithdrawn.Text = ""
    balance.Text = ""
    
End Sub

Private Sub cmdwithdrawn_click()
    frm
    If RS.State = 1 Then RS.Close
    RS.Open "select max(transactionid) from withdrawal", CON
    cboCustomerno.Text = RS.Fields(0) + 1
    Exit Sub
        
End Sub

Private Sub cmdWithdraw_Click()
      frm
    If RS.State = 1 Then RS.Close
    RS.Open "select max(transactionid) from withdrawal", CON
   txtTransactionID.Text = RS.Fields(0) + 1
    Exit Sub
End Sub

Private Sub Form_Load()
    fill
End Sub
