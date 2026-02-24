VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form7 
   Caption         =   "Form7"
   ClientHeight    =   5385
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6330
   LinkTopic       =   "Form7"
   ScaleHeight     =   5385
   ScaleWidth      =   6330
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtCode 
      Appearance      =   0  'Flat
      Height          =   225
      Left            =   4200
      TabIndex        =   11
      Top             =   405
      Width           =   495
   End
   Begin VB.TextBox txtCustomerID 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1425
      TabIndex        =   10
      Top             =   360
      Width           =   3375
   End
   Begin VB.TextBox txtAccountNo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1425
      TabIndex        =   9
      Top             =   735
      Width           =   3375
   End
   Begin VB.TextBox txtNarration 
      Appearance      =   0  'Flat
      Height          =   645
      Left            =   1440
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   1125
      Width           =   3375
   End
   Begin VB.TextBox txtCheckNo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   7
      Top             =   1920
      Width           =   3375
   End
   Begin VB.TextBox txtDebit 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1425
      TabIndex        =   6
      Top             =   2625
      Width           =   1695
   End
   Begin VB.TextBox txtCredit 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1425
      TabIndex        =   5
      Top             =   3000
      Width           =   1695
   End
   Begin VB.TextBox txtBalance 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1425
      TabIndex        =   4
      Top             =   3375
      Width           =   1695
   End
   Begin VB.TextBox txtMode 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1425
      TabIndex        =   3
      Top             =   3765
      Width           =   2055
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      DownPicture     =   "newtransction.frx":0000
      Height          =   375
      Left            =   1320
      Picture         =   "newtransction.frx":1E53
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4155
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      DownPicture     =   "newtransction.frx":895D
      Height          =   375
      Left            =   2640
      Picture         =   "newtransction.frx":A7B0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4155
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker txtDated 
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   2235
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   19857409
      CurrentDate     =   38293
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CustomerID:"
      Height          =   195
      Index           =   0
      Left            =   480
      TabIndex        =   21
      Top             =   405
      Width           =   870
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AccountNo:"
      Height          =   195
      Index           =   1
      Left            =   540
      TabIndex        =   20
      Top             =   780
      Width           =   855
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Narration:"
      Height          =   195
      Index           =   2
      Left            =   705
      TabIndex        =   19
      Top             =   1170
      Width           =   690
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CheckNo:"
      Height          =   195
      Index           =   3
      Left            =   675
      TabIndex        =   18
      Top             =   1905
      Width           =   720
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dated:"
      Height          =   195
      Index           =   4
      Left            =   915
      TabIndex        =   17
      Top             =   2280
      Width           =   480
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Debit:"
      Height          =   195
      Index           =   5
      Left            =   975
      TabIndex        =   16
      Top             =   2670
      Width           =   420
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Credit:"
      Height          =   195
      Index           =   6
      Left            =   945
      TabIndex        =   15
      Top             =   3045
      Width           =   450
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Balance:"
      Height          =   195
      Index           =   7
      Left            =   765
      TabIndex        =   14
      Top             =   3420
      Width           =   630
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mode:"
      Height          =   195
      Index           =   8
      Left            =   945
      TabIndex        =   13
      Top             =   3810
      Width           =   450
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Code:"
      Height          =   195
      Index           =   9
      Left            =   3690
      TabIndex        =   12
      Top             =   435
      Width           =   420
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSave_Click()
frm
CON.Execute ("insert into transaction values (" & txtCustomerID.Text & "," & txtAccountNo.Text & ",'" & txtNarration.Text & "'," & txtCheckNo.Text & "," & txtDated.Value & "," & txtDebit.Text & "," & txtCredit.Text & "," & txtBalance.Text & ",'" & txtMode.Text & "')")

'CON.Execute ("insert into deposit values(" & txtTransactionID.Text & "," & cboCustomerNo.Text & "," & txtAccountNo.Text & ",'" & txtNarration.Text & "'," & txtAmountDeposited.Text & ",'" & mode & "'," & txtCheckNo.Text & ",'" & txtDated.Value & "')")
   
   MsgBox (" SAVE RECORD")


End Sub

Private Sub Form_Load()
'If Not msfgroom.TextMatrix(msfgroom.Row, 0) = "" Then
'z = msfgroom.TextMatrix(msfgroom.Row, 0)
'txtCustomerID.Text = msfgroom.TextMatrix(msfgroom.Row, 0)
'txtAccountNo.Text = msfgroom.TextMatrix(msfgroom.Row, 1)
'txtNarration.Text = msfgroom.TextMatrix(msfgroom.Row, 2)
'txtCheckNo.Text = msfgroom.TextMatrix(msfgroom.Row, 3)
'txtDated.Value = msfgroom.TextMatrix(msfgroom.Row, 4)
'txtDebit.Text = msfgroom.TextMatrix(msfgroom.Row, 5)
'txtCredit.Text = msfgroom.TextMatrix(msfgroom.Row, 6)
'txtBalance.Text = msfgroom.TextMatrix(msfgroom.Row, 7)
'txtMode.Text = msfgroom.TextMatrix(msfgroom.Row, 8)
'End If
End Sub
