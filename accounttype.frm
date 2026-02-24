VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H80000009&
   Caption         =   "Form6"
   ClientHeight    =   9240
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10515
   LinkTopic       =   "Form6"
   Picture         =   "accounttype.frx":0000
   ScaleHeight     =   9240
   ScaleWidth      =   10515
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtAccountID 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5070
      TabIndex        =   14
      Top             =   2880
      Width           =   1335
   End
   Begin VB.TextBox txtAccountName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5070
      TabIndex        =   13
      Top             =   3270
      Width           =   3375
   End
   Begin VB.TextBox txtDescription 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5070
      TabIndex        =   12
      Top             =   3660
      Width           =   3375
   End
   Begin VB.TextBox txtInterestRate 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5040
      TabIndex        =   11
      Top             =   4035
      Width           =   1455
   End
   Begin VB.TextBox txtMinBalance 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5070
      TabIndex        =   10
      Top             =   4410
      Width           =   1935
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   3720
      TabIndex        =   0
      Top             =   4800
      Width           =   5535
      Begin VB.CommandButton cmdLast 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   4920
         Picture         =   "accounttype.frx":4CEF
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   600
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   4560
         Picture         =   "accounttype.frx":5031
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   600
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   600
         Picture         =   "accounttype.frx":5373
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   600
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   240
         Picture         =   "accounttype.frx":56B5
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   600
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add New"
         DownPicture     =   "accounttype.frx":59F7
         Height          =   375
         Left            =   240
         Picture         =   "accounttype.frx":784A
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         DownPicture     =   "accounttype.frx":E354
         Enabled         =   0   'False
         Height          =   375
         Left            =   1440
         Picture         =   "accounttype.frx":101A7
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         DownPicture     =   "accounttype.frx":16CB1
         Enabled         =   0   'False
         Height          =   375
         Left            =   2520
         Picture         =   "accounttype.frx":18B04
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdQuit 
         Caption         =   "&Quit"
         DownPicture     =   "accounttype.frx":1F60E
         Height          =   375
         Left            =   3960
         Picture         =   "accounttype.frx":21461
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   975
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   960
         TabIndex        =   9
         Top             =   600
         Width           =   3600
      End
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "AccountID:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   4065
      TabIndex        =   19
      Top             =   2940
      Width           =   975
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "AccountName:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   3780
      TabIndex        =   18
      Top             =   3315
      Width           =   1260
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Description:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   4005
      TabIndex        =   17
      Top             =   3705
      Width           =   1035
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "InterestRate:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   3
      Left            =   3915
      TabIndex        =   16
      Top             =   4080
      Width           =   1125
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "MinBalance:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   4
      Left            =   3975
      TabIndex        =   15
      Top             =   4455
      Width           =   1065
   End
   Begin VB.Image Image1 
      Height          =   1500
      Left            =   0
      Picture         =   "accounttype.frx":27F6B
      Top             =   0
      Width           =   12000
   End
   Begin VB.Image Image2 
      Height          =   6180
      Left            =   0
      Picture         =   "accounttype.frx":2F621
      Top             =   1560
      Width           =   2400
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub clr()
txtAccountID.Text = ""
 txtAccountName.Text = ""
 txtDescription.Text = ""
 txtInterestRate.Text = ""
 txtMinBalance.Text = ""
End Sub
Private Sub cmdAdd_Click()
cmdSave.Enabled = True
cmdEdit.Enabled = True

frm
  If RS.State = 1 Then RS.Close
   RS.Open "select max(AccountID) from  accounttype", CON
  If RS.EOF = False Then
txtAccountID.Text = RS.Fields(0) + 1

  End If
   txtAccountName.Text = ""
    txtDescription.Text = ""
    txtInterestRate.Text = ""
    txtMinBalance.Text = ""
End Sub

Private Sub cmdEdit_Click()
'frm
 CON.Execute ("update accounttype set AccountID=" & txtAccountID.Text & ",AccountName='" & txtAccountName.Text & "',Description='" & txtDescription.Text & "',InterestRate=" & txtInterestRate.Text & ",MinBalance=" & txtMinBalance.Text & " where AccountID=" & txtAccountID.Text & "")
MsgBox "record updated"
End Sub
'txtAccountID.Text & ",'" & txtAccountName.Text & "','" & txtDescription.Text & "'," & txtInterestRate.Text & "," & txtMinBalance.Text & ")")
Private Sub cmdFirst_Click()
On Error GoTo s
RS.MoveFirst

If Not RS.EOF = True Then
txtAccountID.Text = RS(0).Value
txtAccountName.Text = RS(1).Value
txtDescription.Text = RS(2).Value
txtInterestRate.Text = RS(3).Value
txtMinBalance.Text = RS(4).Value
End If
Exit Sub
s:
MsgBox "this is last record", vbInformation
End Sub



Private Sub cmdLast_Click()
On Error GoTo k
RS.MoveLast
If Not RS.EOF = True Then
txtAccountID.Text = RS(0).Value
txtAccountName.Text = RS(1).Value
txtDescription.Text = RS(2).Value
txtInterestRate.Text = RS(3).Value
txtMinBalance.Text = RS(4).Value
End If
Exit Sub
k:
MsgBox "This is the last file", vbInformation
End Sub

Private Sub cmdNext_Click()
On Error GoTo s
RS.MoveNext
If Not RS.EOF = True Then
txtAccountID.Text = RS(0).Value
txtAccountName.Text = RS(1).Value
txtDescription.Text = RS(2).Value
txtInterestRate.Text = RS(3).Value
txtMinBalance.Text = RS(4).Value
End If
Exit Sub
s:
MsgBox "This is the last file", vbInformation
End Sub

Private Sub cmdPrevious_Click()
On Error GoTo s
RS.MovePrevious


If Not RS.EOF = True Then
txtAccountID.Text = RS(0).Value
txtAccountName.Text = RS(1).Value
txtDescription.Text = RS(2).Value
txtInterestRate.Text = RS(3).Value
txtMinBalance.Text = RS(4).Value
End If
Exit Sub
s:
MsgBox "This is the first file", vbInformation
End Sub

Private Sub cmdQuit_Click()
main.Show
Me.Hide

End Sub

Private Sub cmdSave_Click()
frm
 If txtAccountID.Text = "" Or txtAccountName.Text = "" Or txtDescription.Text = "" Or txtInterestRate.Text = "" Or txtMinBalance.Text = "" Then
       MsgBox "plz enter all details"
    Else
    CON.Execute ("insert into accounttype values(" & txtAccountID.Text & ",'" & txtAccountName.Text & "','" & txtDescription.Text & "'," & txtInterestRate.Text & "," & txtMinBalance.Text & ")")
    MsgBox " save record"
   End If
   clr
   tt
End Sub

Public Sub tt()
frm
If RS.State = 1 Then RS.Close
RS.Open "select * from accounttype", CON, adOpenDynamic, adLockOptimistic
If Not RS.EOF = True Then
txtAccountID.Text = RS(0).Value
txtAccountName.Text = RS(1).Value
txtDescription.Text = RS(2).Value
txtInterestRate.Text = RS(3).Value
txtMinBalance.Text = RS(4).Value
End If
End Sub
Private Sub Form_Load()
tt
End Sub



Private Sub txtAccountID_KeyPress(KeyAscii As Integer)
     If Not (KeyAscii < 48) And (KeyAscii > 57) Then
       MsgBox "PLEASE ENTER ONLY Numbers"
       KeyAscii = 0
     End If
End Sub

Private Sub txtAccountName_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii < 48) And Not (KeyAscii > 57) Then
MsgBox "PLEASE ENTER ALL ONLY ALPHABETIC CHARACTERS"
KeyAscii = 0
End If
End Sub

Private Sub txtDescription_KeyPress(KeyAscii As Integer)
       If Not (KeyAscii < 48) And Not (KeyAscii > 57) Then
MsgBox "PLEASE ENTER ALL ONLY ALPHABETIC CHARACTERS"
KeyAscii = 0
End If
End Sub

Private Sub txtInterestRate_Change()
Dim dig$, i, digi$, digits$
If txtInterestRate.Text <> "" Then
    dig$ = Mid(txtInterestRate.Text, Len(txtInterestRate.Text), 1)
    If Asc(dig$) < 46 Or Asc(dig$) > 57 Then
        For i = 1 To Len(txtInterestRate.Text) - 1
            digi$ = Mid(txtInterestRate.Text, i, 1)
            digits$ = digits$ & digi$
        Next i
     txtInterestRate.Text = digits$
       txtInterestRate.SelStart = Len(txtInterestRate.Text)
    End If
    End If
    
End Sub

Private Sub txtInterestRate_KeyPress(KeyAscii As Integer)
     If Not (KeyAscii < 48) And (KeyAscii > 57) Then
       MsgBox "PLEASE ENTER ONLY Numbers"
       KeyAscii = 0
     End If
End Sub

Private Sub txtMinBalance_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii < 48) And (KeyAscii > 57) Then
       MsgBox "PLEASE ENTER ONLY Numbers"
       KeyAscii = 0
     End If
End Sub
