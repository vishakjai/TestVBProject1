VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "Form5"
   ClientHeight    =   8640
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9510
   LinkTopic       =   "Form5"
   ScaleHeight     =   8640
   ScaleWidth      =   9510
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   3720
      TabIndex        =   5
      Top             =   4800
      Width           =   5535
      Begin VB.CommandButton cmdQuit 
         Caption         =   "&Quit"
         DownPicture     =   "accounttype.frx":0000
         Height          =   375
         Left            =   4320
         Picture         =   "accounttype.frx":1E53
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         DownPicture     =   "accounttype.frx":895D
         Height          =   375
         Left            =   3240
         Picture         =   "accounttype.frx":A7B0
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         DownPicture     =   "accounttype.frx":112BA
         Height          =   375
         Left            =   2160
         Picture         =   "accounttype.frx":1310D
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         DownPicture     =   "accounttype.frx":19C17
         Height          =   375
         Left            =   1200
         Picture         =   "accounttype.frx":1BA6A
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add New"
         DownPicture     =   "accounttype.frx":22574
         Height          =   375
         Left            =   240
         Picture         =   "accounttype.frx":243C7
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmdFirst 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   240
         Picture         =   "accounttype.frx":2AED1
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   600
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   600
         Picture         =   "accounttype.frx":2B213
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
         Picture         =   "accounttype.frx":2B555
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   600
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdLast 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   4920
         Picture         =   "accounttype.frx":2B897
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   600
         UseMaskColor    =   -1  'True
         Width           =   465
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   960
         TabIndex        =   15
         Top             =   600
         Width           =   3600
      End
   End
   Begin VB.TextBox txtMinBalance 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5070
      TabIndex        =   4
      Top             =   4410
      Width           =   1935
   End
   Begin VB.TextBox txtInterestRate 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5070
      TabIndex        =   3
      Top             =   4035
      Width           =   1455
   End
   Begin VB.TextBox txtDescription 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5070
      TabIndex        =   2
      Top             =   3660
      Width           =   3375
   End
   Begin VB.TextBox txtAccountName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5160
      TabIndex        =   1
      Top             =   3270
      Width           =   3375
   End
   Begin VB.TextBox txtAccountID 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5280
      TabIndex        =   0
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Image Image2 
      Height          =   6180
      Left            =   0
      Picture         =   "accounttype.frx":2BBD9
      Top             =   1560
      Width           =   2400
   End
   Begin VB.Image Image1 
      Height          =   1500
      Left            =   0
      Picture         =   "accounttype.frx":3832E
      Top             =   0
      Width           =   12000
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      TabIndex        =   20
      Top             =   4455
      Width           =   1065
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      TabIndex        =   19
      Top             =   4080
      Width           =   1125
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      TabIndex        =   18
      Top             =   3705
      Width           =   1035
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      TabIndex        =   17
      Top             =   3315
      Width           =   1260
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      TabIndex        =   16
      Top             =   2940
      Width           =   975
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
    frm
    If RS.State = 1 Then RS.Close
    RS.Open "select max(accountid) from accounttype", CON
    If RS.EOF = False Then
      txtAccountID.Text = RS.Fields(0) + 1
    End If
    txtAccountName.Text = ""
    txtDescription.Text = ""
    txtInterestRate.Text = ""
    txtMinBalance.Text = ""
    End Sub

Private Sub cmdEdit_Click()
    CON.Execute "update accounttype set accountid=" & txtAccountID.Text & ",accountname='" & txtAccountName.Text & "',description='" & txtDescription.Text & "',interestRate='" & txtInterestRate.Text & "',minbalance='" & txtMinBalance.Text & "' where accountid=" & txtAccountID.Text & ""
    MsgBox "RECORD EDITED"
End Sub

Private Sub cmdFirst_Click()
    On Error GoTo f
    RS.MoveFirst
    If Not RS.EOF = True Then
        txtAccountID.Text = RS.Fields(0).Value
        txtAccountName.Text = RS.Fields(1).Value
        txtDescription.Text = RS.Fields(2).Value
        txtInterestRate.Text = RS.Fields(3).Value
        txtMinBalance.Text = RS.Fields(4).Value
    End If
    Exit Sub
f:  MsgBox "THIS IS FIRST RECORD"
End Sub


Private Sub cmdLast_Click()
     On Error GoTo l
     RS.MoveLast
    If Not RS.EOF = True Then
        txtAccountID.Text = RS.Fields(0).Value
        txtAccountName.Text = RS.Fields(1).Value
        txtDescription.Text = RS.Fields(2).Value
        txtInterestRate.Text = RS.Fields(3).Value
        txtMinBalance.Text = RS.Fields(4).Value
    End If
    Exit Sub
l:  MsgBox "THIS IS LAST RECORD"
End Sub

Private Sub cmdNext_Click()
     On Error GoTo n
     RS.MoveNext
    If Not RS.EOF = True Then
        txtAccountID.Text = RS.Fields(0).Value
        txtAccountName.Text = RS.Fields(1).Value
        txtDescription.Text = RS.Fields(2).Value
        txtInterestRate.Text = RS.Fields(3).Value
        txtMinBalance.Text = RS.Fields(4).Value
    End If
    Exit Sub
n:  MsgBox "THIS IS LAST RECORD"
End Sub

Private Sub cmdPrevious_Click()
    On Error GoTo p
    RS.MovePrevious
    If Not RS.EOF = True Then
        txtAccountID.Text = RS.Fields(0).Value
        txtAccountName.Text = RS.Fields(1).Value
        txtDescription.Text = RS.Fields(2).Value
        txtInterestRate.Text = RS.Fields(3).Value
        txtMinBalance.Text = RS.Fields(4).Value
    End If
    Exit Sub
p:  MsgBox "THIS IS FIRST RECORD"
End Sub

Private Sub cmdSave_Click()
    frm
    If txtAccountID.Text = "" Or txtAccountName.Text = "" Or txtDescription.Text = "" Or txtInterestRate.Text = "" Or txtMinBalance.Text = "" Then
      MsgBox "plz enter all details"
    Else
      CON.Execute ("insert into accounttype values(" & txtAccountID.Text & ",'" & txtAccountName.Text & "','" & txtDescription.Text & "','" & txtInterestRate.Text & "','" & txtMinBalance.Text & "')")
      MsgBox "RECORD SAVED"
      txtAccountID.Text = ""
      txtAccountName.Text = ""
      txtDescription.Text = ""
      txtInterestRate.Text = ""
      txtMinBalance.Text = ""
    End If
    
End Sub

Private Sub Form_Load()
    frm
    If RS.State = 1 Then RS.Close
    RS.Open "select * from accounttype", CON, adOpenDynamic, adLockOptimistic
    If Not RS.EOF = True Then
      txtAccountID.Text = RS.Fields(0).Value
      txtAccountName.Text = RS.Fields(1).Value
      txtDescription.Text = RS.Fields(2).Value
      txtInterestRate.Text = RS.Fields(3).Value
      txtMinBalance.Text = RS.Fields(4).Value
    End If
    
End Sub

