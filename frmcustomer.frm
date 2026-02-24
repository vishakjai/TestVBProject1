VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form2 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form2"
   ClientHeight    =   10680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13755
   LinkTopic       =   "Form2"
   ScaleHeight     =   10680
   ScaleWidth      =   13755
   StartUpPosition =   3  'Windows Default
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   8280
      TabIndex        =   36
      Top             =   4200
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      Format          =   69533697
      CurrentDate     =   39793
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      Left            =   2760
      TabIndex        =   35
      Top             =   2400
      Width           =   2175
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   2760
      TabIndex        =   34
      Top             =   5640
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   3120
      TabIndex        =   28
      Top             =   9240
      Width           =   9495
      Begin VB.CommandButton Command6 
         Caption         =   "QUIT"
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
         Left            =   7680
         Picture         =   "frmcustomer.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         Caption         =   "DELETE"
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
         Left            =   5760
         Picture         =   "frmcustomer.frx":6B0A
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton Command4 
         Caption         =   "EDIT"
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
         Left            =   3840
         Picture         =   "frmcustomer.frx":D614
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton save 
         Caption         =   "SAVE"
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
         Left            =   2040
         Picture         =   "frmcustomer.frx":1411E
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton addnew 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ADD NEW"
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
         Left            =   360
         Picture         =   "frmcustomer.frx":1AC28
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.TextBox Text12 
      Height          =   495
      Left            =   7200
      TabIndex        =   27
      Top             =   8280
      Width           =   3255
   End
   Begin VB.TextBox Text11 
      Height          =   285
      Left            =   3960
      TabIndex        =   26
      Top             =   8160
      Width           =   2055
   End
   Begin VB.TextBox Text10 
      Height          =   285
      Left            =   480
      TabIndex        =   25
      Top             =   8160
      Width           =   2415
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Left            =   8520
      TabIndex        =   21
      Top             =   6840
      Width           =   2415
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   5760
      TabIndex        =   20
      Top             =   6960
      Width           =   2295
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   2640
      TabIndex        =   19
      Top             =   6960
      Width           =   2175
   End
   Begin VB.TextBox Text6 
      Height          =   495
      Left            =   8520
      TabIndex        =   15
      Top             =   5640
      Width           =   2535
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   5640
      TabIndex        =   14
      Top             =   5520
      Width           =   2415
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   5520
      TabIndex        =   10
      Top             =   4200
      Width           =   2295
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmcustomer.frx":21732
      Left            =   2640
      List            =   "frmcustomer.frx":2174E
      TabIndex        =   9
      Top             =   4320
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   8160
      TabIndex        =   5
      Top             =   2400
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   5400
      TabIndex        =   4
      Top             =   2400
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2880
      TabIndex        =   3
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "E-MAIL ADDRESS"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7200
      TabIndex        =   24
      Top             =   7920
      Width           =   3135
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "MOBILE NO"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      TabIndex        =   23
      Top             =   7920
      Width           =   2415
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "PHONE NO"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   22
      Top             =   7920
      Width           =   2895
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "LOCATION/TOWN"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8640
      TabIndex        =   18
      Top             =   6360
      Width           =   2295
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "POSTAL CODE"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   17
      Top             =   6360
      Width           =   2415
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "ADDRESS"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   16
      Top             =   6360
      Width           =   2655
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "OPENING BALANCE"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8520
      TabIndex        =   13
      Top             =   5040
      Width           =   2415
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "ACCOUNT NO"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   12
      Top             =   5040
      Width           =   2295
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "ACCOUNT TYPE"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   11
      Top             =   5040
      Width           =   2415
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "DATE JOINED"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8280
      TabIndex        =   8
      Top             =   3600
      Width           =   2535
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "NATIONAL ID NO"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   7
      Top             =   3600
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "CONTACT TITLE"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   6
      Top             =   3600
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "LAST NAME"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8280
      TabIndex        =   2
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "FIRST NAME"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5400
      TabIndex        =   1
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "CUSTOMER ID"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   0
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Image Image2 
      Height          =   6180
      Left            =   0
      Picture         =   "frmcustomer.frx":2177A
      Top             =   1440
      Width           =   2400
   End
   Begin VB.Image Image1 
      Height          =   1500
      Left            =   0
      Picture         =   "frmcustomer.frx":2DBF5
      Top             =   0
      Width           =   12000
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub frm1()
frm
If RS.State = 1 Then RS.Close
RS.Open "select  AccountName from accounttype", CON
Do While Not RS.EOF
Combo3.AddItem ((RS.Fields(0).Value))
RS.MoveNext
Loop
End Sub
Sub clr()
Combo4.Clear

Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Combo1.Text = ""
Text4.Text = ""

Combo3.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""


End Sub
Sub fill()

 frm
If RS.State = 1 Then RS.Close
RS.Open "select  customerid from customer", CON
Do While Not RS.EOF
Combo4.AddItem ((RS.Fields(0).Value))
RS.MoveNext
Loop

   
End Sub


Private Sub addnew_Click()
frm
Combo4.Visible = False


  If RS.State = 1 Then RS.Close
   RS.Open "select max(customerid) from  customer", CON
  If RS.EOF = False Then
  Text1.Text = RS.Fields(0) + 1

  End If
 
 Combo4.Clear
 

'Text2.Text = ""
Text3.Text = ""
Combo1.Text = ""
Text4.Text = ""
Text5.Text = ""
Combo3.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text11.Text = ""
Text12.Text = ""
Text10.Text = ""
Text2.Text = ""
End Sub

Private Sub Combo3_Change()
'If RS.State = 1 Then RS.Close
'rs.Open("select
End Sub

Private Sub Combo4_Click()

If RS.State = 1 Then RS.Close
RS.Open "select * from customer where customerid= " & Combo4.Text & "", CON
If RS.EOF = False Then
Text1.Text = RS.Fields(0).Value
Text2.Text = RS.Fields(1).Value
Text3.Text = RS.Fields(2).Value
Combo1.Text = RS.Fields(3).Value
Text4.Text = RS.Fields(4).Value
DTPicker1.Value = RS.Fields(5).Value
Combo3.Text = RS.Fields(6).Value
Text5.Text = RS.Fields(7).Value
Text6.Text = RS.Fields(8).Value
Text7.Text = RS.Fields(9).Value
Text8.Text = RS.Fields(10).Value
Text9.Text = RS.Fields(11).Value
Text10.Text = RS.Fields(12).Value
Text11.Text = RS.Fields(13).Value
Text12.Text = RS.Fields(14).Value
End If
frm



End Sub

Private Sub Command1_Click()

End Sub

Private Sub Command4_Click()
'fill
If Combo1.Text = "" Then
MsgBox "plz select Customer id"
Else
Dim str As String
   str = MsgBox(("Do you want to Save or Modify This Record?"), vbYesNo, "Customer")
If str = vbYes Then
 CON.Execute ("update customer set customerid=" & Combo4.Text & ",firstname='" & Text2.Text & "',lastname='" & Text3.Text & "',contacttitle='" & Combo1.Text & "',nationalidno=" & Text4.Text & ",datejoined='" & DTPicker1.Value & "',accounttype= '" & Combo3.Text & "',accountno= " & Text5.Text & ",openingbalance=" & Text6.Text & ",address='" & Text7.Text & "',postalcode=" & Text8.Text & ",location='" & Text9.Text & "',phoneno=" & Text10.Text & ",mobileno=" & Text11.Text & ",email='" & Text12.Text & "'where customerid=" & Combo4.Text & " ")
MsgBox "record updated"
End If
End If
clr
fill
End Sub

Private Sub Command5_Click()
Dim str1
If Combo1.Text = "" Then
MsgBox "plz select Customer id"
Else

 str1 = MsgBox(("Do You To Delete This Record Permanantly???"), vbYesNo, "Customer")
            If str1 = vbYes Then

CON.Execute ("delete * from customer where customerid=" & Combo4.Text & "")

MsgBox (" record deleted")
End If
End If
clr

fill

End Sub

Private Sub Command6_Click()
main.Show
Me.Hide
 
End Sub

Private Sub Form_Load()



frm1
'if rs.EOF
'save.Enabled = False
'Command3.Enabled = False
fill

End Sub

Private Sub save_Click()

If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Combo1.Text = "" Or Text4.Text = "" Or Combo3.Text = "" Or Text5.Text = "" Or Text6.Text = "" Or Text7.Text = "" Or Text8.Text = "" Or Text9.Text = "" Or Text10.Text = "" Or Text11.Text = "" Or Text12.Text = "" Then
   MsgBox "please enter all details"
Else
  
 If RS.State = 1 Then RS.Close
 RS.Open ("select MinBalance from accounttype where AccountName='" & Combo3.Text & "'")
 If RS.EOF = False Then
 If Val(Text6.Text) >= RS.Fields(0) Then
  
  If CON.State = 1 Then CON.Close
  CON.Open
  CON.Execute ("insert into customer values(" & Text1.Text & ",'" & Text2.Text & "','" & Text3.Text & "','" & Combo1.Text & "'," & Text4.Text & ",'" & DTPicker1.Value & "','" & Combo3.Text & "'," & Text5.Text & "," & Text6.Text & ",'" & Text7.Text & "'," & Text8.Text & ",'" & Text9.Text & "'," & Text10.Text & "," & Text11.Text & ",'" & Text12.Text & "')")
   MsgBox ("THANK YOU " & UCase(Text2.Text) & ". SAVE RECORD")
   CON.Execute ("insert into balancedt values(" & Text5.Text & "," & Text6.Text & ")")
   
   
   
 Combo4.Visible = True
  clr
 fill
 Else
 MsgBox "Error"
  Exit Sub
 End If
 End If
 End If
End Sub

Private Sub Text1_Change()
Dim dig$, i, digi$, digits$
If Text1.Text <> "" Then
    dig$ = Mid(Text1.Text, Len(Text1.Text), 1)
    If Asc(dig$) < 46 Or Asc(dig$) > 57 Then
        For i = 1 To Len(Text1.Text) - 1
            digi$ = Mid(Text1.Text, i, 1)
            digits$ = digits$ & digi$
        Next i
        Text1.Text = digits$
        Text1.SelStart = Len(Text1.Text)
    End If
End If
End Sub



Private Sub Text10_KeyPress(KeyAscii As Integer)
If Not (KeyAscii < 48) And (KeyAscii > 57) Then
MsgBox "PLEASE ENTER ONLY NUMBERS"
KeyAscii = 0
End If
End Sub



Private Sub Text11_KeyPress(KeyAscii As Integer)
If Not (KeyAscii < 48) And (KeyAscii > 57) Then
MsgBox "PLEASE ENTER ONLY NUMBERS"
KeyAscii = 0
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If Not (KeyAscii < 48) And Not (KeyAscii > 57) Then
MsgBox "PLEASE ENTER ALL ONLY ALPHABETIC CHARACTERS"
KeyAscii = 0
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If Not (KeyAscii < 48) And Not (KeyAscii > 57) Then
MsgBox "PLEASE ENTER ALL ONLY ALPHABETIC CHARACTERS"
KeyAscii = 0
End If
End Sub

Private Sub Text4_Change()

'Dim dig$, i, digi$, digits$
'If Text4.Text <> "" Then
   ' dig$ = Mid(Text4.Text, Len(Text4.Text), 1)
    'If Asc(dig$) < 46 Or Asc(dig$) > 57 Then
       ' For i = 1 To Len(Text4.Text) - 1
          '  digi$ = Mid(Text4.Text, i, 1)
            'digits$ = digits$ & digi$
        'Next i
       ' Text4.Text = digits$
        'Text4.SelStart = Len(Text4.Text)
    'End If
'End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If Not (KeyAscii < 48) And (KeyAscii > 57) Then
MsgBox "PLEASE ENTER ONLY Numbers"
KeyAscii = 0
End If
End Sub



Private Sub Text5_KeyPress(KeyAscii As Integer)
If Not (KeyAscii < 48) And (KeyAscii > 57) Then
MsgBox "PLEASE ENTER ONLY NUMBERS"
KeyAscii = 0
End If

End Sub



Private Sub Text6_KeyPress(KeyAscii As Integer)
If Not (KeyAscii < 48) And (KeyAscii > 57) Then
MsgBox "PLEASE ENTER ONLY NUMBERS"
KeyAscii = 0
End If
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If Not (KeyAscii < 48) And Not (KeyAscii > 57) Then
MsgBox "PLEASE ENTER ALL ONLY ALPHABETIC CHARACTERS"
KeyAscii = 0
End If
End Sub



Private Sub Text8_KeyPress(KeyAscii As Integer)
If Not (KeyAscii < 48) And (KeyAscii > 57) Then
MsgBox "PLEASE ENTER ONLY NUMBERS"
KeyAscii = 0
End If
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
If Not (KeyAscii < 48) And Not (KeyAscii > 57) Then
MsgBox "PLEASE ENTER ALL ONLY ALPHABETIC CHARACTERS"
KeyAscii = 0
End If
End Sub
