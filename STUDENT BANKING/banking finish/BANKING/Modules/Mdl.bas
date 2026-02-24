Attribute VB_Name = "Mdl"
'Coding by : William Mwangi
'
'A simple Banking application ideal for beginerswho want to know how to use list
'view control and MaskEDbox
'
'If you want assistance with this code emeail me at kremlin_sniffer@yahoo.com
'Last modified 20th June 2005

Public cnBank As Connection
Public rsCustomers As Recordset
Public rsDeposit As Recordset
Public rsWithdrawal As Recordset
Public rsTransactions As Recordset
Public rsAccTypes As Recordset
Public rsBalances As Recordset
Public rsTemp As Recordset

Public X As Integer
Public NewRecord As Boolean

Public Sub connectDatabase()

Set cnBank = New ADODB.Connection
With cnBank
    .Provider = "Microsoft.JET.OLEDB.4.0"
    .ConnectionString = App.Path & "\dbBank.mdb"
    .Open
End With

Set rsCustomers = New ADODB.Recordset
rsCustomers.Open "tblCustomers", cnBank, adOpenKeyset, adLockOptimistic

Set rsDeposit = New ADODB.Recordset
rsDeposit.Open "tblDeposits", cnBank, adOpenKeyset, adLockOptimistic

Set rsWithdrawal = New ADODB.Recordset
rsWithdrawal.Open "tblWithdrawals", cnBank, adOpenKeyset, adLockOptimistic

Set rsBalances = New ADODB.Recordset
rsBalances.Open "tblBalances", cnBank, adOpenKeyset, adLockOptimistic

Set rsTransactions = New ADODB.Recordset
rsTransactions.Open "tblTransactions", cnBank, adOpenKeyset, adLockOptimistic

Set rsAccTypes = New ADODB.Recordset
rsAccTypes.Open "tblAccTypes", cnBank, adOpenKeyset, adLockOptimistic

End Sub
Public Sub clear_Form_Controls(frm As Form)
Dim ctrl As Control
    For Each ctrl In frm.Controls
        If TypeOf ctrl Is TextBox Then
            ctrl.Text = ""
        ElseIf TypeOf ctrl Is ComboBox Then
            ctrl.Text = ""
'        ElseIf TypeOf ctrl Is MaskEdBox Then
'            ctrl.Text = ""
        Else
'        frm.Refresh
        End If
    Next ctrl
End Sub

Public Sub Lock_Form_Controls(frm As Form)
Dim ctrl As Control
    For Each ctrl In frm.Controls
        If TypeOf ctrl Is TextBox Then
            ctrl.Locked = True
        ElseIf TypeOf ctrl Is ComboBox Then
            ctrl.Locked = True
        ElseIf TypeOf ctrl Is MaskEdBox Then
            ctrl.Enabled = True
        End If
    Next ctrl
End Sub
Public Sub UnLock_Form_Controls(frm As Form)
Dim ctrl As Control
    For Each ctrl In frm.Controls
        If TypeOf ctrl Is TextBox Then
            ctrl.Locked = False
        ElseIf TypeOf ctrl Is ComboBox Then
            ctrl.Locked = False
        ElseIf TypeOf ctrl Is MaskEdBox Then
            ctrl.Enabled = True
        End If
    Next ctrl
End Sub

Public Sub disconnectDatabase()
cnBank.Close
End Sub
Public Sub selectTextControl(txtCtrl As TextBox)
    txtCtrl.SelStart = 0
    txtCtrl.SelLength = Len(txtCtrl.Text)
End Sub

Public Sub selectMaskControl(maskCtrl As MaskEdBox)
    maskCtrl.SelStart = 0
    maskCtrl.SelLength = Len(maskCtrl.Text)
End Sub
Public Sub ValidNonNumeric(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
Select Case KeyAscii
 Case Asc(" ")
 Case 65 To 90
 Case 97 To 122
 Case 32
 Case 13
 Case 8
 Case 127
 Case Else
  MsgBox "Invalid Input. Please Don't Enter Numerics...", vbOKOnly + vbExclamation
  KeyAscii = 0
End Select
End Sub

Public Sub ValidNumeric(KeyAscii As Integer)
Select Case KeyAscii
Case 8
Case 97
Case 110
Case 47
Case 13
Case 32
Case 48 To 57
 Case Else
  MsgBox "Invalid Input.Please Enter Numeric Types Only..", vbOKOnly + vbExclamation
  KeyAscii = 0
End Select
End Sub

Public Sub CheckDatabaseStatus(rsStat As Recordset)
With rsStat
If .BOF = True And .EOF = True Then
MsgBox "There are currently No records Available for this module", vbInformation
Exit Sub
End If
End With
End Sub
Public Sub MoveToFirst(rsFirst As Recordset)
With rsFirst
Call CheckDatabaseStatus(rsFirst)
.MoveFirst
If .BOF Then
.MoveFirst
MsgBox "This is the first Record..", vbInformation
Exit Sub
End If
End With
End Sub
Public Sub MoveToPrev(rsPrev As Recordset)
With rsPrev
Call CheckDatabaseStatus(rsPrev)
.MovePrevious
If .BOF Then
.MoveFirst
MsgBox "This is the first Record..", vbInformation
Exit Sub
End If

End With

End Sub
Public Sub MoveToNext(rsNext As Recordset)
With rsNext
Call CheckDatabaseStatus(rsNext)
.MoveNext
If .EOF Then
.MoveLast
MsgBox "This is the last Record..", vbInformation
Exit Sub
End If
End With
End Sub

Public Sub MoveToLast(rsLast As Recordset)
With rsLast
Call CheckDatabaseStatus(rsLast)
.MoveLast
If .EOF Then
.MoveLast
MsgBox "This is the last Record..", vbInformation
Exit Sub
End If
End With
End Sub

Public Sub Messager()
MsgBox "Please Ensure that all fields are Complete", vbExclamation
End Sub
