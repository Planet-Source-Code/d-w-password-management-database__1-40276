VERSION 5.00
Begin VB.Form Login 
   Appearance      =   0  'Flat
   ClientHeight    =   1755
   ClientLeft      =   1095
   ClientTop       =   1170
   ClientWidth     =   3450
   ControlBox      =   0   'False
   LinkTopic       =   "Login"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   3450
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame 
      Caption         =   "Login"
      Height          =   1500
      Index           =   0
      Left            =   0
      TabIndex        =   17
      Top             =   240
      Visible         =   0   'False
      Width           =   3450
      Begin VB.TextBox EnterID 
         Height          =   336
         Left            =   255
         TabIndex        =   6
         Top             =   435
         Width           =   1692
      End
      Begin VB.CommandButton NoSubmit 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Cancel"
         Height          =   372
         Left            =   2220
         TabIndex        =   9
         Top             =   885
         Width           =   972
      End
      Begin VB.CommandButton Submit 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "OK"
         Height          =   372
         Left            =   2220
         TabIndex        =   8
         Top             =   285
         Width           =   972
      End
      Begin VB.TextBox EnterPass 
         Height          =   336
         IMEMode         =   3  'DISABLE
         Left            =   255
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   1035
         Width           =   1692
      End
      Begin VB.Label EnterLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Enter your password:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   255
         TabIndex        =   18
         Top             =   825
         Width           =   1755
      End
      Begin VB.Label EnterLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Enter your User ID:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   255
         TabIndex        =   19
         Top             =   225
         Width           =   1575
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "Options"
      Height          =   1500
      Index           =   5
      Left            =   0
      TabIndex        =   10
      Top             =   240
      Visible         =   0   'False
      Width           =   3450
      Begin VB.CommandButton Button 
         Caption         =   "&Close"
         Height          =   300
         Index           =   5
         Left            =   1800
         TabIndex        =   16
         Top             =   1065
         Width           =   1500
      End
      Begin VB.CommandButton Button 
         Caption         =   "&Log In"
         Height          =   300
         Index           =   4
         Left            =   1800
         TabIndex        =   15
         Top             =   660
         Width           =   1500
      End
      Begin VB.CommandButton Button 
         Caption         =   "&Change Password"
         Height          =   300
         Index           =   3
         Left            =   1785
         TabIndex        =   14
         Top             =   255
         Width           =   1500
      End
      Begin VB.CommandButton Button 
         Caption         =   "&Remove User"
         Height          =   300
         Index           =   2
         Left            =   135
         TabIndex        =   13
         Top             =   1065
         Width           =   1500
      End
      Begin VB.CommandButton Button 
         Caption         =   "&Add User"
         Height          =   300
         Index           =   1
         Left            =   150
         TabIndex        =   12
         Top             =   660
         Width           =   1500
      End
      Begin VB.CommandButton Button 
         Caption         =   "&List Users"
         Height          =   300
         Index           =   0
         Left            =   150
         TabIndex        =   11
         Top             =   255
         Width           =   1500
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "Remove User"
      Height          =   1500
      Index           =   2
      Left            =   0
      TabIndex        =   29
      Top             =   225
      Visible         =   0   'False
      Width           =   3450
      Begin VB.ListBox cboDeleteList 
         Height          =   1035
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   32
         Top             =   300
         Width           =   2055
      End
      Begin VB.CommandButton NoDeleteUser 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Cancel"
         Height          =   372
         Left            =   2325
         TabIndex        =   31
         Top             =   975
         Width           =   870
      End
      Begin VB.CommandButton DeleteUser 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Delete"
         Height          =   372
         Left            =   2325
         TabIndex        =   30
         Top             =   375
         Width           =   852
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "Change password"
      Height          =   1500
      Index           =   3
      Left            =   0
      TabIndex        =   33
      Top             =   240
      Visible         =   0   'False
      Width           =   3450
      Begin VB.CommandButton NoChange 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Cancel"
         Height          =   372
         Left            =   2175
         TabIndex        =   37
         Top             =   870
         Width           =   972
      End
      Begin VB.CommandButton Change 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   372
         Left            =   2175
         TabIndex        =   36
         Top             =   315
         Width           =   972
      End
      Begin VB.TextBox NewPswd 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   240
         PasswordChar    =   "*"
         TabIndex        =   35
         Top             =   960
         Width           =   1550
      End
      Begin VB.TextBox OldPwd 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   240
         PasswordChar    =   "*"
         TabIndex        =   34
         Top             =   420
         Width           =   1550
      End
      Begin VB.Label MsgLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   240
         TabIndex        =   40
         Top             =   1245
         Width           =   1455
      End
      Begin VB.Label ChangeLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "New Password: "
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   39
         Top             =   750
         Width           =   1335
      End
      Begin VB.Label ChangeLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Old Password:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   38
         Top             =   210
         Width           =   1215
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "Users"
      Height          =   1500
      Index           =   4
      Left            =   0
      TabIndex        =   25
      Top             =   240
      Visible         =   0   'False
      Width           =   3450
      Begin VB.CommandButton MoreInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "More Info"
         Height          =   372
         Left            =   2460
         TabIndex        =   28
         Top             =   360
         Width           =   852
      End
      Begin VB.CommandButton NoMoreInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Cancel"
         Height          =   372
         Left            =   2460
         TabIndex        =   27
         Top             =   870
         Width           =   852
      End
      Begin VB.ListBox cboUserList 
         Height          =   1035
         Left            =   90
         Sorted          =   -1  'True
         TabIndex        =   26
         Top             =   315
         Width           =   2295
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "Add user"
      Height          =   1500
      Index           =   1
      Left            =   0
      TabIndex        =   20
      Top             =   240
      Visible         =   0   'False
      Width           =   3450
      Begin VB.ComboBox cboTaskLevel 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   675
         TabIndex        =   2
         Top             =   825
         Width           =   1530
      End
      Begin VB.TextBox NewName 
         Height          =   288
         Left            =   720
         TabIndex        =   0
         Top             =   195
         Width           =   2610
      End
      Begin VB.TextBox NewID 
         Height          =   288
         Left            =   450
         TabIndex        =   1
         Top             =   510
         Width           =   1770
      End
      Begin VB.TextBox NewPwd 
         Height          =   288
         IMEMode         =   3  'DISABLE
         Left            =   870
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1170
         Width           =   1260
      End
      Begin VB.CommandButton NoNewUser 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   372
         Left            =   2355
         TabIndex        =   5
         Top             =   885
         Width           =   972
      End
      Begin VB.CommandButton NewUser 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "OK"
         Height          =   372
         Left            =   2355
         TabIndex        =   4
         Top             =   495
         Width           =   972
      End
      Begin VB.Label ConfirmLabel 
         Caption         =   "Confirm Password"
         Height          =   165
         Left            =   2130
         TabIndex        =   42
         Top             =   1245
         Visible         =   0   'False
         Width           =   1290
      End
      Begin VB.Label NewLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Level:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   60
         TabIndex        =   24
         Top             =   885
         Width           =   510
      End
      Begin VB.Label NewLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "ID:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   23
         Top             =   570
         Width           =   240
      End
      Begin VB.Label NewLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Name:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   60
         TabIndex        =   22
         Top             =   240
         Width           =   555
      End
      Begin VB.Label NewLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         Caption         =   "Password:"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   45
         TabIndex        =   21
         Top             =   1230
         Width           =   870
      End
   End
   Begin VB.Label TitleBar 
      BackColor       =   &H80000002&
      Caption         =   "Password Management"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   270
      Left            =   0
      TabIndex        =   41
      Top             =   0
      Width           =   3450
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SendMessage Lib "User32" _
    Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
    ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub ReleaseCapture Lib "User32" ()
Private Declare Function SetWindowPos Lib "User32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
Private Const TermLength = 60
Private Const MaxNameLen = 50
Private Const MinIDLen = 3
Private Const MaxIDLen = 25
Private Const MinPassLen = 6
Private Const MaxPassLen = 12
Private Const LoginDatabase = "Password.mdb"
Private Const AdminBackdoor = "*" 'if database is empty tables
Private Const AppTitle = "Password Program" 'encryption key
Private Const AllowedAttempts = 3
Private Db As Database
Private Rc As Recordset
Private Response As VbMsgBoxResult
Private UserID As String
Private Password As String
Private UserName As String
Private AccessLevel As String
Private ActivationDate As String
Private ExpireDate As String
Private Attempts As Integer
Private PasswordExpired As Boolean

Private Sub ClearData()
Dim i As Integer
Dim ctrl As Control
Dim frm As Form
Set frm = Me
For Each ctrl In frm
    If TypeOf ctrl Is TextBox Then
    ctrl.Text = ""
    ElseIf TypeOf ctrl Is ComboBox Then
    ctrl.Clear
    End If
Next
For i = 1 To 5
Frame(i).Visible = False
Next
End Sub

Private Sub DecriptDatabase()
On Local Error Resume Next

Dim UserID As String
Dim UserName As String
Dim Temp As String

Rc.MoveFirst

Do
UserID = Rc("UserID")
UserName = Rc("UserName")
Temp = Rc("Password")
Temp = Crypt("D", AppTitle, Temp)
Debug.Print UserName & " " & UserID & " " & Temp
Rc.MoveNext
If Rc.EOF Then Exit Sub
Loop

End Sub

Private Function HasNoNumbers(Words As String) As Boolean
Dim i As Integer
For i = 0 To 9
    If InStr(1, Words, CStr(i)) <> 0 Then
    HasNoNumbers = False
    Exit Function
    End If
Next
HasNoNumbers = True
End Function

Private Sub ReturnToPrevious()
ClearData
If Left(Me.Tag, 1) = "5" Then
ShowOptions
Else
Me.Hide
End If
End Sub

Private Sub ShowChangePassword()
Dim i As Integer
For i = 0 To 5
    If i <> 3 Then
    Frame(i).Visible = False
    Else
    Frame(i).Visible = True
    End If
Next
End Sub

Private Sub ShowLogin()
Dim i As Integer

For i = 0 To 5
    If i <> 0 Then
    Frame(i).Visible = False
    Else
    Frame(i).Visible = True
    End If
Next
EnterID = GetSetting("Password Program", "Users", "Last", UserID)
EnterID.SelStart = 0
EnterID.SelLength = Len(EnterID)

End Sub

Private Sub ShowOptions()
Dim i As Integer
For i = 0 To 5
    If i <> 5 Then
    Frame(i).Visible = False
    Else
    Frame(i).Visible = True
    End If
Next
End Sub

Private Sub ShowSelectedOption()
Select Case Left(Me.Tag, 1)
Case "0"
ShowLogin
Case "1"
If Val(AccessLevel) = 5 Then
ShowNewUser
Else
MsgBox "Access Denied", vbCritical, "Insufficient Access Level"
ReturnToPrevious
End If
Case "2"
If Val(AccessLevel) = 5 Then
ShowDeleteUser
Else
MsgBox "Access Denied", vbCritical, "Insufficient Access Level"
ReturnToPrevious
End If
Case "3"
ShowChangePassword
Case "4"
ShowUsers
Case "5"
ShowOptions
Case Else
ShowLogin
End Select
End Sub

Private Sub Button_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 27 Then
Button_Click (5)
End If
End Sub

Private Sub cboDeleteList_DblClick()
DeleteUser_Click
End Sub


Private Sub ConfirmLabel_Click()
'label
End Sub

Private Sub Form_Activate()
Initialize
ShowSelectedOption
'the sample password database contains 5 users
'the passwords are the same as the name(min 6 characters)
'and must have letters and numbers
'guest1
'staff1
'super1
'manager1
'admin1
'the "AdminBackdoor" to creating a dtabase from empty is
'username = "*"
'password = "*"
'consider removing this or, if known, access could be
'gained by deleting the database which is recreated
'empty if not found on startup.
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'this should not really be necessary since the scope
'is private but to avoid damage to database it's a
'good idea
Rc.Close
Set Rc = Nothing
Db.Close
Set Db = Nothing
End Sub


Private Sub Frame_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
'controls are in sets on frame containers
'to edit the user interface, right click on a
'blank spot in a frame and choose send to back
'and the successive control sets will come to front
End Sub

Private Sub NewPwd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
NewUser_Click
KeyAscii = 0
End If
End Sub
Private Sub ShowDeleteUser()
On Local Error Resume Next
Dim i As Integer
For i = 0 To 5
    If i <> 2 Then
    Frame(i).Visible = False
    Else
    Frame(i).Visible = True
    End If
Next
cboDeleteList.Clear
Rc.MoveFirst
Do Until Rc.EOF
cboDeleteList.AddItem Rc("UserID") & ", " & Rc("UserName")  'field
Rc.MoveNext
Loop
End Sub

Private Sub NoChange_Click()
If PasswordExpired Then End
Attempts = 0
ReturnToPrevious
End Sub

Private Sub NoDeleteUser_Click()
cboDeleteList.Clear
ReturnToPrevious
End Sub

Private Sub NoMoreInfo_Click()
cboUserList.Clear
ReturnToPrevious
End Sub

Private Sub NoNewUser_Click()
cboTaskLevel.Clear
ReturnToPrevious
End Sub


Private Sub NoSubmit_Click()
If Len(Me.Tag) < 2 Then
End
Else
ReturnToPrevious
End If
End Sub
Private Sub NewPswd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Change_Click
KeyAscii = 0
End If
End Sub
Private Sub OldPwd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{Tab}"
KeyAscii = 0
End If
End Sub

Private Sub NewID_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then 'enter key
SendKeys "{Tab}"
KeyAscii = 0
End If
End Sub
Private Sub NewName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then 'enter key
SendKeys "{Tab}"
KeyAscii = 0
End If
End Sub

Private Sub DeleteUser_Click()

Dim RemoveID As String
Dim Mark As Integer

Response = MsgBox("Delete User?" & cboDeleteList, 20, "Delete User")
If Response = 6 Then
RemoveID = cboDeleteList
Mark = InStr(1, RemoveID, ",")
RemoveID = Left(RemoveID, Mark - 1)
DeletePassword RemoveID
cboDeleteList.RemoveItem cboDeleteList.ListIndex
cboDeleteList.SetFocus
Exit Sub
End If

End Sub

Private Sub ShowNewUser()

Dim i As Integer
For i = 0 To 5
    If i <> 1 Then
    Frame(i).Visible = False
    Else
    Frame(i).Visible = True
    End If
Next
cboTaskLevel.Clear
cboTaskLevel.AddItem "1 - Guest"
cboTaskLevel.AddItem "2 - Staff"
cboTaskLevel.AddItem "3 - Supervisor"
cboTaskLevel.AddItem "4 - Management"
cboTaskLevel.AddItem "5 - Administrator"
cboTaskLevel.Text = "2 - Staff"
Attempts = 0
End Sub
Private Sub NewUser_Click()

Static Entry1 As String
Static Entry2 As String
Dim TempID As String
Dim AddNew As Boolean
Dim TempName As String
Dim Msg As String
Dim i As Integer
AddNew = True

If NewName = "" Then
NewPwd.SetFocus
MsgBox "User name required."
Exit Sub

ElseIf NewID = "" Then
NewID.SetFocus
MsgBox "User ID required."
Exit Sub

ElseIf NewPwd = "" Then
NewPwd.SetFocus
MsgBox "Password is required."
Exit Sub
End If

If Len(NewName) > MaxNameLen Then
Msg = "User name too long - maximum is "
Msg = Msg + Str(MaxNameLen) + " characters."
MsgBox Msg, vbExclamation, "User Name"
NewName = ""
NewName.SetFocus
Exit Sub
End If

If Len(NewID) < MinIDLen Then
Msg = "User ID too short - minimum is "
Msg = Msg + Str(MinIDLen) + " characters."
MsgBox Msg, vbExclamation, "User ID"
NewID = ""
NewID.SetFocus
Exit Sub

ElseIf Len(NewID) > MaxIDLen Then
Msg = "User ID too long - maximum is "
Msg = Msg + Str(MaxIDLen) + " characters."
MsgBox Msg, vbExclamation, "User ID"
NewID = ""
NewID.SetFocus
Exit Sub
End If

If Len(NewPwd) < MinPassLen Then
Msg = "Password too short - minimum is "
Msg = Msg + Str(MinPassLen) + " characters."
MsgBox Msg, vbExclamation, "Password"
NewPwd = ""
NewPwd.SetFocus
Exit Sub

ElseIf Len(NewPwd) > MaxPassLen Then
Msg = "Password too long - maximum is "
Msg = Msg + Str(MaxPassLen) + " characters."
MsgBox Msg, vbExclamation, "Password"
NewPwd = ""
NewPwd.SetFocus
Exit Sub

ElseIf IsNumeric(NewPwd) Or HasNoNumbers(NewPwd) Then
Msg = "Password must contain both letters and numbers."
MsgBox Msg, vbExclamation, "Password"
NewPwd = ""
NewPwd.SetFocus
Exit Sub
End If

Attempts = Attempts + 1
If Attempts = 1 Then
Entry1 = UCase(NewPwd)
NewPwd = ""
ConfirmLabel.Visible = True
NewPwd.SetFocus
Exit Sub
Else
Entry2 = UCase(NewPwd)
    If Entry1 <> Entry2 Then
    MsgBox "Entries do not match", vbInformation, "Password"
    NewPwd.SelStart = 0
    NewPwd.SelLength = Len(NewPwd)
    Attempts = 0
    ConfirmLabel.Visible = True
    Exit Sub
    End If
End If
Rc.FindFirst "[UserID] = '" & NewID & "'"
    If Not Rc.NoMatch Then
    Response = MsgBox("User Exists. Is this an Edit?", vbCritical + vbYesNo, "Edit existing user?")
        If Response = vbNo Then
        NewName = ""
        NewID = ""
        NewPwd = ""
        NewName.SetFocus
        Exit Sub
        End If
   AddNew = False
   End If

Entry1 = Entry1 & "|" & Left(cboTaskLevel, 1) & "|" & Now & "|" & DateAdd("d", TermLength, Now)
TempID = UCase(NewID)
TempName = UCase(NewName)
WritePassword AddNew, TempID, TempName, Entry1
ConfirmLabel.Visible = False
Attempts = 0
cboTaskLevel.Clear
ReturnToPrevious
End Sub

Public Sub CreateNewDB()

Dim Tb As TableDef
Dim DbName As String

DbName = App.Path & "\" & LoginDatabase
Set Db = CreateDatabase(DbName, dbLangGeneral)
Set Tb = Db.CreateTableDef("Passwords")
With Tb
    .Fields.Append .CreateField("UserID", dbText, 25)
    .Fields.Append .CreateField("UserName", dbText, 55)
    .Fields.Append .CreateField("Password", dbText, 255)
End With
Db.TableDefs.Append Tb
Set Tb = Nothing
End Sub


Private Sub Submit_Click()
On Local Error Resume Next
Dim Info As String
Dim DaysLeft As Integer
Dim Temp As String
Dim Mark As Integer
Dim Spot As Integer
Static Attempts As Integer

If EnterID = "" Then
EnterID.SetFocus
Exit Sub
ElseIf EnterPass = "" Then
EnterPass.SetFocus
Exit Sub
End If

Attempts = Attempts + 1
UserID = UCase(EnterID)
Info = ReadInformation(UserID)
If Password = AdminBackdoor Then
ClearData
ShowOptions
Exit Sub
End If
Mark = InStr(1, Info, "|")
Password = Left(Info, Mark - 1)
'AccessLevel
Spot = InStr(Mark + 1, Info, "|")
AccessLevel = Mid(Info, Mark + 1, 1)
Mark = Spot
'Activation date
Spot = InStr(Mark + 1, Info, "|")
ActivationDate = Mid(Info, Mark + 1, Spot - Mark - 1)
'Expire date
ExpireDate = Right(Info, Len(Info) - Spot)

If UCase(EnterPass) = Password Then
    'password validated
    ClearData
    SaveSetting "Password Program", "Users", "Last", UserID
    If Len(Me.Tag) < 2 Then
    Me.Tag = " " & AccessLevel
    Else
    Me.Tag = Left(Me.Tag, 1) & AccessLevel
    End If
    'check for expiration date
    DaysLeft = DateDiff("d", Now, ExpireDate)
    If DaysLeft < 1 Then
    MsgBox "Your password has expired and must be changed.", vbCritical, "Expired User"
    PasswordExpired = True
    ShowChangePassword
    Exit Sub
    Else
        If Len(Me.Tag) = 2 Then
        EnterID.SetFocus
        ReturnToPrevious
        Else
        ClearData
        Me.Hide
        End If
    End If
Exit Sub
ElseIf Attempts >= AllowedAttempts Then
MsgBox "Access denied.", vbCritical, "Password"
    If Me.Tag = "" Then
    End
    Else
    ClearData
    ReturnToPrevious
    End If
End If
MsgBox "Invalid username or password.", vbExclamation, "Password"
EnterPass = ""
EnterPass.SetFocus
End Sub

Private Sub EnterID_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys "{Tab}"
KeyAscii = 0
End If
End Sub

Private Sub EnterPass_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Submit_Click
KeyAscii = 0
End If
End Sub
Private Sub ShowUsers()

On Local Error Resume Next

Dim i As Integer

For i = 0 To 5
    If i <> 4 Then
    Frame(i).Visible = False
    Else
    Frame(i).Visible = True
    End If
Next
cboUserList.Clear
If Val(AccessLevel) >= 4 Then
Rc.MoveFirst
Do Until Rc.EOF
cboUserList.AddItem Rc("UserID") & ", " & Rc("UserName")
Rc.MoveNext
Loop
Else
cboUserList.AddItem UserID & ", " & UserName
End If

End Sub

Private Sub MoreInfo_Click()

Dim Temp As String
Dim UserTaskLevel As String
Dim UserActivation As String
Dim UserExpire As String
Dim Mark As Integer
Dim Info As String
Dim Spot As Integer

Temp = cboUserList
Mark = InStr(1, Temp, ",")
If Mark > 0 Then
Temp = Left(Temp, Mark - 1) 'assign the ID only
Info = ReadInformation(Temp)
Mark = InStr(1, Info, "|")
'AccessLevel
Spot = InStr(Mark + 1, Info, "|")
UserTaskLevel = Mid(Info, Mark + 1, 1)
Mark = Spot
'activation date
Spot = InStr(Mark + 1, Info, "|")
UserActivation = Mid(Info, Mark + 1, Spot - Mark - 1)
'expire date
UserExpire = Right(Info, Len(Info) - Spot)
MsgBox "ID & Name: " & cboUserList & Chr(13) & "Access Level is " & UserTaskLevel & Chr(13) & "Last change was " & UserActivation & Chr(13) & "Expires on " & UserExpire, vbInformation, "Details"
Else
MsgBox "Select a user."
End If
End Sub

Private Sub TitleBar_Click()
'fake title bar from label
End Sub

Private Sub TitleBar_DblClick()
DecriptDatabase 'to debug window
End Sub


Private Sub TitleBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lngReturnValue As Long
If Button = 1 Then
ReleaseCapture
lngReturnValue = SendMessage(Login.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End If
End Sub


Private Sub cboUserList_DblClick()
MoreInfo_Click
End Sub
Private Sub Change_Click()

Dim Msg As String
Dim Action As Boolean
Static Entry1 As String
Static Entry2 As String

If Trim(OldPwd) = "" Then
OldPwd.SetFocus
MsgLabel.Caption = "Old Password is required."
Exit Sub
ElseIf Trim(NewPswd) = "" Then
NewPswd.SetFocus
MsgLabel.Caption = "New Password is required."
Exit Sub
End If
'is old password correct?
OldPwd = Trim(OldPwd)
NewPswd = Trim(NewPswd)
If UCase(OldPwd) = Password Then
NewPswd.SetFocus
Else
MsgBox "Password incorrect for logged on user.", vbCritical, "Password"
NewPswd = ""
OldPwd = ""
OldPwd.SetFocus
Attempts = 0
MsgLabel.Caption = "New Password is required."
Exit Sub
End If
'does new password pass tests?

If Attempts = 0 Then
    If Len(NewPswd) < MinPassLen Then
    Msg = "Password too short - minimum is "
    Msg = Msg + Str(MinPassLen) + " characters."
    MsgBox Msg, vbExclamation, "Password"
    NewPswd = ""
    NewPswd.SetFocus
    Attempts = 0
    Exit Sub
    
    ElseIf Len(NewPswd) > MaxPassLen Then
    Msg = "Password too long - maximum is "
    Msg = Msg + Str(MaxPassLen) + " characters."
    MsgBox Msg, vbExclamation, "Password"
    NewPswd = ""
    NewPswd.SetFocus
    Attempts = 0
    Exit Sub
    
    ElseIf IsNumeric(NewPswd) Or HasNoNumbers(NewPswd) Then
    Msg = "Password must contain both letters and numbers."
    MsgBox Msg, vbExclamation, "Password"
    NewPswd = ""
    NewPswd.SetFocus
    Attempts = 0
    Exit Sub
    
    ElseIf NewPswd = OldPwd Then
    Msg = "New password must be different from old password."
    MsgBox Msg, vbExclamation, "Password"
    NewPswd = ""
    NewPswd.SetFocus
    Attempts = 0
    Exit Sub
    End If
End If
'old passed
Attempts = Attempts + 1
If Attempts = 1 Then
Entry1 = UCase(NewPswd)
NewPswd = ""
MsgLabel = "Verify New password:"
NewPswd.SetFocus
Exit Sub
Else
Entry2 = UCase(NewPswd)
    If Entry1 <> Entry2 Then
    MsgBox "Sorry - Entries do not match", vbInformation, "Password"
    NewPswd = ""
    NewPswd.SetFocus
    Exit Sub
    End If
End If
'if the two entries match.
Password = Entry1
Entry1 = Entry1 & "|" & AccessLevel & "|" & Now & "|" & DateAdd("d", TermLength, Now)
WritePassword Action, UserID, UserName, Entry1
PasswordExpired = False
MsgLabel = ""
Attempts = 0

If Me.Tag = "" Then
ClearData
Me.Hide
Else
ReturnToPrevious
End If
End Sub


Private Sub Initialize()

Dim Find As String
SetWindowPos Login.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE
Screen.MousePointer = vbHourglass
DoEvents
Find = Dir(App.Path & "\" & LoginDatabase)

If Find = "" Then
MsgBox "No existing password database. Please use preset administrator password to populate it with data."
CreateNewDB
End If
Set Db = OpenDatabase(App.Path & "\" & LoginDatabase)
Set Rc = Db.OpenRecordset("Passwords", dbOpenDynaset)
Screen.MousePointer = 0
End Sub
Private Function ReadInformation(UserID As String) As String

On Local Error GoTo Err1:

Screen.MousePointer = vbHourglass

If Rc.RecordCount = 0 Then
UserID = AdminBackdoor 'reset in general declarations
Password = AdminBackdoor 'for staring with empty database
AccessLevel = "5"
ExpireDate = DateAdd("d", 1, Now)
Screen.MousePointer = vbDefault
Exit Function
End If

Rc.MoveLast
Rc.FindFirst "[UserID] = '" & UserID & "'"
    If Rc.NoMatch Then
    Password = Chr(177)
    Screen.MousePointer = vbDefault
    Exit Function
    End If
UserName = Rc("UserName")
ReadInformation = Crypt("D", AppTitle, Rc("Password"))

Screen.MousePointer = vbDefault
Exit Function
Err1:
MsgBox "Database error, " & Error & ", please remove and re-add this user.", vbCritical
Resume Next
End Function

Private Sub DeletePassword(TempID As String)

On Local Error GoTo Err1:
Screen.MousePointer = vbHourglass

Rc.FindFirst "[UserID] = '" & TempID & "'"

If Rc.NoMatch Then
Screen.MousePointer = vbDefault
Exit Sub
End If

Rc.Delete
MsgBox "User removed.", vbInformation
Screen.MousePointer = vbDefault

Exit Sub
Err1:
MsgBox "Error removing user.", vbCritical
Exit Sub
End Sub

Private Function Crypt(Action As String, Key As String, Src As String) As String
On Error GoTo 0
Dim Count As Integer
Dim KeyPos As Integer
Dim KeyLen As Integer
Dim SrcAsc As Integer
Dim Dest As String
Dim Offset As Integer
Dim TmpSrcAsc As Integer
Dim SrcPos As Integer

KeyLen = Len(Key)

If Action = "E" Then
Randomize
Offset = (Rnd * 10000 Mod 255) + 1
Dest = Hex(Offset)

For SrcPos = 1 To Len(Src)
SrcAsc = (Asc(Mid(Src, SrcPos, 1)) + Offset) Mod 255
    If KeyPos < KeyLen Then
    KeyPos = KeyPos + 1
    Else
    KeyPos = 1
    End If
SrcAsc = SrcAsc Xor Asc(Mid(Key, KeyPos, 1))
Dest = Dest + Format(Hex(SrcAsc), "@@")
Offset = SrcAsc
Next

ElseIf Action = "D" Then
Offset = Val("&H" + Left(Src, 2))
    For SrcPos = 3 To Len(Src) Step 2
    SrcAsc = Val("&H" + Trim(Mid(Src, SrcPos, 2)))
        If KeyPos < KeyLen Then
        KeyPos = KeyPos + 1
        Else
        KeyPos = 1
        End If
    TmpSrcAsc = SrcAsc Xor Asc(Mid(Key, KeyPos, 1))
        If TmpSrcAsc <= Offset Then
        TmpSrcAsc = 255 + TmpSrcAsc - Offset
        Else
        TmpSrcAsc = TmpSrcAsc - Offset
        End If
    Dest = Dest + Chr(TmpSrcAsc)
    Offset = SrcAsc
    Next
End If
Crypt = Dest
End Function


Private Sub WritePassword(AddNew As Boolean, TempID As String, TempName As String, Temp As String)

On Local Error GoTo Err1:

Screen.MousePointer = vbHourglass
Temp = Crypt("E", AppTitle, Temp)

If AddNew Then
Rc.AddNew
Rc("UserID") = TempID
Rc("UserName") = TempName
Rc("Password") = Temp
Rc.Update
Else
Rc.FindFirst "[UserID] = '" & TempID & "'"
Rc.Edit
Rc("UserName") = TempName
Rc("Password") = Temp
Rc.Update
End If
MsgBox "Update successful.", vbInformation
Screen.MousePointer = vbDefault

Exit Sub
Err1:
MsgBox "Error updating password."
Exit Sub
End Sub









Private Sub Button_Click(Index As Integer)
Select Case Index
Case 0
ShowUsers
Case 1
If Val(AccessLevel) = 5 Then
ShowNewUser
Else
MsgBox "Access Denied", vbCritical, "Insufficient Access Level"
End If
Case 2
If Val(AccessLevel) = 5 Then
ShowDeleteUser
Else
MsgBox "Access Denied.", vbCritical, "Insufficient Access Level"
End If
Case 3
ShowChangePassword
Case 4
ShowLogin
Case 5
ClearData
Me.Hide
End Select
End Sub


