VERSION 5.00
Begin VB.Form Test 
   Caption         =   "Form1"
   ClientHeight    =   3180
   ClientLeft      =   165
   ClientTop       =   765
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3180
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuPass 
      Caption         =   "&Password"
      Begin VB.Menu mnuShow 
         Caption         =   "&Show Options"
      End
   End
   Begin VB.Menu mnuOpt 
      Caption         =   "&Options"
      Begin VB.Menu mnuLog 
         Caption         =   "&Login As..."
      End
      Begin VB.Menu mnuChg 
         Caption         =   "&Change Password"
      End
      Begin VB.Menu mnuAdd 
         Caption         =   "&Add New User"
      End
      Begin VB.Menu mnuDel 
         Caption         =   "&Remove User"
      End
      Begin VB.Menu mnuList 
         Caption         =   "&User Information"
      End
   End
   Begin VB.Menu mnuAccess 
      Caption         =   "&Access Levels"
      Begin VB.Menu mnuCurrent 
         Caption         =   "&Current Level"
      End
   End
End
Attribute VB_Name = "Test"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Function AppAccessLevel() As Integer
'this function is not needed for password validation
'it is included so that you can restrict access to
'selected parts of your program.
If Len(Login.Tag) = 2 Then
AppAccessLevel = CInt(Right(Login.Tag, 1))
Else
AppAccessLevel = 0
End If
End Function


Private Sub Form_Load()
Login.Show vbModal
Debug.Print "AccessLevel = " & CStr(AppAccessLevel)
'This is just a test form to show how login form
'is used. You only need to add Login.frm to your
'project and show it modally (important) to add
'password management to a program. The AppAccessLevel
'function in the test form is to give levels of
'access to the main program and is not needed for
'levels of access to the password management features.
'the sample password database contains 5 users
'the passwords are the 'SAME AS THE NAME' (min 6 characters)
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


Private Sub Form_Unload(Cancel As Integer)
Unload Login
End Sub


Private Sub mnuAdd_Click()
Login.Tag = "1" & AppAccessLevel
Login.Show vbModal
End Sub

Private Sub mnuChg_Click()
Login.Tag = "3" & AppAccessLevel
Login.Show vbModal
End Sub


Private Sub mnuCurrent_Click()
MsgBox "Application Access Level is " & AppAccessLevel & " on a scale of 0 to 5."
End Sub

Private Sub mnuDel_Click()
Login.Tag = "2" & AppAccessLevel
Login.Show vbModal
End Sub


Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnuList_Click()
Login.Tag = "4" & AppAccessLevel
Login.Show vbModal
End Sub


Private Sub mnuLog_Click()
Login.Tag = "0" & AppAccessLevel
Login.Show vbModal
End Sub

Private Sub mnuShow_Click()
Login.Tag = "5" & AppAccessLevel
Login.Show vbModal
End Sub

