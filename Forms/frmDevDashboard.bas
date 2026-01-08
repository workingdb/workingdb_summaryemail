Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub backBtn_Click()
Application.Quit
End Sub

Private Sub disShift_Click()
If (privilege("admin", Environ("username")) = False) Then
    MsgBox "You need admin privilege to do this", vbCritical, "Access Denied"
    Exit Sub
End If
If (privilege("developer", Environ("username")) = False) Then
    MsgBox "You need developer privilege to do this", vbCritical, "Access Denied"
    Exit Sub
End If

ap_DisableShift
Me.ShortcutMenu = False

End Sub

Private Sub enableShift_Click()

If (privilege("admin", Environ("username")) = False) Then
    MsgBox "You need admin privilege to do this", vbCritical, "Access Denied"
    Exit Sub
End If
If (privilege("developer", Environ("username")) = False) Then
    MsgBox "You need developer privilege to do this", vbCritical, "Access Denied"
    Exit Sub
End If

ap_EnableShift
Me.ShortcutMenu = True

End Sub

Private Sub Form_Load()

Me.ShortcutMenu = False

If (privilege("admin", Environ("username")) = False) Then
    MsgBox "You need admin privilege to open this page", vbCritical, "Access Denied"
    GoTo killThis
End If
If (privilege("developer", Environ("username")) = False) Then
    MsgBox "You need developer privilege to open this page", vbCritical, "Access Denied"
    GoTo killThis
End If

Exit Sub

killThis:
MsgBox "You must be a developer to open this", vbCritical, "Denied."
Application.Quit
End Sub

Private Sub hideNav_Click()

Call DoCmd.NavigateTo("acNavigationCategoryObjectType")
Call DoCmd.RunCommand(acCmdWindowHide)

End Sub

Private Sub showNav_Click()

Call DoCmd.SelectObject(acTable, , True)
Me.ShortcutMenu = True

End Sub
