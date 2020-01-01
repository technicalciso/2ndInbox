Sub MoveToSecond()
On Error Resume Next

Dim ns As Outlook.NameSpace
Set ns = Application.GetNamespace("MAPI")

Dim moveToFolder As Outlook.MAPIFolder
Set moveToFolder = ns.Folders("email@example.com").Folders("Inbox").Folders("2nd inbox")

If moveToFolder Is Nothing Then
    MsgBox "Target folder not found.", vbOKOnly + vbExclamation, "Move Macro Error"
End If

If Application.ActiveExplorer.Selection.Count = 0 Then
    MsgBox "No item selected."
    Exit Sub
End If

For Each objItem In Application.ActiveExplorer.Selection
    objItem.Move moveToFolder
Next

Set ns = Nothing
Set moveToFolder = Nothing
Set objItem = Nothing

End Sub

Sub MoveToSecondMsg()
On Error Resume Next

Dim ns As Outlook.NameSpace
Set ns = Application.GetNamespace("MAPI")

Dim moveToFolder As Outlook.MAPIFolder
Set moveToFolder = ns.Folders("email@example.com").Folders("Inbox").Folders("2nd inbox")

If moveToFolder Is Nothing Then
    MsgBox "Target folder not found.", vbOKOnly + vbExclamation, "Move Macro Error"
End If

Dim myInspector As Outlook.Inspector
Dim myItem As Outlook.MailItem
Set myInspector = Application.ActiveInspector
Set myItem = myInspector.CurrentItem
myItem.Move moveToFolder

Set ns = Nothing
Set moveToFolder = Nothing
Set objItem = Nothing
Set myInspector = Nothing
Set myItem = Nothing

End Sub

