Attribute VB_Name = "modList"
Option Explicit

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
 
 Public Const WM_VSCROLL = &H115
 Public Const SB_BOTTOM = 7

Public Sub ListBoxSave(Directory As String, TheList As ListBox)
    Dim SaveList As Long
    On Error Resume Next
    Open Directory$ For Output As #1
    For SaveList& = 0 To TheList.ListCount - 1
        Print #1, TheList.List(SaveList&)
    Next SaveList&
    Close #1
End Sub


Public Sub ListBoxRemove(Src As Object)

    Dim X As Integer
    If Src.ListIndex < 0 Then Exit Sub
    If Src.ListCount < 1 Then Exit Sub
    If Src.SelCount > 1 Then
restart:
        For X = 0 To Src.ListCount - 1
            If Src.Selected(X) Then
                Src.RemoveItem X
                GoTo restart
            End If
        Next X
    Else
        Src.RemoveItem Src.ListIndex
    End If

End Sub


Public Sub ListBoxMoveUp(lb As Object)

    Dim tmpField As String
    Dim i As Integer
    i = lb.ListIndex
    If lb.ListCount < 1 Then Exit Sub

    If i > 0 And i < lb.ListCount Then
        tmpField = lb.List(i - 1)
        lb.List(i - 1) = lb.List(i)
        lb.List(i) = tmpField
        lb.ListIndex = i - 1
        lb.Selected(i - 1) = True
        lb.Selected(i) = False
    End If

End Sub

Public Sub ListBoxMoveDown(lb As Object)

    Dim tmpField As String
    Dim i As Integer
    i = lb.ListIndex
    If lb.ListCount < 1 Then Exit Sub

    If i > -1 And i < lb.ListCount - 1 Then
        tmpField = lb.List(i + 1)
        lb.List(i + 1) = lb.List(i)
        lb.List(i) = tmpField
        lb.ListIndex = i + 1
        lb.Selected(i + 1) = True
        lb.Selected(i) = False
    End If

End Sub


Public Sub ListBoxLoad(Directory As String, TheList As ListBox)
    Dim MyString As String
    On Error Resume Next
    Open Directory$ For Input As #1
    While Not EOF(1)
        Input #1, MyString$
        DoEvents
        TheList.AddItem MyString$
    Wend
    Close #1
End Sub




