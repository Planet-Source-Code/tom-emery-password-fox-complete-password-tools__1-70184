Attribute VB_Name = "modEncryption"
Public Function EnCrypt(Data As String) As String
    Dim iCount As Integer
    For iCount = 1 To Len(Data)
        EnCrypt = EnCrypt & Chr(Asc(Mid(Data, iCount, 1)) + 100)
        DoEvents
    Next iCount
End Function

Public Function DeCrypt(Data As String) As String
    Dim iCount As Integer
    For iCount = 1 To Len(Data)
        DeCrypt = DeCrypt & Chr(Asc(Mid(Data, iCount, 1)) - 100)
        DoEvents
    Next iCount
End Function
