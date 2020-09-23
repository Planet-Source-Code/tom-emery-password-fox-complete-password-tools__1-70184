Attribute VB_Name = "modCheckKey"

'
Public Function KeyQuality(ByVal aKey As String) As Integer
' returns an integer value (0 to 100) rating the key quality
Dim QC As Integer
Dim LN As Integer
Dim k As Integer
Dim Uc As Boolean
Dim Lc As Boolean
Dim Wid As Integer
Dim ValidKey As Boolean
LN = Len(aKey)
QC = LN * 4
'check key lenght (at least 5 chars!)
If Len(aKey) < 5 Then KeyQuality = 0: Exit Function
' check for repetitions (abcabc, aaaaa, 121212, etc.)
For Wid = 1 To Int(Len(aKey) / 2)
    ValidKey = False
    For k = Wid + 1 To Len(aKey) Step Wid
        If Mid(aKey, 1, Wid) <> Mid(aKey, k, Wid) Then ValidKey = True: Exit For
    Next
If ValidKey = False Then Exit For
Next
If ValidKey = False Then KeyQuality = 0: Exit Function
'check ucases and lcases
For k = 1 To Len(aKey)
    If Asc(Mid(aKey, k, 1)) > 64 And Asc(Mid(aKey, k, 1)) < 91 Then Uc = True
    If Asc(Mid(aKey, k, 1)) > 96 And Asc(Mid(aKey, k, 1)) < 123 Then Lc = True
Next
If Uc = True And Lc = True Then QC = QC * 1.5
'check numbers
For k = 1 To Len(aKey)
    If Asc(Mid(aKey, k, 1)) > 47 And Asc(Mid(aKey, k, 1)) < 58 Then
        If Uc = True Or Lc = True Then QC = QC * 1.5
        Exit For
        End If
Next
'check signs
For k = 1 To Len(aKey)
    If Asc(Mid(aKey, k, 1)) < 48 Or Asc(Mid(aKey, k, 1)) > 122 Or (Asc(Mid(aKey, k, 1)) > 57 And Asc(Mid(aKey, k, 1)) < 65) Then QC = QC * 1.5: Exit For
Next
If QC > 100 Then QC = 100
KeyQuality = Int(QC)
End Function

