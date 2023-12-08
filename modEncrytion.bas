Attribute VB_Name = "modEncrytion"
Option Explicit

Public Function edwinEncryption(strInput) As String
  Dim key As String, j As Integer
  Dim i As Integer, a As String, p As String, keyCount As Integer
  key = "HydrographicOfficeMarineDepartment"     'Encryption Key
  keyCount = Len(key)
  j = 1
  edwinEncryption = ""
  For i = 1 To Len(strInput)
    p = Asc(Mid(key, j, 1)) Xor Asc(Mid(strInput, i, 1))
    j = j + 1
    If j > keyCount Then
      j = 1
    End If
    edwinEncryption = edwinEncryption & Chr(p)
  Next
End Function
