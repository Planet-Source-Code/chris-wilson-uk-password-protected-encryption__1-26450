Attribute VB_Name = "Enc"
Public Function Encrypt(Text As String, Password As String) As String
    Dim cipher As New Coding
    Dim Password2 As String * 10
    Dim PTag As String * 3
    Password2 = Password
    Text = Text & Password2

    cipher.KeyString = Password2
    cipher.Text = Text
    cipher.DoXor
    cipher.Stretch
    Encrypt = cipher.Text & "<e"
    End Function
    
Public Function Decrypt(Text As String, Password As String) As String

    If Not Right(Text, 2) = "<e" Then Decrypt = "ERR1": Exit Function
    Text = Mid(Text, 1, Len(Text) - 2)
       
    Dim cipher As New Coding
    Dim Password2 As String * 10
    Password2 = Password
       
    cipher.KeyString = Password2
    cipher.Text = Text
    cipher.Shrink
    cipher.DoXor
    
       
    On Error GoTo 10
    
    If Not Mid(cipher.Text, Len(cipher.Text) - 9) = Password2 Then
10
    Decrypt = "ERR2"
    Exit Function
    End If
    
    
    Decrypt = Mid(cipher.Text, 1, Len(cipher.Text) - 10)
    
End Function
