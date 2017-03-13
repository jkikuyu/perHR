Attribute VB_Name = "PASSWORD"
Option Explicit
'Declare variables.
    Global NewUser As Boolean
    Dim Counter
    Global x, f, t
    Dim Encrypt
    Global Timecount As Integer
    Dim ClientName
    Dim BeforeSpace As Integer, afterspace As Integer, Length As Integer, sizeofword As Integer
    Dim Position
    Dim PasswordEntry
    Dim FileData
    Dim spacefound, loginname, TypedPassword, RetrievedPassword, LastLetter
    Dim checkdigit
    Dim Decrypt
    'Global Whattodo
    Global trackstatus
    Global RealPass
    Global username As String
    Global userid As Long
    Global grpid As Long
    'Global loginname

Function DecryptPassword()
On Error GoTo ErrDetected
         Counter = Counter + 1
    Select Case Counter
        Case 1
            t = t / 11
            t = (t) - 7
        Case 2
            t = t / 22
            t = (t) - 3
        Case 3
            t = t / 33
            t = (t) - 18
        Case 4
            t = t / 44
            t = (t) - 53
        Case 5
            t = t / 55
            t = (t) - 27
        Case 6
            t = t / 66
            t = (t) - 1
        Case 7
            t = t / 77
            t = (t) - 36
        Case 8
            t = t / 88
            t = (t) - 37
        Case 9
            t = t / 99
            t = (t) - 13
        Case 10
            t = t / 110
            t = (t) - 14
    End Select
    
    DecryptPassword = Chr(t)

ErrDetected:
If Err Then
   'MsgBox "Wrong password typed!"
   resetvariables
   Exit Function
End If

End Function

Function Encryptpassword()
    'This functioning the password and then uses an encryption key to encrypt
    'the password.
    
    Counter = Counter + 1
    Select Case Counter
        Case 1
            t = Asc(t) + 7
            t = t * 11
        Case 2
            t = Asc(t) + 3
            t = t * 22
        Case 3
            t = Asc(t) + 18
            t = t * 33
        Case 4
            t = Asc(t) + 53
            t = t * 44
        Case 5
            t = Asc(t) + 27
            t = t * 55
        Case 6
            t = Asc(t) + 1
            t = t * 66
        Case 7
            t = Asc(t) + 36
            t = t * 77
        Case 8
            t = Asc(t) + 37
            t = t * 88
        Case 9
            t = Asc(t) + 13
            t = t * 99
        Case 10
            t = Asc(t) + 14
            t = t * 110
    End Select
    
    Encryptpassword = t

End Function

Sub resetvariables()
    'Dim ClientName
    BeforeSpace = 1
    afterspace = 0
    Length = 0
    sizeofword = 0
    Position = 0
    'count = 0
    PasswordEntry = ""
    'Dim FileData
    spacefound = 0
    loginname = ""
    LastLetter = 0
    checkdigit = 0
    Decrypt = 0
    Counter = 0
    'Timecount = 0
End Sub

