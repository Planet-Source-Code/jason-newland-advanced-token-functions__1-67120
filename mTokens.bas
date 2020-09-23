Attribute VB_Name = "modTokens"
'Gettok functions, etc, for use with IRC scripting but
'may server other purposes
Option Explicit

'GetTok("text with delimiters", "position as a string eg: 1", character (ie: 32))
Public Function GetTok(ByVal strSource As String, ByVal strPosition As String, ByVal Token As Integer, Optional ByVal NumTokens As Integer) As String
    'by Jason James Newland (2006)
    'eg: gettok("hi world this is a test","2-",32) would
    'return "world this is a test" and
    'gettok("hi world this is a test","-2",32) would
    'return "hi world"
    'gettok("hi world this is a test","0",32) would return
    '6 as in 6 tokens delimited by chr(32) in the source
    'string
    'gettok("hi world this is a test","2 to 4",32) would
    'return "world this is"
    On Error Resume Next
    Dim Tokens() As String
    Dim intPosition As Integer
    Dim intTemp As Integer
    'pointer for gettok position
    Dim intPos As Integer
    'return rest of string after pointer or not
    Dim intBool As Boolean
    Dim intRev As Boolean
    Dim i As Long
    'temp string to return
    Dim strToken As String
    Dim NumOfToks As Integer
    
    'first if the position is 0 return the total number of tokens
    If Val(Replace(strPosition, "-", vbNullString)) = 0 Then
        Tokens() = Split(strSource, Chr$(Token))
        GetTok = UBound(Tokens) + 1
        Exit Function
    End If
    
    intPosition = InStrRev(strPosition, "-")
    intTemp = Len(strPosition) - 1
    
    If intPosition > 0 Then
        If intPosition - intTemp = 1 Then
            'ie its at the end of the string
            intBool = True
            intRev = False
            intPos = Val(Replace(strPosition, "-", vbNullString)) - 1
        Else
            'must be at the start so set single token only
            intBool = False
            intRev = True
            intPos = Val(Replace(strPosition, "-", vbNullString)) - 1
        End If
    Else
        'just go from start
        intBool = False
        intRev = False
        intPos = Val(strPosition) - 1
    End If
    'ok, we have our token positions lets do the dirty work
    'first split the tokens
    Tokens() = Split(strSource, Chr$(Token))
    
    'if position is #- go from position to end
    'also check NumTokens
    If intBool = True Then
        If intRev = False Then
            If NumTokens = 0 Then
                NumTokens = UBound(Tokens)
            Else
                NumTokens = NumTokens - 1
            End If
            NumOfToks = 0
            For i = LBound(Tokens) To UBound(Tokens)
                DoEvents
                If i >= intPos Then
                    If NumOfToks <= NumTokens Then
                        NumOfToks = NumOfToks + 1
                        strToken = strToken & Tokens(i) & Chr$(Token)
                    End If
                End If
            Next i
        End If
    End If
    
    'if position is -# go from beginning to position
    If intBool = False Then
        If intRev = True Then
            For i = LBound(Tokens) To UBound(Tokens)
                DoEvents
                If i <= intPos Then
                    strToken = strToken & Tokens(i) & Chr$(Token)
                End If
            Next i
        End If
    End If
    If intBool = False Then
        If intRev = False Then
            'just return the token
            strToken = Tokens(intPos)
        End If
    End If
    'trim the end of string
    If Right$(strToken, 1) = Chr$(Token) Then
        strToken = Left$(strToken, Len(strToken) - 1)
    End If
    'return it
    GetTok = strToken
End Function

'IsTok("text with delimiters", "comparator string", Character (ie: 44))
Function IsTok(ByVal strSource As String, ByVal strCompare As String, ByVal Token As Integer) As Boolean
    'eg: IsTok("this,is,a,test", "test", 44) = True
    'compares a source string to see if the occurance of
    'the string exists
    
    'ok first split the string into arrays
    On Error Resume Next
    Dim Tokens() As String
    Dim i As Integer
    
    Tokens = Split(strSource, Chr$(Token))
    
    'now do some matching
    For i = 0 To UBound(Tokens)
        DoEvents
        If LCase(Tokens(i)) = LCase(strCompare) Then
            IsTok = True
            Exit Function
        End If
    Next i
    IsTok = False
End Function

Function FindTok(ByVal strSource As String, ByVal strCompare As String, ByVal Occurance As Integer, ByVal Token As Integer) As Integer
    'finds and matches a token in a string of text and returns
    'its position number as an integer
    'use 0 as the 'Occurances' delimter to return the total
    'number of times the same string occurs in the source and
    '1 to return the first occurance token position
    
    'ok, first we have to split the string into arrays
    On Error Resume Next
    Dim Tokens() As String
    Dim i As Integer
    Dim Tok As Integer
    
    Tokens = Split(strSource, Chr$(Token))
    Tok = 0
    
    'now do some matching
    For i = 0 To UBound(Tokens)
        DoEvents
        If LCase(Tokens(i)) = LCase(strCompare) Then
            If Occurance = 0 Then
                Tok = Tok + 1
            Else
                Tok = i + 1
                Exit For
            End If
        End If
    Next i
    FindTok = Tok
End Function

Function AddTok(ByVal strSource As String, ByVal strAddString As String, ByVal Token As Integer) As String
    On Error Resume Next
    Dim strTemp As String
    
    'add the token
    strTemp = strSource & Chr$(Token) & strAddString
    
    'trim the token at the front of the string
    If Left$(strTemp, 1) = Chr$(Token) Then
        strTemp = Mid$(strTemp, 2)
    End If
    AddTok = strTemp
End Function

Function DelTok(ByVal strSource As String, ByVal Position As Integer, ByVal Token As Integer) As String
    'first we split the source into an array
    'then loop through looking for the position
    On Error Resume Next
    Dim Tokens() As String
    Dim i As Integer
    Dim strTemp As String
    Dim Tok As Integer
    Tok = Position - 1
    
    Tokens = Split(strSource, Chr$(Token))

    'remove the token
    For i = 0 To UBound(Tokens)
        DoEvents
        If i <> Tok Then
            strTemp = strTemp & Tokens(i) & Chr$(Token)
        End If
    Next i
    
    'trim the token off the end of the string
    If Right$(strTemp, 1) = Chr$(Token) Then
        strTemp = Left$(strTemp, Len(strTemp) - 1)
    End If
    DelTok = strTemp
End Function

'misc functions, InsTok (insert token at position), RepTok
'(replace token at position), PutTok (overwrites a token)
Function RepTok(ByVal strSource As String, ByVal strNewToken As String, ByVal Position As Integer, ByVal Token As Integer) As String
    'first split the tokens in to an array
    On Error Resume Next
    Dim Tokens() As String
    Dim i As Integer
    Dim Tok As Integer
    Dim TokTotal As Integer
    Dim strTemp As String
    
    Tok = Position - 1
    
    'get total number of tokens already
    TokTotal = GetTok(strSource, "0", Token) - 1
    
    Tokens = Split(strSource, Chr$(Token))
    
    'now to replace, if the token position is out of range
    'then simply add the token to the end
    If Tok <= TokTotal Then
        For i = 0 To UBound(Tokens)
            DoEvents
            If i <> Tok Then
                strTemp = strTemp & Tokens(i) & Chr$(Token)
            Else
                'now insert the new token
                strTemp = strTemp & strNewToken & Chr$(Token)
            End If
        Next i
    Else
        'if the token is out of range add it to the end
        strTemp = strSource & Chr$(Token) & strNewToken
    End If
    
    'trim the token off the end of string
    If Right$(strTemp, 1) = Chr$(Token) Then
        strTemp = Left$(strTemp, Len(strTemp) - 1)
    End If
    
    RepTok = strTemp
End Function

'InsTok doesn't overwrite a token but merly inserts it at
'the position
Function InsTok(ByVal strSource As String, ByVal strNewToken As String, ByVal Position As Integer, ByVal Token As Integer) As String
    'first split the tokens
    On Error Resume Next
    Dim Tokens() As String
    Dim i As Integer
    Dim Tok As Integer
    Dim TokTotal As Integer
    Dim strTemp As String
    
    Tokens = Split(strSource, Chr$(Token))
    Tok = Position - 1
    TokTotal = GetTok(strSource, "0", Token) - 1
    
    'insert the token at position or at end if position is
    'out of range
    If Tok <= TokTotal Then
        For i = 0 To UBound(Tokens)
            DoEvents
            If i <> Tok Then
                strTemp = strTemp & Tokens(i) & Chr$(Token)
            Else
                strTemp = strTemp & strNewToken & Chr$(Token) & Tokens(i) & Chr$(Token)
            End If
        Next i
    Else
        'add the token at the end if its out of range
        strTemp = strSource & Chr$(Token) & strNewToken
    End If
    
    'trim the token off the end of string
    If Right$(strTemp, 1) = Chr$(Token) Then
        strTemp = Left$(strTemp, Len(strTemp) - 1)
    End If
    
    InsTok = strTemp
End Function

'PutTok overwrites a token at the specified position
Function PutTok(ByVal strSource As String, ByVal strNewToken As String, ByVal Position As Integer, ByVal Token As Integer) As String
    'first split the tokens
    On Error Resume Next
    Dim Tokens() As String
    Dim i As Integer
    Dim Tok As Integer
    Dim strTemp As String
    
    Tokens = Split(strSource, Chr$(Token))
    Tok = Position - 1
    
    'insert the token at position or at end if position is
    'out of range
    For i = 0 To UBound(Tokens)
        DoEvents
        If i <> Tok Then
            strTemp = strTemp & Tokens(i) & Chr$(Token)
        Else
            strTemp = strTemp & strNewToken & Chr$(Token)
        End If
    Next i
    
    'trim the token off the end of string
    If Right$(strTemp, 1) = Chr$(Token) Then
        strTemp = Left$(strTemp, Len(strTemp) - 1)
    End If
    
    PutTok = strTemp
End Function
