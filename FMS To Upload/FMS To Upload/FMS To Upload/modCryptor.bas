Attribute VB_Name = "modCryptor"
Option Explicit

Dim strInput As String
Dim strOutput As String
Dim strOneLet As String
Dim strAlLet As String
Dim strSearchString As String

Const strDLLPass = "HawkeR"
Public Const strPassword = "HawkeR"

Dim lngEncoded As Double
Dim dbl255Times As Double
Dim dblMultX As Double
Dim dblX, dblY As Double

Public Function Encrypt(strInput As String, strPassword As String) As String
If strPassword = strDLLPass Then
    strOneLet = ""
    strAlLet = ""
    dblX = 1
    dblY = 100
    
    If Len(strInput) <= 2000 And Len(strInput) > 0 Then
        While dblX <= Len(strInput)
            strOneLet = Mid(strInput, dblX, 1)
            
            Call setDblX
            Call passEncr
            
            Randomize
            strAlLet = strAlLet & strOneLet
            dblX = dblX + 1
        Wend
        
        Randomize
        
        'set this to the ouput and add header and footer
        If Len(strAlLet) > 0 Then
            strOutput = Chr(Int((255 * Rnd) + 1)) & Chr(Int((255 * Rnd) + 1)) & strAlLet & Chr(Int((255 * Rnd) + 1)) & Chr(Int((255 * Rnd) + 1))
            Encrypt = strOutput
             If Decrypt(strOutput, strDLLPass) <> strInput Then
                MsgBox "This text cannot be encrypted.", vbExclamation
                Encrypt = ""
            End If
        Else
            Encrypt = ""
        End If
    ElseIf Len(strInput) > 2000 Then
        Encrypt = "<Cannot encrypt more than 2000 characters at one step. Convert using a loop.>"
    Else
        Encrypt = ""
    End If
Else
    MsgBox "The password is incorrect.", vbExclamation
End If
End Function

Private Sub passEncr()
lngEncoded = Asc(strOneLet) + Len(strInput) - (dblX * dblMultX)
dbl255Times = 1
If lngEncoded < 0 Then
    While lngEncoded + (255 * dbl255Times) < 0
        dbl255Times = dbl255Times + 1
    Wend
    strOneLet = Chr(lngEncoded + (255 * dbl255Times))
ElseIf lngEncoded > 255 Then
    While lngEncoded - (255 * dbl255Times) > 255
        dbl255Times = dbl255Times + 1
    Wend
    strOneLet = Chr(lngEncoded - (255 * dbl255Times))
Else
    strOneLet = Chr(lngEncoded)
End If
End Sub

Private Sub setDblX()
If dblX Mod 9 = 0 Then
    dblMultX = 5.16
ElseIf dblX Mod 8 = 0 Then
    dblMultX = 8.47
ElseIf dblX Mod 7 = 0 Then
    dblMultX = 4.29
ElseIf dblX Mod 6 = 0 Then
    dblMultX = 3.64
ElseIf dblX Mod 5 = 0 Then
    dblMultX = 4.56
ElseIf dblX Mod 4 = 0 Then
    dblMultX = 7.88
ElseIf dblX Mod 3 = 0 Then
    dblMultX = 1.79
ElseIf dblX Mod 2 = 0 Then
    dblMultX = 2.12
Else
    dblMultX = 1.84
End If
End Sub

Public Function Decrypt(strInput As String, strPassword As String) As String
If strPassword = strDLLPass Then
    If Len(strInput) >= 5 And Len(strInput) <= 2005 Then
        'set this text
        strOneLet = ""
        strAlLet = ""
        dblX = 1
        
        strInput = Mid(strInput, 3, Len(strInput) - 4)
        
        While dblX <= Len(strInput)
            strOneLet = Mid(strInput, dblX, 1)
        
            Call setDblX
            Call passDecr
        
            strAlLet = strAlLet & strOneLet
            dblX = dblX + 1
        Wend
        Decrypt = strAlLet
    ElseIf Len(strInput) > 2005 Then
        Decrypt = "<Cannot decrypt more than 2000 characters at one step. Convert using a loop.>"
    Else
        Decrypt = ""
    End If
Else
    MsgBox "The password is incorrect.", vbExclamation
End If
End Function

Private Sub passDecr()
lngEncoded = Asc(strOneLet) - Len(strInput) + (dblX * dblMultX)
dbl255Times = 1
If lngEncoded < 0 Then
    While lngEncoded + (255 * dbl255Times) < 0
        dbl255Times = dbl255Times + 1
    Wend
    strOneLet = Chr(lngEncoded + (255 * dbl255Times))
ElseIf lngEncoded > 255 Then
    While lngEncoded - (255 * dbl255Times) > 255
        dbl255Times = dbl255Times + 1
    Wend
    strOneLet = Chr(lngEncoded - (255 * dbl255Times))
Else
    strOneLet = Chr(lngEncoded)
End If
End Sub
