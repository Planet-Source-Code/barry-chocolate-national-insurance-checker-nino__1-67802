Attribute VB_Name = "modNINO"
Option Explicit
'checks whether a National Insurance Number (NINO) is in the correct format
'Returns true if the NINO is in the correct format
'The NINO should be passed as a string
'The length should be a minimum of 8 (without a suffix)
'And a maximum of 9 (with a suffix)
'1. Must be 8 or 9 characters.
'2. First 2 characters must be alpha.
'3. Next 6 characters must be numeric.
'4. Final character(suffix, if included) can be A, B, C, D.
'5. First character must not be D,F,I,Q,U or V
'6. Second characters must not be D, F, I, O, Q, U or V.
'7. First 2 characters must not be combinations of GB, NK, TN or ZZ
'(the term combinations covers both GB and BG etc.)
Public Function IsNinoValid(ByRef NINO As String) As Boolean
    On Error GoTo ErrIsNinoValid
    If Len(NINO) < 8 Or Len(NINO) > 9 Then
        IsNinoValid = False
        Exit Function
    End If
    NINO = UCase(NINO)
    If Len(NINO) = 9 Then
        If NINO Like "[A-CEG-HJ-PR-TW-Z][A-CEG-HJ-NPR-TW-Z]######[A-D]" Then
            IsNinoValid = True
        End If
    Else
        If NINO Like "[A-CEG-HJ-PR-TW-Z][A-CEG-HJ-NPR-TW-Z]######" Then
            IsNinoValid = True
        End If
    End If
    If IsNinoValid = True Then
        Select Case Left(NINO, 2)
            Case "GB"
                IsNinoValid = False
            Case "NK"
                IsNinoValid = False
            Case "TN"
                IsNinoValid = False
            Case "ZZ"
                IsNinoValid = False
        End Select
    End If
    Exit Function
ErrIsNinoValid:
    IsNinoValid = False
    MsgBox "Error " & Err.Number & vbCrLf & Err.Description, vbCritical + vbOKOnly, "Nino Check Error"
End Function
