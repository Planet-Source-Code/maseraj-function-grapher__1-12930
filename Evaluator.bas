Attribute VB_Name = "Evaluator"
Option Explicit

Function Evaluate(ByRef expression As String, Optional ErrorHandler As Boolean = True, Optional extras1 = "", Optional extras2 = "") As String

Dim dummy As String
Dim errors As Double
Dim i As Variant
Dim a, e
Dim vals()
Dim vars()

vars = Array("e", "pi")
vals = Array("2.71828182845905", "3.14159265358979")


For i = 1 To Len(expression)
    If Mid(expression, i, 1) <> " " Then
        If dummy = "" Then
            dummy = Mid(expression, i, 1)
        Else
            dummy = dummy + Mid(expression, i, 1)
        End If
    End If
Next

dummy = LCase(dummy)

If IsArray(extras1) And IsArray(extras2) Then
    If UBound(extras1) = UBound(extras2) Then
        ReDim Preserve vars(0 To UBound(extras1) + 2)
        ReDim Preserve vals(0 To UBound(extras1) + 2)
        For i = 0 To UBound(extras1)
            vars(i + 2) = extras1(i)
            vals(i + 2) = extras2(i)
        Next
    End If

Else
    e = ((Exp(Val(1)) - Exp(-Val(1))) / 2) + ((Exp(Val(1)) + Exp(-Val(1))) / 2)

    vars = Array("e", "pi")
    vals = Array("2.71828182845905", "3.14159265358979")
End If


dummy = eval(dummy, errors, , vars, vals)

If dummy = "" And errors = 0 Then errors = 101

If errors <> 0 Then

If ErrorHandler <> False Then

Select Case errors

    Case 100
    MsgBox "ERROR!" & Chr(13) & Chr(13) & "Unforseen error!", vbDefaultButton1, "Error!"
    dummy = "ERROR"
    
    Case 101
    MsgBox "ERROR!" & Chr(13) & Chr(13) & "Missing argument!", vbDefaultButton1, "Error!"
    dummy = "ERROR"
    
    Case 1
    MsgBox "ERROR!" & Chr(13) & Chr(13) & "Division by zero!", vbDefaultButton1, "Error!"
    dummy = "ERROR"
    
    Case 2
    MsgBox "ERROR!" & Chr(13) & Chr(13) & "Invalid value!", vbDefaultButton1, "Error!"
    dummy = "ERROR"
    
    Case 3
    MsgBox "ERROR!" & Chr(13) & Chr(13) & "Square root of negative!", vbDefaultButton1, "Error!"
    dummy = "ERROR"
    
    Case 4
    MsgBox "ERROR!" & Chr(13) & Chr(13) & "Power of negative value!", vbDefaultButton1, "Error!"
    dummy = "ERROR"
    
    Case 5
    MsgBox "ERROR!" & Chr(13) & Chr(13) & "Variable undefined!", vbDefaultButton1, "Error!"
    dummy = "ERROR"
    
    Case 6
    MsgBox "ERROR!" & Chr(13) & Chr(13) & "Operant undefined!", vbDefaultButton1, "Error!"
    dummy = "ERROR"
    
    Case 7
    MsgBox "ERROR!" & Chr(13) & Chr(13) & "Negative value!", vbDefaultButton1, "Error!"
    dummy = "ERROR"
    
    Case 8
    MsgBox "ERROR!" & Chr(13) & Chr(13) & "Value non integer!", vbDefaultButton1, "Error!"
    dummy = "ERROR"

    

End Select

Else

dummy = "ERROR"

End If

End If

Evaluate = dummy

End Function

Private Function eval(expression As String, errors As Double, Optional OpIn As String, Optional extras1 = "", Optional extras2 = "") As String

Dim dummy As String, num As String, op As String, result As String, operation As String
Dim errs As Double
errs = errors
dummy = expression

Call TermCut(dummy, num, op, errs, extras1, extras2)
result = num
While ((dummy <> "") And (op <> ")")) Or (op = "pi" Or op = "e")
    If (Prec(OpIn) >= Prec(op) And OpIn <> "") Or op = "" Then
        If num = "" Then num = result
        eval = num
        OpIn = op
        expression = dummy
        errors = errs
        Exit Function
    Else
        operation = op
        result = solve(operation, result, eval(dummy, errs, op, extras1, extras2), errs, extras1, extras2)
    End If
Wend
OpIn = op
eval = result
expression = dummy
errors = errs

End Function

Private Function TermCut(expression As String, num As String, op As String, errors As Double, Optional extras1 = "", Optional extras2 = "") As String

Dim dummy As String
Dim expr As String
Dim errs As Double
Dim j As Double

errs = errors
dummy = expression

If Mid(dummy, 1, 1) = "(" Then
    
    dummy = Right(dummy, Len(dummy) - 1)
    op = ""
    num = eval(dummy, errs, op, extras1, extras2)
    op = isop(dummy)
    dummy = isop(dummy, True)
    
ElseIf isop(dummy) = "+" Or isop(dummy) = "-" Then
    num = "0"
    If isop(dummy) = "+" Then
        op = "p"
    ElseIf isop(dummy) = "-" Then
        op = "n"
    End If
    dummy = isop(dummy, True)

Else
    
    num = isnum(dummy, False, extras1, extras2, errs)
    If num <> "" Then dummy = isnum(dummy, True, extras1, extras2, errs)
    op = isop(dummy)
    dummy = isop(dummy, True)
            
End If

errors = errs
expression = dummy
End Function

Private Function solve(operator As String, value1 As String, value2 As String, errors As Double, Optional extras1 = "", Optional extras2 = "") As String

On Error GoTo errors
Dim j As Double
Dim errs As Double
errs = errors
Dim e As Double
e = Exp(1)
Dim pi As Double
pi = 4 * Atn(1)
If value2 = "" Then errs = 101

Select Case operator

    Case "!"
    If Val(value2) < 0 Then errs = 7
    If (Val(value2) - Fix(Val(value2))) <> 0 Then errs = 8
    solve = 1
    For j = 1 To Val(value2)
        solve = Val(solve) * j
    Next
    If Val(value2) = 0 Or Val(value2) = 1 Then solve = 1
        
    Case "imp"
    solve = Val(value1) Imp Val(value2)
    
    Case "xor"
    solve = Val(value1) Xor Val(value2)
    
    Case "or"
    solve = Val(value1) Or Val(value2)

    Case "and"
    solve = Val(value1) And Val(value2)
    
    Case "deg_rad"
    solve = Val(value2) * (pi / 180)

    Case "rad_deg"
    solve = Val(value2) * (180 / pi)

    Case "exp"
    solve = Exp(Val(value2))
    
    Case "abs"
    solve = Abs(Val(value2))
    
    Case "atn"
    solve = Atn(Val(value2))
    
    Case "fix"
    solve = Fix(Val(value2))
    
    Case "int"
    solve = Int(Val(value2))
    
    Case "rnd"
    If value2 = "" Then
    Randomize Time
    solve = Rnd
    Else
    solve = Rnd(-Abs(Val(value2)))
    End If
    
    Case "sgn"
    solve = Sgn(Val(value2))
    
    Case "sec"
    If Val(value2) <> pi / 2 Then
    solve = 1 / Cos(Val(value2))
    Else
    errs = 2
    End If
    
    Case "cosec"
    If Val(value2) <> 0 And Val(value2) <> pi Then
    solve = 1 / Sin(Val(value2))
    ElseIf Val(value2) = pi Then
    errs = 1
    ElseIf Val(value2) = 0 Then
    errs = 2
    End If

    Case "cotan"
    solve = 1 / Tan(Val(value2))
    
    Case "arcsin"
    If Val(value2) >= -1 And Val(value2) <= 1 Then
    solve = Atn(Val(value2) / Sqr(-Val(value2) * Val(value2) + 1))
    Else
    errs = 2
    End If
    
    Case "arccos"
    If Val(value2) >= -1 And Val(value2) <= 1 Then
    solve = Atn(-Val(value2) / Sqr(-Val(value2) * Val(value2) + 1)) + 2 * Atn(1)
    Else
    errs = 2
    End If

    Case "arcsec"
    If Val(value2) >= -1 And Val(value2) <= 1 Then
    solve = Atn(Val(value2) / Sqr(Val(value2) * Val(value2) - 1)) + Sgn(Val(value2) - 1) * (2 * Atn(1))
    Else
    errs = 2
    End If
    
    Case "arccosec"
    If Val(value2) >= -1 And Val(value2) <= 1 Then
    solve = Atn(Val(value2) / Sqr(Val(value2) * Val(value2) - 1)) + (Sgn(Val(value2)) - 1) * (2 * Atn(1))
    Else
    errs = 2
    End If
    
    Case "arccotan"
    solve = Atn(Val(value2)) + 2 * Atn(1)
    
    Case "hsin"
    solve = (Exp(Val(value2)) - Exp(-Val(value2))) / 2
    
    Case "hcos"
    solve = (Exp(Val(value2)) + Exp(-Val(value2))) / 2
    
    Case "htan"
    solve = (Exp(Val(value2)) - Exp(-Val(value2))) / (Exp(Val(value2)) + Exp(-Val(value2)))
    
    Case "hsec"
    solve = 2 / (Exp(Val(value2)) + Exp(-Val(value2)))
    
    Case "hcosec"
    solve = 2 / (Exp(Val(value2)) - Exp(-Val(value2)))
    
    Case "hcotan"
    solve = (Exp(Val(value2)) + Exp(-Val(value2))) / (Exp(Val(value2)) - Exp(-Val(value2)))
    
    Case "harcsin"
    solve = Log(Val(value2) + Sqr(Val(value2) * Val(value2) + 1))
    
    Case "harccos"
    solve = Log(Val(value2) + Sqr(Val(value2) * Val(value2) - 1))
    
    Case "harctan"
    solve = Log((1 + Val(value2)) / (1 - Val(value2))) / 2
        
    Case "harcsec"
    solve = Log((Sqr(-Val(value2) * Val(value2) + 1) + 1) / Val(value2))
    
    Case "harccosec"
    solve = Log((Sgn(Val(value2)) * Sqr(Val(value2) * Val(value2) + 1) + 1) / Val(value2))
    
    Case "harccotan"
    solve = Log((Val(value2) + 1) / (Val(value2) - 1)) / 2
    
    Case "sqr"
    If Val(value2) >= 0 Then
    solve = Sqr(Val(value2))
    Else
    errs = 3
    End If
    
    Case "sin"
    solve = Sin(Val(value2))

    Case "cos"
    solve = Cos(Val(value2))
    
    Case "log"
    If Val(value2) > 0 Then
    solve = Log(Val(value2)) / Log(10)
    Else
    errs = 2
    End If
    
    Case "ln"
    If Val(value2) > 0 Then
    solve = Log(Val(value2))
    Else
    errs = 2
    End If

    Case "tan"
    If Val(value2) <> pi / 2 Then
    solve = Tan(Val(value2))
    Else
    errs = 2
    End If
        
    Case value1 = ""
    errs = 101
    
    Case "="
    If Val(value1) = Val(value2) Then
    solve = 1
    Else
    solve = 0
    End If
    
    Case "<"
    If Val(value1) < Val(value2) Then
    solve = 1
    Else
    solve = 0
    End If
    
    Case ">"
    If Val(value1) > Val(value2) Then
    solve = 1
    Else
    solve = 0
    End If

    Case "<>"
    If Val(value1) <> Val(value2) Then
    solve = 1
    Else
    solve = 0
    End If

    Case "<="
    If Val(value1) <= Val(value2) Then
    solve = 1
    Else
    solve = 0
    End If

    Case ">="
    If Val(value1) >= Val(value2) Then
    solve = 1
    Else
    solve = 0
    End If

    Case "\"
    If Val(value2) <> 0 Then
    solve = Val(value1) \ Val(value2)
    Else
    errs = 1
    End If
    
    Case "+", "p"
    solve = Val(value1) + Val(value2)

    Case "-", "n"
    solve = Val(value1) - Val(value2)
    
    Case "*"
    solve = Val(value1) * Val(value2)
    
    Case "/"
    If Val(value2) <> 0 Then
    solve = Val(value1) / Val(value2)
    Else
    errs = 1
    End If
    
    Case "^"
    If Val(value1) <= 0 And Abs(Val(value2)) < 1 Then
    errs = 4
    Else
    solve = Val(value1) ^ Val(value2)
    End If
    
    Case Else
    errs = 6
    
    End Select
    
errors = errs

Exit Function

errors:
errors = 100

End Function

Private Function isnum(value As String, Optional cut As Boolean, Optional extras1 = "", Optional extras2 = "", Optional errors As Double) As String

Dim dummy As String
Dim i, j, ncons, ncd
Dim cons As String, cd As String

For i = 1 To Len(value)

    Select Case Mid(value, i, 1)
    
    Case " "
        
    Case "a" To "z", "_", "!"
    If dummy = "" Then
    If cons = Empty Then
        cons = Mid(value, i, 1)
    Else
        cons = cons + Mid(value, i, 1)
    End If
        
    If IsArray(extras1) Then
    For j = 0 To UBound(extras1)
        If cons = extras1(j) Then
            cd = extras2(j)
        End If
    Next
    If cd = "" And cons = "" Then errors = 5
    End If
    
    Else
    Exit For
    End If
    
    Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
        If dummy = Empty Then
            dummy = Mid(value, i, 1)
        Else
            dummy = dummy + Mid(value, i, 1)
        End If
        
    Case "."
        If dummy = Empty Then
            dummy = Mid(value, i, 1)
        Else
            dummy = dummy + "."
        End If
        
    Case ","
        If dummy = Empty Then
            dummy = Mid(value, i, 1)
        Else
            dummy = dummy + "."
        End If

    Case Else
    
    Exit For
    
    End Select

Next

If cd <> "" Then
ncons = cons
ncd = ""
If IsArray(extras1) Then
For j = 0 To UBound(extras1)
    If ncons = extras1(j) Then
        ncd = extras2(j)
    End If
Next
If ncd = "" Then
cd = ""
i = 0
End If
End If
End If

If cut And i <> 0 Then
    isnum = Mid(value, i)
    Exit Function
End If
If cd <> "" Then dummy = cd
isnum = dummy
    
End Function


Private Function isop(value As String, Optional cut As Boolean) As String

Dim i, mids, ascval
Dim dummy As String
For i = 1 To Len(value)
    mids = Mid(value, i, 1)
    
    Select Case mids
    
    Case " "
    
    Case ")"
        If i = 1 Then
        dummy = ")"
        Exit For
        End If
        
    Case "<", "=", ">"
        Select Case Right(dummy, 1)
            
            Case "a" To "z", "_", "!"
            Exit For
            
            Case Else
            If dummy = Empty Then
                dummy = mids
            Else
                dummy = dummy + mids
            End If
            
        End Select
            
    Case "a" To "z", "_", "!"

        Select Case Right(dummy, 1)
        
            Case "<", "=", ">"
                Exit For
            Case Else
                If dummy = Empty Then
                    dummy = mids
                Else
                    dummy = dummy + mids
                End If
                
        End Select
        
    Case "/", "*", "-", "+", "^", "\"
        If dummy = "" Then
        dummy = Mid(value, i, 1)
        Exit For
        Else
        Exit For
        End If
        
    Case Else
    
    Exit For
    
    End Select

Next

If cut And i <> 0 Then
    If i = 1 Then i = 2
    isop = Mid(value, i, Len(value))
    Exit Function
End If

isop = dummy

End Function

Private Function Prec(string1 As String) As Integer

Dim ops, opvals, i
ops = Array("+", "-", "*", "/", "\", "^", "sqr", "sin", "cos", "log", "ln", "tan", "atn", "logs", "p", "n", "(", ")", "abs", "fix", "int", "rnd", "sgn", "sec", "cosec", "cotan", "arcsin", "arccos", "arcsec", "arccosec", "arccotan", "hsin", "hcos", "htan", "hsec", "hcosec", "hcotan", "harcsin", "harcos", "harctan", "harsec", "harcosec", "harcotan", "=", "<>", "<", ">", "<=", ">=", "and", "or", "xor", "imp", "int", "!", "exp")
opvals = Array(1, 1, 2, 2, 2, 3, 4, 4, 4, 4, 4, 4, 4, 4, 5, 5, 6, 6, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 0, -1, -2, -3, -4, -5, -6, -7, -8, -9, 4, 4, 4)
For i = 0 To UBound(ops)
    If string1 = ops(i) Then
    Prec = opvals(i)
    End If
Next
End Function
