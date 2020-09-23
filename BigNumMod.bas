Attribute VB_Name = "BigNumMod"
Option Explicit

'##########################################################
'##########################################################
'######## Name:   Big numbers Library              ########
'########                                          ########
'######## Description: With this library           ########
'######## you can do everything                    ########
'######## around large numbers:                    ########
'######## standard operations: add, substract,     ########
'######## divide, multiply, power, mod, powermod   ########
'######## boolean operations: and, or, not, xor,   ########
'######## shl, shr, inc, dec                       ########
'######## convert bases: from 2-64 to 2-64         ########
'########                                          ########
'######## AND THAT ALL IN A (FOR VB) SHORT TIME    ########
'########                                          ########
'######## Author: HAANDI                           ########
'######## e-mail: haandi@gmx.de                    ########
'########------------------------------------------########
'########   !IF YOU LIKE THIS CODE PLEASE VOTE!    ########
'########------------------------------------------########
'##########################################################
'##########################################################

Private Const BigNumbers As String = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz{|"

'StrAdd = sA + sB
Public Function StrAdd(ByVal sA As String, ByVal sB As String) As String
Dim i As Integer, iCarry As Integer
Dim dTemp As Double, sTemp As String
Dim iLenA As Integer, iLenB As Integer, iLenT As Integer
Dim sOut As String
    'sA = StrTrim(sA)
    'sB = StrTrim(sB)
    If sB = "0" Or sB = "" Then StrAdd = sA: Exit Function
    If sA = "0" Or sA = "" Then StrAdd = sB: Exit Function
    
    iLenA = Len(sA)
    iLenB = Len(sB)
    If iLenA >= iLenB Then
        sA = String$(14 - (iLenA Mod 14), "0") & sA
        iLenA = Len(sA)
        sB = String$(iLenA - iLenB, "0") & sB
        iLenB = iLenA
    Else
        sB = String$(14 - (iLenB Mod 14), "0") & sB
        iLenB = Len(sB)
        sA = String$(iLenB - iLenA, "0") & sA
        iLenA = iLenB
    End If
    For i = iLenA - 13 To 1 Step -14
        dTemp = Int(Mid$(sA, i, 14)) + Int(Mid$(sB, i, 14)) + iCarry
        sTemp = LTrim$(Str$(dTemp))
        iLenT = Len(sTemp)
        iCarry = 0
        If iLenT < 14 Then sTemp = String$(14 - iLenT, "0") & sTemp
        sOut = Right$(sTemp, 14) & sOut
        If iLenT = 15 Then iCarry = 1
    Next i
    If iCarry Then sOut = iCarry & sOut
    StrAdd = StrTrim(sOut)
End Function

'StrSub = sA - sB
Public Function StrSub(ByVal sA As String, ByVal sB As String) As String
Dim i As Integer, iCarry As Integer
Dim iLenA As Integer, iLenB As Integer, iLenT As Integer
Dim dTemp As Double, sTemp As String, sOut As String
    'sA = StrTrim(sA)
    'sB = StrTrim(sB)
    If sB = "0" Or sB = "" Then StrSub = sA: Exit Function
    
    iLenA = Len(sA)
    iLenB = Len(sB)
    sA = String$(14 - (iLenA Mod 14), "0") & sA
    iLenA = Len(sA)
    sB = String$(iLenA - iLenB, "0") & sB
    iLenB = iLenA
    
    For i = iLenA - 13 To 1 Step -14
        dTemp = Int(Mid$(sA, i, 14)) - Int(Mid$(sB, i, 14)) - iCarry
        If dTemp < 0 Then
            iCarry = 1
            dTemp = dTemp + 100000000000000#
        Else: iCarry = 0
        End If
        sTemp = LTrim$(Str$(dTemp))
        iLenT = Len(sTemp)
        sTemp = String$(14 - iLenT, "0") & sTemp
        sOut = Right$(sTemp, 14) & sOut
    Next i
    StrSub = StrTrim(sOut)
End Function

'StrMult = sA * sB
Public Function StrMult(ByVal sA As String, ByVal sB As String) As String
Dim i As Integer, j As Integer
Dim iLenA As Integer, iLenB As Integer, iLenT As Integer
Dim iCarry As Integer, dTemp As Double
Dim sOut As String, sTemp As String, AddArray(9) As String
    'sA = StrTrim(sA)
    'sB = StrTrim(sB)
    If sB = "0" Or sB = "" Then StrMult = "0": Exit Function
    If sA = "0" Or sA = "" Then StrMult = "0": Exit Function
    
    sA = String$(14 - (iLenA Mod 14), "0") & sA
    iLenA = Len(sA)
    iLenB = Len(sB)
    For i = 1 To 9
        For j = iLenA - 13 To 1 Step -14
            dTemp = Int(Mid$(sA, j, 14)) * i + iCarry
            sTemp = LTrim$(Str$(dTemp))
            iLenT = Len(sTemp)
            iCarry = 0
            If iLenT < 14 Then sTemp = String$(14 - iLenT, "0") & sTemp
            AddArray(i) = Right$(sTemp, 14) & AddArray(i)
            If iLenT = 15 Then iCarry = Int(Left(sTemp, 1))
        Next j
        If iCarry Then AddArray(i) = iCarry & AddArray(i): iCarry = 0
    Next i
    For i = iLenB To 1 Step -1
        If Mid$(sB, i, 1) = "0" Then GoTo NextDigit
        sOut = StrAdd(sOut, AddArray(Int(Mid$(sB, i, 1))) & String$(iLenB - i, "0"))
NextDigit:
    Next i
    StrMult = sOut
End Function

'StrDiv = sA \ sB
Public Function StrDiv(ByVal sA As String, ByVal sB As String) As String
Dim sCurrent As String, sTmp As String, sCount As String, i As Long, sOut As String
Dim iLenB As Integer
    sCurrent = sA
    If StrGT(sB, sA) Then StrDiv = "0": Exit Function
    
    iLenB = Len(sB)
    While Not (StrGT(sB, sCurrent))
        i = Len(sCurrent) - iLenB + StrGT(sB, Left(sCurrent, iLenB))
        sTmp = sB & String$(i, "0")
        sCount = "1" & String$(i, "0")
        While StrGT(sCurrent, sTmp) Or sCurrent = sTmp
            sCurrent = StrSub(sCurrent, sTmp)
            sOut = StrAdd(sOut, sCount)
        Wend
    Wend
    StrDiv = sOut
End Function

'StrMod = sA mod sB
Public Function StrMod(ByVal sA As String, ByVal sB As String) As String
Dim sCurrent As String, sTmp As String, iLenB As Integer
    sCurrent = sA
    iLenB = Len(sB)
    If StrGT(sB, sA) Then StrMod = sA: Exit Function
    While Not (StrGT(sB, sCurrent))
        sTmp = sB & String$(Len(sCurrent) - iLenB + StrGT(sB, Left(sCurrent, iLenB)), "0")
        While StrGT(sCurrent, sTmp) Or sCurrent = sTmp
            sCurrent = StrSub(sCurrent, sTmp)
        Wend
    Wend
    StrMod = sCurrent
End Function

'StrPow = sA ^ sB
Public Function StrPow(ByVal sA As String, ByVal sB As String) As String
Dim sOut As String, PowerArray() As String, i As Byte, Counter As Long
    If sB = "0" Then StrPow = "1": Exit Function
    If sB = "1" Then StrPow = sA: Exit Function
    sOut = sA
    For Counter = 10 To 0 Step -1
        If StrGT(sB, LTrim$(Str$(2 ^ Counter))) Or sB = LTrim$(Str$(2 ^ Counter)) Then Exit For
    Next Counter
    ReDim PowerArray(Counter) As String
    PowerArray(0) = sA
    For i = 1 To Counter
        PowerArray(i) = StrMult(PowerArray(i - 1), PowerArray(i - 1))
    Next i
    sOut = "1"
    For Counter = 10 To 0 Step -1
        For i = 1 To Int(StrDiv(sB, LTrim$(Str$(2 ^ Counter))))
            sOut = StrMult(sOut, PowerArray(Counter))
        Next i
        sB = StrMod(sB, LTrim$(Str$(2 ^ Counter)))
    Next Counter
    StrPow = sOut
End Function

'StrInc = sA + 1
Public Function StrInc(ByVal sA As String) As String
Dim i As Long, iCarry As Byte, sOut As String, iTemp As Byte
    i = Len(sA)
    sOut = sA
    iCarry = 1
    While iCarry = 1
        iTemp = Int(Mid$(sA, i, 1)) + iCarry
        Mid$(sOut, i, 1) = iTemp Mod 10
        i = i - 1
        If i = 0 Then sOut = "1" & sOut: StrInc = sOut: Exit Function
        iCarry = iTemp \ 10
    Wend
    StrInc = sOut
End Function

'StrDec = sA - 1
Public Function StrDec(ByVal sA As String) As String
Dim i As Long, j As Long, iTemp As Byte
    For i = Len(sA) To 1 Step -1
        iTemp = Int(Mid$(sA, i, 1))
        If iTemp > 0 Then
            Mid$(sA, i, 1) = iTemp - 1
            Exit For
        End If
    Next i
    If i = 0 Then
        Err.Raise vbObjectError + 100, , "Cannot deal with negative numbers!"
        Exit Function
    End If
    For j = Len(sA) To i + 1 Step -1
        Mid$(sA, j, 1) = "9"
        Next j
    StrDec = StrTrim(sA)
End Function

'StrGT = (sA > sB) (True or False)
Public Function StrGT(sA As String, sB As String) As Boolean
Dim i As Integer, iTemp As Integer, iTemp2 As Integer
Dim iLenA As Integer, iLenB As Integer
    sA = StrTrim(sA)
    sB = StrTrim(sB)
    iLenA = Len(sA)
    iLenB = Len(sB)
    
    If iLenA > iLenB Then StrGT = True: Exit Function
    If iLenA < iLenB Then StrGT = False: Exit Function
    
    For i = 1 To iLenA
        iTemp = Asc(Mid$(sA, i, 1))
        iTemp2 = Asc(Mid$(sB, i, 1))
        If iTemp > iTemp2 Then StrGT = True: Exit Function
        If iTemp2 > iTemp Then StrGT = False: Exit Function
    Next i
End Function

'StrTrim = sA without leading "0"s
Public Function StrTrim(sA As String) As String
    If Len(sA) = 0 Then StrTrim = "0": Exit Function
    Do While Left(sA, 1) = "0"
        sA = Right(sA, Len(sA) - 1)
    Loop
    If sA = "" Then sA = "0"
    StrTrim = sA
End Function

'StrPowMod = sA ^ sB mod sC (best for RSA encryption)
Public Function StrPowMod(sA As String, sB As String, sC As String) As String
Dim Tmp As String, Tmp2 As String
If sB = "0" Then StrPowMod = "1": Exit Function
If sB = "1" Then StrPowMod = StrMod(sA, sC): Exit Function
'Tmp = StrPowMod(sA, StrHalf(sB), sC)
Tmp = StrPowMod(sA, StrHalf(sB), sC)
Tmp2 = StrMod(StrMult(Tmp, Tmp), sC)
If StrStraight(sB) Then StrPowMod = Tmp2: Exit Function
StrPowMod = StrMod(StrMult(Tmp2, sA), sC)
End Function

Public Function StrShl(ByVal sA As String, ByVal sB As String) As String
StrShl = StrMult(sA, StrPow("2", sB))
End Function

Public Function StrShr(ByVal sA As String, ByVal sB As String) As String
StrShr = StrDiv(sA, StrPow("2", sB))
End Function

'StrHalf = sA \ 2
Public Function StrHalf(sA As String) As String
Dim i As Integer, iCarry As Long, lTemp As Long, iLenA As Integer, sTemp As String
Dim sOut As String, iLenT As Integer
iLenA = Len(sA)
sA = String$(8 - (iLenA Mod 8), "0") & sA
iLenA = Len(sA)
For i = 1 To iLenA - 7 Step 8
lTemp = (Int(Mid$(sA, i, 8)) + iCarry)
sTemp = LTrim$(Str$(lTemp \ 2))
iLenT = Len(sTemp)
If iLenT < 8 Then sTemp = String$(8 - iLenT, "0") & sTemp
sOut = sOut & sTemp
iCarry = (lTemp Mod 2) * 100000000
Next i
StrHalf = StrTrim(sOut)
End Function

'StrStraight = (sA / 2 = sA \ 2) (True or False)
Public Function StrStraight(sA As String) As Boolean
If Int(Right$(sA, 1)) Mod 2 = 0 Then StrStraight = True
End Function

'BooleanOperation = sA Operator sB 'sA,sB have to be in HEX format (base=16)
'The operator NOT dos not need sB, so let it be a vbNullString there
Public Function BooleanOperation(ByVal sA As String, ByVal sB As String, Operator As String) As String
Dim iLenA As Integer, iLenB As Integer, i As Integer
Dim iTemp As Integer, sOut As String
    iLenA = Len(sA)
    iLenB = Len(sB)
    If iLenA > iLenB Then
        sB = String$(iLenA - iLenB, "0") & sB
        iLenB = iLenA
    ElseIf iLenA < iLenB Then
        sA = String$(iLenB - iLenA, "0") & sA
        iLenA = iLenB
    End If
    For i = 1 To iLenA
        Select Case Operator
            Case "XOR"
            iTemp = (InStr(1, BigNumbers, Mid$(sA, i, 1)) - 1) Xor (InStr(1, BigNumbers, Mid$(sB, i, 1)) - 1)
            sOut = sOut & Mid$(BigNumbers, iTemp + 1, 1)
            Case "AND"
            iTemp = (InStr(1, BigNumbers, Mid$(sA, i, 1)) - 1) And (InStr(1, BigNumbers, Mid$(sB, i, 1)) - 1)
            sOut = sOut & Mid$(BigNumbers, iTemp + 1, 1)
            Case "OR"
            iTemp = (InStr(1, BigNumbers, Mid$(sA, i, 1)) - 1) Or (InStr(1, BigNumbers, Mid$(sB, i, 1)) - 1)
            sOut = sOut & Mid$(BigNumbers, iTemp + 1, 1)
            Case "NOT"
            iTemp = (InStr(1, BigNumbers, Mid$(sA, i, 1)) - 1) Xor 15
            sOut = sOut & Mid$(BigNumbers, iTemp + 1, 1)
        End Select
    Next i
    BooleanOperation = StrTrim(sOut)
End Function

Public Function ConvFromBase10(ByVal sA As String, ToBase As String) As String
Dim NextDigit As Integer
    While sA <> "0"
        NextDigit = Int(StrMod(sA, ToBase))
        ConvFromBase10 = Mid$(BigNumbers, NextDigit + 1, 1) & ConvFromBase10
        sA = StrDiv(sA, ToBase)
    Wend
End Function

Public Function ConvToBase10(ByVal sA As String, FromBase As String) As String
Dim i As Long, sTemp As String
    If Len(sA) = 0 Then ConvToBase10 = "0": Exit Function
    For i = 0 To Len(sA) - 1
    sTemp = LTrim$(Str$(InStr(1, BigNumbers, Mid$(sA, Len(sA) - i, 1)) - 1))
        ConvToBase10 = StrAdd(ConvToBase10, StrMult(sTemp, StrPow(FromBase, LTrim$(Str$(i)))))
    Next i
End Function

Public Function ConvertBases(ByVal sA As String, FromBase As String, ToBase As String) As String
Dim sTemp As String
sTemp = ConvToBase10(sA, FromBase)
ConvertBases = ConvFromBase10(sTemp, ToBase)
End Function

