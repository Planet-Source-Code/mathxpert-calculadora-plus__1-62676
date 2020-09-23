VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculadora Plus"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   5145
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMod 
      Caption         =   "Mod"
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton cmdAns 
      Caption         =   "Ans"
      Height          =   375
      Left            =   960
      TabIndex        =   5
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton cmdInt 
      Caption         =   "Int"
      Height          =   375
      Left            =   120
      TabIndex        =   47
      Top             =   5040
      Width           =   735
   End
   Begin VB.CommandButton cmdRound 
      Caption         =   "Round"
      Height          =   375
      Left            =   2640
      TabIndex        =   50
      ToolTipText     =   ">RoundTo([how many decimal places?])"
      Top             =   5040
      Width           =   735
   End
   Begin VB.CommandButton cmdDecPoint 
      Caption         =   "."
      Height          =   375
      Left            =   1800
      TabIndex        =   49
      Top             =   5040
      Width           =   735
   End
   Begin VB.CommandButton cmdNum0 
      Caption         =   "0"
      Height          =   375
      Left            =   960
      TabIndex        =   48
      Top             =   5040
      Width           =   735
   End
   Begin VB.CommandButton cmdNum3 
      Caption         =   "3"
      Height          =   375
      Left            =   2640
      TabIndex        =   44
      Top             =   4560
      Width           =   735
   End
   Begin VB.CommandButton cmdNum2 
      Caption         =   "2"
      Height          =   375
      Left            =   1800
      TabIndex        =   43
      Top             =   4560
      Width           =   735
   End
   Begin VB.CommandButton cmdNum1 
      Caption         =   "1"
      Height          =   375
      Left            =   960
      TabIndex        =   42
      Top             =   4560
      Width           =   735
   End
   Begin VB.CommandButton cmdNum6 
      Caption         =   "6"
      Height          =   375
      Left            =   2640
      TabIndex        =   38
      Top             =   4080
      Width           =   735
   End
   Begin VB.CommandButton cmdNum5 
      Caption         =   "5"
      Height          =   375
      Left            =   1800
      TabIndex        =   37
      Top             =   4080
      Width           =   735
   End
   Begin VB.CommandButton cmdNum4 
      Caption         =   "4"
      Height          =   375
      Left            =   960
      TabIndex        =   36
      Top             =   4080
      Width           =   735
   End
   Begin VB.CommandButton cmdEquals 
      Caption         =   "="
      Height          =   855
      Left            =   4320
      TabIndex        =   46
      Top             =   4560
      Width           =   735
   End
   Begin VB.CommandButton cmdPi 
      Caption         =   "Pi"
      Height          =   375
      Left            =   3480
      TabIndex        =   51
      Top             =   5040
      Width           =   735
   End
   Begin VB.CommandButton cmdPercent 
      Caption         =   "%"
      Height          =   375
      Left            =   3480
      TabIndex        =   45
      Top             =   4560
      Width           =   735
   End
   Begin VB.CommandButton cmdPlus 
      Caption         =   "+"
      Height          =   375
      Left            =   3480
      TabIndex        =   39
      Top             =   4080
      Width           =   735
   End
   Begin VB.CommandButton cmdDividedBy 
      Caption         =   "÷"
      Height          =   375
      Left            =   4320
      TabIndex        =   40
      Top             =   4080
      Width           =   735
   End
   Begin VB.CommandButton cmdTimes 
      Caption         =   "×"
      Height          =   375
      Left            =   4320
      TabIndex        =   34
      Top             =   3600
      Width           =   735
   End
   Begin VB.CommandButton cmdMinusOrNeg 
      Caption         =   "-"
      Height          =   375
      Left            =   3480
      TabIndex        =   33
      Top             =   3600
      Width           =   735
   End
   Begin VB.CommandButton cmdNum9 
      Caption         =   "9"
      Height          =   375
      Left            =   2640
      TabIndex        =   32
      Top             =   3600
      Width           =   735
   End
   Begin VB.CommandButton cmdNum8 
      Caption         =   "8"
      Height          =   375
      Left            =   1800
      TabIndex        =   31
      Top             =   3600
      Width           =   735
   End
   Begin VB.CommandButton cmdNum7 
      Caption         =   "7"
      Height          =   375
      Left            =   960
      TabIndex        =   30
      Top             =   3600
      Width           =   735
   End
   Begin VB.CommandButton cmdClosedPar 
      Caption         =   ")"
      Height          =   375
      Left            =   120
      TabIndex        =   41
      Top             =   4560
      Width           =   735
   End
   Begin VB.CommandButton cmdOpenPar 
      Caption         =   "("
      Height          =   375
      Left            =   120
      TabIndex        =   35
      Top             =   4080
      Width           =   735
   End
   Begin VB.CommandButton cmdBackspace 
      Caption         =   "Bksp"
      Height          =   375
      Left            =   4320
      TabIndex        =   28
      Top             =   3120
      Width           =   735
   End
   Begin VB.CommandButton cmdExp 
      Caption         =   "e^x"
      Height          =   375
      Left            =   4320
      TabIndex        =   16
      Top             =   2160
      Width           =   735
   End
   Begin VB.CommandButton cmdNeg1Power 
      Caption         =   "x^-1"
      Height          =   375
      Left            =   2640
      TabIndex        =   26
      Top             =   3120
      Width           =   735
   End
   Begin VB.CommandButton cmdSquared 
      Caption         =   "x^2"
      Height          =   375
      Left            =   1800
      TabIndex        =   25
      Top             =   3120
      Width           =   735
   End
   Begin VB.CommandButton cmdSqr 
      Caption         =   "Sqr"
      Height          =   375
      Left            =   120
      TabIndex        =   23
      Top             =   3120
      Width           =   735
   End
   Begin VB.CommandButton cmdToFraction 
      Caption         =   "> Frac"
      Height          =   375
      Left            =   120
      TabIndex        =   29
      Top             =   3600
      Width           =   735
   End
   Begin VB.CommandButton cmdGetBase 
      Caption         =   "x^1/y"
      Height          =   375
      Left            =   960
      TabIndex        =   24
      Top             =   3120
      Width           =   735
   End
   Begin VB.CommandButton cmdXtoYpower 
      Caption         =   "x^y"
      Height          =   375
      Left            =   3480
      TabIndex        =   27
      Top             =   3120
      Width           =   735
   End
   Begin VB.CommandButton cmdAtn 
      Caption         =   "Tan^-1"
      Height          =   375
      Left            =   4320
      TabIndex        =   22
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton cmdAcs 
      Caption         =   "Cos^-1"
      Height          =   375
      Left            =   2640
      TabIndex        =   20
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton cmdAsn 
      Caption         =   "Sin^-1"
      Height          =   375
      Left            =   960
      TabIndex        =   18
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton cmdTan 
      Caption         =   "Tan"
      Height          =   375
      Left            =   3480
      TabIndex        =   21
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton cmdCos 
      Caption         =   "Cos"
      Height          =   375
      Left            =   1800
      TabIndex        =   19
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton cmdSin 
      Caption         =   "Sin"
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton cmdLn 
      Caption         =   "Ln"
      Height          =   375
      Left            =   3480
      TabIndex        =   15
      Top             =   2160
      Width           =   735
   End
   Begin VB.CommandButton cmdLog 
      Caption         =   "Log"
      Height          =   375
      Left            =   2640
      TabIndex        =   14
      Top             =   2160
      Width           =   735
   End
   Begin VB.CommandButton cmdENeg 
      Caption         =   "E-"
      Height          =   375
      Left            =   1800
      TabIndex        =   13
      Top             =   2160
      Width           =   735
   End
   Begin VB.CommandButton cmdEPlus 
      Caption         =   "E+"
      Height          =   375
      Left            =   960
      TabIndex        =   12
      Top             =   2160
      Width           =   735
   End
   Begin VB.CommandButton cmdFactorial 
      Caption         =   "x!"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   2160
      Width           =   735
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   10
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton cmdRecip 
      Caption         =   "1 / x"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   735
   End
   Begin VB.Frame framAngleUnits 
      Caption         =   "Unit of Angles"
      Height          =   645
      Left            =   2610
      TabIndex        =   7
      Top             =   1470
      Width           =   1665
      Begin VB.CommandButton cmdChange 
         Caption         =   "Change"
         Height          =   375
         Left            =   870
         TabIndex        =   9
         Top             =   210
         Width           =   735
      End
      Begin VB.Label lblDRG 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         Caption         =   "DEG"
         Height          =   255
         Left            =   150
         TabIndex        =   8
         Top             =   275
         Width           =   495
      End
   End
   Begin VB.TextBox txtOutput 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1080
      Width           =   4935
   End
   Begin VB.TextBox txtInput 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   4935
   End
   Begin VB.Label lblOutput 
      Caption         =   "Output:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   615
   End
   Begin VB.Label lblInput 
      Caption         =   "Input:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function PurelyNumeric(Expression) As Boolean
PurelyNumeric = IsNumeric(Expression) And (InStr(CStr(Expression), "(") = 0 And InStr(CStr(Expression), ")") = 0)
End Function

Private Function IsNegative(ByVal Number As Double) As Boolean
IsNegative = (Number < 0)
End Function

Private Function Opposite(ByVal Number As Double) As Double
Opposite = -Number
End Function

Private Function DoCount(ByVal sInput As String, ByVal sMatch As String, Optional ByVal Compare As VbCompareMethod = vbBinaryCompare) As Long
Dim lLen As Long
Dim i As Long
Dim j As Long

lLen = Len(sMatch)
If lLen = 0 Then Exit Function

Do
    i = InStr(i + lLen, sInput, sMatch, Compare)
    If i > 0 Then
        j = j + 1
    Else
        Exit Do
    End If
Loop

DoCount = j
End Function

Private Function MyRound(ByVal Number As Double, Optional ByVal NumDecimalPlaces As Long = 0) As Double
Dim tNumber As Double
Dim iNumber As Double

tNumber = Number * 10 ^ CDbl(NumDecimalPlaces)
iNumber = Fix(tNumber)

If Abs(tNumber - iNumber) >= 0.5 Then
    If IsNegative(tNumber) Then
        iNumber = iNumber - 1
    Else
        iNumber = iNumber + 1
    End If
End If

MyRound = iNumber / 10 ^ CDbl(NumDecimalPlaces)
End Function

Private Function GetFraction(ByVal Number As Double) As String
On Error GoTo ReturnDecimal

Dim Numerator As Long
Dim Denominator As Long
Dim fNumber As Double
Dim IterationControl As Long

Numerator = 1
Denominator = 1

fNumber = CDbl(Numerator) / CDbl(Denominator)

Do While CStr(fNumber) <> CStr(Number)
    
    If fNumber < Number Then
        Numerator = Numerator + 1
    Else
        Denominator = Denominator + 1
        Numerator = CLng(Number * CDbl(Denominator))
    End If
    fNumber = CDbl(Numerator) / CDbl(Denominator)
    
    IterationControl = IterationControl + 1
    If IterationControl >= 1000000 Then GoTo ReturnDecimal
Loop

If Denominator = 1 Then GoTo ReturnDecimal
GetFraction = CStr(Numerator) & "/" & CStr(Denominator)

Exit Function
ReturnDecimal: GetFraction = CStr(Number)
End Function

Private Function RangeIt(ByVal Exp As String, ByVal Start As Long, InPercent As Boolean, InFactorial As Boolean, Optional Reverse As Boolean = False, Optional InPar As Boolean) As Long
On Error Resume Next

Dim tStart As Long
Dim X As String
Dim y As String
Dim bSkip As Boolean
Dim MyExpr As String

bSkip = False
InPercent = False
InFactorial = False
InPar = False
tStart = Start

Do
    tStart = tStart + IIf(Reverse, -1, 1)
    If tStart > Len(Exp) Or tStart <= 0 Then Exit Do
    
    If bSkip Then
        bSkip = False
    Else
        X = UCase$(Mid$(Exp, tStart, 1))
        Select Case X
            Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", ".", "%", "!", "(", ")", ","
            Case "E"
                If Not Reverse Then
                    y = Mid$(Exp, tStart + 1, 1)
                    If y = "+" Or y = "-" Then
                        bSkip = True
                    Else
                        Exit Do
                    End If
                End If
            Case "-"
                If Reverse Then
                    y = UCase$(Mid$(Exp, tStart - 1, 1))
                    Select Case y
                        Case "+", "-", "*", "/", "": tStart = tStart - 1: Exit Do
                        Case "(": tStart = tStart - 2: Exit Do
                        Case "E": bSkip = True
                        Case Else: Exit Do
                    End Select
                Else
                    If tStart - Start > 1 Then Exit Do
                End If
            Case "+"
                If Reverse Then
                    y = UCase$(Mid$(Exp, tStart - 1, 1))
                    If y = "E" Then
                        bSkip = True
                    Else
                        Exit Do
                    End If
                Else
                    Exit Do
                End If
            Case Else: Exit Do
        End Select
    End If
Loop

tStart = tStart + IIf(Reverse, 1, -1)

If Reverse Then
    MyExpr = Mid$(Exp, tStart, Start - tStart)
Else
    MyExpr = Mid$(Exp, Start + 1, tStart - Start)
End If

InPercent = (Right$(Replace(MyExpr, "!", ""), 1) = "%")
InFactorial = (Right$(Replace(MyExpr, "%", ""), 1) = "!")
InPar = ((Left$(MyExpr, 1) = "(") And (Right$(MyExpr, 1) = ")"))

RangeIt = tStart
End Function

'THE MOST IMPORTANT FUNCTION OF THE CALCULATOR!!!
Private Function EvaluateInput(ByVal sInput As String) As String
On Error GoTo ProcExit

Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long
Dim m As Long
Dim tOutput As String
Dim tmpStr As String
Dim tmpStr2 As String
Dim tmpStr3 As String
Dim tmpExp As String
Dim tmpExp2 As String
Dim tmpNum As Double
Dim tmpNumerator As Double
Dim tmpDenominator As Double
Dim NoFraction As Boolean

Dim lOpenParCount As Long
Dim lClosedParCount As Long

Dim ColPar As Collection

Dim bPercent As Boolean
Dim bPercent2 As Boolean

Dim bFactorial As Boolean
Dim bFactorial2 As Boolean

Dim bPar As Boolean

Dim D1 As Double
Dim D2 As Double

NoFraction = True
tmpStr = LCase$(sInput)
tmpStr = Replace(tmpStr, " ", "")
tmpStr = Replace(tmpStr, "pi", Replace(CStr(Pi), ",", "."))

If txtOutput = "" Then tOutput = "0" Else tOutput = "(" & txtOutput & ")"
tmpStr = Replace(tmpStr, "ans", tOutput)

i = InStr(tmpStr, ">")
If i > 0 Then
    tmpStr2 = Left$(tmpStr, i - 1)
Else
    tmpStr2 = tmpStr
End If

lOpenParCount = DoCount(tmpStr2, "(")
lClosedParCount = DoCount(tmpStr2, ")")

If lOpenParCount > lClosedParCount Then
    tmpStr2 = tmpStr2 & String$(lOpenParCount - lClosedParCount, ")")
ElseIf lClosedParCount > lOpenParCount Then
    EvaluateInput = "Syntax error"
    Exit Function
End If

Do
    i = InStr(k + 1, tmpStr2, "(")
    j = InStr(k + 1, tmpStr2, ")")
    
    If i > 1 Or j > 1 Then
        If i <= 1 Or j < i Then
            k = j
            If PurelyNumeric(Mid$(tmpStr2, k + 1, 1)) Or (Mid$(tmpStr2, k + 1, 1) = "x") Or (Mid$(tmpStr2, k + 1, 1) = "(") Then
                tmpStr2 = Left$(tmpStr2, k) & "*" & Mid$(tmpStr2, k + 1)
            End If
        Else
            k = i
            If PurelyNumeric(Mid$(tmpStr2, k - 1, 1)) Or (Mid$(tmpStr2, k - 1, 1) = "x") Or (Mid$(tmpStr2, k - 1, 1) = ")") Then
                tmpStr2 = Left$(tmpStr2, k - 1) & "*" & Mid$(tmpStr2, k)
            ElseIf Mid$(tmpStr2, k - 1, 1) = "-" Then
                If k > 2 Then
                    If Mid$(tmpStr2, k - 2, 1) = "*" Or Mid$(tmpStr2, k - 2, 1) = "/" Then
                        tmpStr2 = Left$(tmpStr2, k - 2) & "(-1*" & Mid$(tmpStr2, k, j - k + 1) & ")" & Mid$(tmpStr2, j + 1)
                    Else
                        GoTo Continuation
                    End If
                Else
Continuation:       tmpStr2 = Left$(tmpStr2, k - 1) & "1*" & Mid$(tmpStr2, k)
                End If
            End If
        End If
    Else
        If i = 0 And j = 0 Then Exit Do
    End If
Loop

RedoOperations:

If Not PurelyNumeric(tmpStr2) Then  'Expression, not a number
    
    ' Refresh
    
    i = 0
    j = 0
    
SignFixProc:
    tmpStr3 = tmpStr2
    Do
        i = InStr(i + 1, tmpStr2, "--")
        If i = 1 Then
            tmpStr2 = Mid$(tmpStr2, 3)
        Else
            If i > 0 Then
                tmpStr2 = Left$(tmpStr2, i - 1) & "+" & Mid$(tmpStr2, i + 2)
            Else
                Exit Do
            End If
        End If
    Loop
    
    Do
        i = InStr(i + 1, tmpStr2, "+-")
        If i = 1 Then
            tmpStr2 = Mid$(tmpStr2, 2)
        Else
            If i > 0 Then
                tmpStr2 = Left$(tmpStr2, i - 1) & Mid$(tmpStr2, i + 1)
            Else
                Exit Do
            End If
        End If
    Loop
    
    Do
        i = InStr(i + 1, tmpStr2, "-+")
        If i = 1 Then
            tmpStr2 = "-" & Mid$(tmpStr2, 3)
        Else
            If i > 0 Then
                tmpStr2 = Left$(tmpStr2, i) & Mid$(tmpStr2, i + 2)
            Else
                Exit Do
            End If
        End If
    Loop
    
    tmpStr2 = Replace(tmpStr2, "++", "+")
    
    If tmpStr3 <> tmpStr2 Then GoTo SignFixProc
    
    'Refresh
    
    i = 0
    j = 0
    tmpStr3 = tmpStr2
    
    Set ColPar = Nothing
    Set ColPar = New Collection
    
    'Do order of operations
    
    Do
        j = InStr(j + 1, tmpStr3, ")")
        If j > 0 Then
            i = InStrRev(tmpStr3, "(", j - 1)
            If i = 0 Then
                EvaluateInput = "Syntax error"
                Exit Function
            Else
                tmpStr3 = Left$(tmpStr3, i - 1) & "<" & Mid$(tmpStr3, i + 1)
                tmpStr3 = Left$(tmpStr3, j - 1) & ">" & Mid$(tmpStr3, j + 1)
                ColPar.Add j, "P" & CStr(i)
            End If
        Else
            Exit Do
        End If
    Loop
    
    i = InStr(tmpStr2, "log(")
    If i > 0 Then
        j = ColPar("P" & CStr(i + 3))
        If j > 0 Then
            tmpExp = EvaluateInput(Mid$(tmpStr2, i + 4, j - i - 4))
            If PurelyNumeric(tmpExp) Then
                tmpStr2 = Left$(tmpStr2, i - 1) & CStr(Log10(CDbl(tmpExp))) & Mid$(tmpStr2, j + 1)
            Else
                GoTo ProcExit
            End If
        Else
            GoTo ProcExit
        End If
        GoTo RedoOperations
    End If
    
    i = InStr(tmpStr2, "ln(")
    If i > 0 Then
        j = ColPar("P" & CStr(i + 2))
        If j > 0 Then
            tmpExp = EvaluateInput(Mid$(tmpStr2, i + 3, j - i - 3))
            If PurelyNumeric(tmpExp) Then
                tmpStr2 = Left$(tmpStr2, i - 1) & CStr(Log(CDbl(tmpExp))) & Mid$(tmpStr2, j + 1)
            Else
                GoTo ProcExit
            End If
        Else
            GoTo ProcExit
        End If
        GoTo RedoOperations
    End If
    
    i = InStr(tmpStr2, "exp(")
    If i > 0 Then
        j = ColPar("P" & CStr(i + 3))
        If j > 0 Then
            tmpExp = EvaluateInput(Mid$(tmpStr2, i + 4, j - i - 4))
            If PurelyNumeric(tmpExp) Then
                tmpStr2 = Left$(tmpStr2, i - 1) & CStr(Exp(CDbl(tmpExp))) & Mid$(tmpStr2, j + 1)
            Else
                GoTo ProcExit
            End If
        Else
            GoTo ProcExit
        End If
        GoTo RedoOperations
    End If
    
    i = InStr(tmpStr2, "sin(")
    If i > 0 Then
        j = ColPar("P" & CStr(i + 3))
        If j > 0 Then
            tmpExp = EvaluateInput(Mid$(tmpStr2, i + 4, j - i - 4))
            If PurelyNumeric(tmpExp) Then
                Select Case lblDRG
                    Case "DEG": tmpNum = SinD(CDbl(tmpExp))
                    Case "RAD": tmpNum = Sin(CDbl(tmpExp))
                    Case "GRAD": tmpNum = SinG(CDbl(tmpExp))
                End Select
                tmpStr2 = Left$(tmpStr2, i - 1) & CStr(tmpNum) & Mid$(tmpStr2, j + 1)
            Else
                GoTo ProcExit
            End If
        Else
            GoTo ProcExit
        End If
        GoTo RedoOperations
    End If
    
    i = InStr(tmpStr2, "asn(")
    If i > 0 Then
        j = ColPar("P" & CStr(i + 3))
        If j > 0 Then
            tmpExp = EvaluateInput(Mid$(tmpStr2, i + 4, j - i - 4))
            If PurelyNumeric(tmpExp) Then
                Select Case lblDRG
                    Case "DEG": tmpNum = AsnD(CDbl(tmpExp))
                    Case "RAD": tmpNum = Asn(CDbl(tmpExp))
                    Case "GRAD": tmpNum = AsnG(CDbl(tmpExp))
                End Select
                tmpStr2 = Left$(tmpStr2, i - 1) & CStr(tmpNum) & Mid$(tmpStr2, j + 1)
            Else
                GoTo ProcExit
            End If
        Else
            GoTo ProcExit
        End If
        GoTo RedoOperations
    End If
    
    i = InStr(tmpStr2, "cos(")
    If i > 0 Then
        j = ColPar("P" & CStr(i + 3))
        If j > 0 Then
            tmpExp = EvaluateInput(Mid$(tmpStr2, i + 4, j - i - 4))
            If PurelyNumeric(tmpExp) Then
                Select Case lblDRG
                    Case "DEG": tmpNum = CosD(CDbl(tmpExp))
                    Case "RAD": tmpNum = Cos(CDbl(tmpExp))
                    Case "GRAD": tmpNum = CosG(CDbl(tmpExp))
                End Select
                tmpStr2 = Left$(tmpStr2, i - 1) & CStr(tmpNum) & Mid$(tmpStr2, j + 1)
            Else
                GoTo ProcExit
            End If
        Else
            GoTo ProcExit
        End If
        GoTo RedoOperations
    End If
    
    i = InStr(tmpStr2, "acs(")
    If i > 0 Then
        j = ColPar("P" & CStr(i + 3))
        If j > 0 Then
            tmpExp = EvaluateInput(Mid$(tmpStr2, i + 4, j - i - 4))
            If PurelyNumeric(tmpExp) Then
                Select Case lblDRG
                    Case "DEG": tmpNum = AcsD(CDbl(tmpExp))
                    Case "RAD": tmpNum = Acs(CDbl(tmpExp))
                    Case "GRAD": tmpNum = AcsG(CDbl(tmpExp))
                End Select
                tmpStr2 = Left$(tmpStr2, i - 1) & CStr(tmpNum) & Mid$(tmpStr2, j + 1)
            Else
                GoTo ProcExit
            End If
        Else
            GoTo ProcExit
        End If
        GoTo RedoOperations
    End If
    
    i = InStr(tmpStr2, "tan(")
    If i > 0 Then
        j = ColPar("P" & CStr(i + 3))
        If j > 0 Then
            tmpExp = EvaluateInput(Mid$(tmpStr2, i + 4, j - i - 4))
            If PurelyNumeric(tmpExp) Then
                Select Case lblDRG
                    Case "DEG": tmpNum = TanD(CDbl(tmpExp))
                    Case "RAD": tmpNum = Tan(CDbl(tmpExp))
                    Case "GRAD": tmpNum = TanG(CDbl(tmpExp))
                End Select
                tmpStr2 = Left$(tmpStr2, i - 1) & CStr(tmpNum) & Mid$(tmpStr2, j + 1)
            Else
                GoTo ProcExit
            End If
        Else
            GoTo ProcExit
        End If
        GoTo RedoOperations
    End If
    
    i = InStr(tmpStr2, "atn(")
    If i > 0 Then
        j = ColPar("P" & CStr(i + 3))
        If j > 0 Then
            tmpExp = EvaluateInput(Mid$(tmpStr2, i + 4, j - i - 4))
            If PurelyNumeric(tmpExp) Then
                Select Case lblDRG
                    Case "DEG": tmpNum = AtnD(CDbl(tmpExp))
                    Case "RAD": tmpNum = Atn(CDbl(tmpExp))
                    Case "GRAD": tmpNum = AtnG(CDbl(tmpExp))
                End Select
                tmpStr2 = Left$(tmpStr2, i - 1) & CStr(tmpNum) & Mid$(tmpStr2, j + 1)
            Else
                GoTo ProcExit
            End If
        Else
            GoTo ProcExit
        End If
        GoTo RedoOperations
    End If
    
    i = InStrRev(tmpStr2, "sqr(")
    If i > 0 Then
        j = ColPar("P" & CStr(i + 3))
        If j > 0 Then
            tmpExp = EvaluateInput(Mid$(tmpStr2, i + 4, j - i - 4))
            If PurelyNumeric(tmpExp) Then
                tmpStr2 = Left$(tmpStr2, i - 1) & CStr(Sqr(CDbl(tmpExp))) & Mid$(tmpStr2, j + 1)
            Else
                GoTo ProcExit
            End If
        Else
            GoTo ProcExit
        End If
        GoTo RedoOperations
    End If
    
    i = InStr(tmpStr2, "int(")
    If i > 0 Then
        j = ColPar("P" & CStr(i + 3))
        If j > 0 Then
            tmpExp = EvaluateInput(Mid$(tmpStr2, i + 4, j - i - 4))
            If PurelyNumeric(tmpExp) Then
                tmpStr2 = Left$(tmpStr2, i - 1) & CStr(Fix(CDbl(tmpExp))) & Mid$(tmpStr2, j + 1)
            Else
                GoTo ProcExit
            End If
        Else
            GoTo ProcExit
        End If
        GoTo RedoOperations
    End If
    
    i = InStr(tmpStr2, "abs(")
    If i > 0 Then
        j = ColPar("P" & CStr(i + 3))
        If j > 0 Then
            tmpExp = EvaluateInput(Mid$(tmpStr2, i + 4, j - i - 4))
            If PurelyNumeric(tmpExp) Then
                tmpStr2 = Left$(tmpStr2, i - 1) & CStr(Abs(CDbl(tmpExp))) & Mid$(tmpStr2, j + 1)
            Else
                GoTo ProcExit
            End If
        Else
            GoTo ProcExit
        End If
        GoTo RedoOperations
    End If
    
    i = 0
ProcParCheck:
    i = InStr(i + 1, tmpStr2, "(")
    If i > 0 Then
        j = ColPar("P" & CStr(i))
        If j > 0 Then
            tmpExp = Mid$(tmpStr2, i + 1, j - i - 1)
            If PurelyNumeric(tmpExp) Then
                If Mid$(tmpStr2, j + 1, 1) = "^" Then
                    GoTo ProcParCheck
                Else
                    tmpExp2 = EvaluateInput(Mid$(tmpStr2, i + 1, j - i - 1))
                    If PurelyNumeric(tmpExp2) Then
                        tmpStr2 = Left$(tmpStr2, i - 1) & tmpExp2 & Mid$(tmpStr2, j + 1)
                    Else
                        GoTo ProcExit
                    End If
                End If
            Else
                tmpExp2 = EvaluateInput(Mid$(tmpStr2, i + 1, j - i - 1))
                If PurelyNumeric(tmpExp2) Then
                    tmpStr2 = Left$(tmpStr2, i) & tmpExp2 & Mid$(tmpStr2, j)
                Else
                    GoTo ProcExit
                End If
            End If
        Else
            GoTo ProcExit
        End If
        GoTo RedoOperations
    End If
    
    i = InStrRev(tmpStr2, "^")
    If i > 0 Then
        j = RangeIt(tmpStr2, i, bPercent, bFactorial)
        k = RangeIt(tmpStr2, i, bPercent2, bFactorial2, True, bPar)
        
        tmpExp = Replace(Replace(Replace(Replace(Mid$(tmpStr2, k, i - k), "%", ""), "!", ""), "(", ""), ")", "")
        If PurelyNumeric(tmpExp) Then
            D1 = CDbl(tmpExp)
            If bPercent2 Then D1 = D1 / 100
            If bFactorial2 Then D1 = Factorial(D1)
        Else
            GoTo ProcExit
        End If
        
        tmpExp = Replace(Replace(Replace(Replace(Mid$(tmpStr2, i + 1, j - i), "%", ""), "!", ""), "(", ""), ")", "")
        If PurelyNumeric(tmpExp) Then
            D2 = CDbl(tmpExp)
            If bPercent Then D2 = D2 / 100
            If bFactorial Then D2 = Factorial(D2)
        Else
            GoTo ProcExit
        End If
        
        If bPar Then
            tmpNum = D1 ^ D2
        Else
            tmpNum = Abs(D1) ^ D2
            If IsNegative(D1) Then tmpNum = Opposite(tmpNum)
        End If
        
        tmpStr2 = Left$(tmpStr2, k - 1) & CStr(tmpNum) & Mid$(tmpStr2, j + 1)
        GoTo RedoOperations
    End If
    
    i = InStr(tmpStr2, "*")
    j = InStr(tmpStr2, "/")
    k = InStr(tmpStr2, "mod")
    
    If i > 0 And j = 0 And k = 0 Then GoTo MulProc
    If j > 0 And i = 0 And k = 0 Then GoTo DivProc
    If k > 0 And i = 0 And j = 0 Then GoTo ModProc
    
    If ((i > 0 And j > 0 And k = 0) And j > i) Or ((i > 0 And k > 0 And j = 0) And k > i) Or ((i > 0 And j > 0 And k > 0) And j > i And k > i) Then
MulProc:
        l = RangeIt(tmpStr2, i, bPercent, bFactorial)
        m = RangeIt(tmpStr2, i, bPercent2, bFactorial2, True)
        
        tmpExp = Replace(Replace(Mid$(tmpStr2, m, i - m), "%", ""), "!", "")
        If PurelyNumeric(tmpExp) Then
            D1 = CDbl(tmpExp)
            If bFactorial2 Then D1 = Factorial(D1)
            If bPercent2 Then D1 = D1 / 100
        Else
            GoTo ProcExit
        End If
        
        tmpExp = Replace(Replace(Mid$(tmpStr2, i + 1, l - i), "%", ""), "!", "")
        If PurelyNumeric(tmpExp) Then
            D2 = CDbl(tmpExp)
            If bFactorial Then D2 = Factorial(D2)
            If bPercent Then D2 = D2 / 100
        Else
            GoTo ProcExit
        End If
        
        tmpStr2 = Left$(tmpStr2, m - 1) & CStr(D1 * D2) & Mid$(tmpStr2, l + 1)
        GoTo RedoOperations
    ElseIf ((i > 0 And j > 0 And k = 0) And i > j) Or ((j > 0 And k > 0 And i = 0) And k > j) Or ((i > 0 And j > 0 And k > 0) And i > j And k > j) Then
DivProc:
        l = RangeIt(tmpStr2, j, bPercent, bFactorial)
        m = RangeIt(tmpStr2, j, bPercent2, bFactorial2, True)
        
        tmpExp = Replace(Replace(Mid$(tmpStr2, m, j - m), "%", ""), "!", "")
        If PurelyNumeric(tmpExp) Then
            D1 = CDbl(tmpExp)
            If bFactorial2 Then D1 = Factorial(D1)
            If bPercent2 Then D1 = D1 / 100
        Else
            GoTo ProcExit
        End If
        
        tmpExp = Replace(Replace(Mid$(tmpStr2, j + 1, l - j), "%", ""), "!", "")
        If PurelyNumeric(tmpExp) Then
            D2 = CDbl(tmpExp)
            If bFactorial Then D2 = Factorial(D2)
            If bPercent Then D2 = D2 / 100
        Else
            GoTo ProcExit
        End If
        
        tmpStr2 = Left$(tmpStr2, m - 1) & CStr(D1 / D2) & Mid$(tmpStr2, l + 1)
        GoTo RedoOperations
    ElseIf ((i > 0 And k > 0 And j = 0) And i > k) Or ((j > 0 And k > 0 And i = 0) And j > k) Or ((i > 0 And j > 0 And k > 0) And i > k And j > k) Then
ModProc:
        l = RangeIt(tmpStr2, k + 2, bPercent, bFactorial)
        m = RangeIt(tmpStr2, k, bPercent2, bFactorial2, True)
        
        tmpExp = Replace(Replace(Mid$(tmpStr2, m, k - m), "%", ""), "!", "")
        If PurelyNumeric(tmpExp) Then
            D1 = CDbl(tmpExp)
            If bFactorial2 Then D1 = Factorial(D1)
            If bPercent2 Then D1 = D1 / 100
        Else
            GoTo ProcExit
        End If
        
        tmpExp = Replace(Replace(Mid$(tmpStr2, k + 3, l - k - 2), "%", ""), "!", "")
        If PurelyNumeric(tmpExp) Then
            D2 = CDbl(tmpExp)
            If bFactorial Then D2 = Factorial(D2)
            If bPercent Then D2 = D2 / 100
        Else
            GoTo ProcExit
        End If
        
        tmpStr2 = Left$(tmpStr2, m - 1) & CStr(Modulo(D1, D2)) & Mid$(tmpStr2, l + 1)
        GoTo RedoOperations
    End If
    
    i = InStr(tmpStr2, "+")
    j = InStr(tmpStr2, "-")
    If j = 1 Then j = InStr(2, tmpStr2, "-")
    
AddSubProc:
    
    If i > 0 And j = 0 Then GoTo AddProc
    If j > 0 And i = 0 Then GoTo SubProc
    
    If (i > 0 Or j > 0) And j > i Then
AddProc:
        If Mid$(tmpStr2, i - 1, 1) <> "E" Then
            k = RangeIt(tmpStr2, i, bPercent, bFactorial)
            l = RangeIt(tmpStr2, i, bPercent2, bFactorial2, True)
            
            tmpExp = Replace(Replace(Mid$(tmpStr2, l, i - 1), "%", ""), "!", "")
            If PurelyNumeric(tmpExp) Then
                D1 = CDbl(tmpExp)
                If bFactorial2 Then D1 = Factorial(D1)
                If bPercent2 Then D1 = D1 / 100
            Else
                GoTo ProcExit
            End If
            
            tmpExp = Replace(Replace(Mid$(tmpStr2, i + 1, k - i), "%", ""), "!", "")
            If PurelyNumeric(tmpExp) Then
                D2 = CDbl(tmpExp)
                If bFactorial Then D2 = Factorial(D2)
                If bPercent Then D2 = D2 / 100
            Else
                GoTo ProcExit
            End If
            
            tmpStr2 = Left$(tmpStr2, l - 1) & CStr(D1 + D2) & Mid$(tmpStr2, k + 1)
            GoTo RedoOperations
        Else
            i = InStr(i + 1, tmpStr2, "-")
            GoTo AddSubProc
        End If
    ElseIf (i > 0 Or j > 0) And i > j Then
SubProc:
        If Mid$(tmpStr2, j - 1, 1) <> "E" Then
            k = RangeIt(tmpStr2, j, bPercent, bFactorial)
            l = RangeIt(tmpStr2, j, bPercent2, bFactorial2, True)
            
            tmpExp = Replace(Replace(Mid$(tmpStr2, l, j - 1), "%", ""), "!", "")
            If PurelyNumeric(tmpExp) Then
                D1 = CDbl(tmpExp)
                If bFactorial2 Then D1 = Factorial(D1)
                If bPercent2 Then D1 = D1 / 100
            Else
                GoTo ProcExit
            End If
            
            tmpExp = Replace(Replace(Mid$(tmpStr2, j + 1, k - j), "%", ""), "!", "")
            If PurelyNumeric(tmpExp) Then
                D2 = CDbl(tmpExp)
                If bFactorial Then D2 = Factorial(D2)
                If bPercent Then D2 = D2 / 100
            Else
                GoTo ProcExit
            End If
            
            tmpStr2 = Left$(tmpStr2, l - 1) & CStr(D1 - D2) & Mid$(tmpStr2, k + 1)
            GoTo RedoOperations
        Else
            j = InStr(j + 1, tmpStr2, "-")
            GoTo AddSubProc
        End If
    End If
    
    bPercent = (Right$(Replace(tmpStr2, "!", ""), 1) = "%")
    bFactorial = (Right$(Replace(tmpStr2, "%", ""), 1) = "!")
    
    If bFactorial Then tmpStr2 = CStr(Factorial(CDbl(Replace(Replace(tmpStr2, "%", ""), "!", ""))))
    If bPercent Then tmpStr2 = CStr(CDbl(Replace(Replace(tmpStr2, "%", ""), "!", "")) / 100)
Else
    tmpStr2 = CStr(CDbl(tmpStr2))
End If

If Not PurelyNumeric(tmpStr2) Then
    EvaluateInput = "Error"
    Exit Function
End If

i = InStr(tmpStr, ">roundto(")
If i > 0 Then
    j = InStr(i + 1, tmpStr, ")")
    If j > 0 Then
        tmpExp = Mid$(tmpStr, i + 9, j - i - 9)
    Else
        tmpExp = Mid$(tmpStr, i + 9)
    End If
    
    If PurelyNumeric(tmpExp) Then
        tmpStr2 = MyRound(CDbl(tmpStr2), CLng(tmpExp))
    Else
        GoTo ProcExit
    End If
End If

i = InStr(tmpStr, ">frac")
If i > 0 Then tmpStr2 = GetFraction(CDbl(tmpStr2))

EvaluateInput = tmpStr2

Exit Function
ProcExit:
    EvaluateInput = "Error"
End Function

Private Sub AddTextToInput(ByVal Text As String)
txtInput = txtInput & Text
End Sub

Private Sub FocusAndSetCursor()
txtInput.SetFocus
txtInput.SelStart = Len(txtInput)
End Sub

Private Sub AddTextAndFocus(ByVal Text As String)
AddTextToInput Text
FocusAndSetCursor
End Sub

Private Sub cmdAcs_Click()
AddTextAndFocus "Acs("
End Sub

Private Sub cmdAns_Click()
AddTextAndFocus "Ans"
End Sub

Private Sub cmdAsn_Click()
AddTextAndFocus "Asn("
End Sub

Private Sub cmdAtn_Click()
AddTextAndFocus "Atn("
End Sub

Private Sub cmdBackspace_Click()
On Error Resume Next
If txtInput <> "" Then txtInput = Left$(txtInput, Len(txtInput) - 1)
FocusAndSetCursor
End Sub

Private Sub cmdChange_Click()
Select Case lblDRG
    Case "DEG": lblDRG = "RAD"
    Case "RAD": lblDRG = "GRAD"
    Case "GRAD": lblDRG = "DEG"
End Select
FocusAndSetCursor
End Sub

Private Sub cmdClear_Click()
txtInput = ""
txtOutput = ""
lblDRG = "DEG"
FocusAndSetCursor
End Sub

Private Sub cmdClosedPar_Click()
AddTextAndFocus ")"
End Sub

Private Sub cmdCos_Click()
AddTextAndFocus "Cos("
End Sub

Private Sub cmdDecPoint_Click()
AddTextAndFocus "."
End Sub

Private Sub cmdDividedBy_Click()
AddTextAndFocus "/"
End Sub

Private Sub cmdENeg_Click()
AddTextAndFocus "E-"
End Sub

Private Sub cmdEPlus_Click()
AddTextAndFocus "E+"
End Sub

Private Sub cmdEquals_Click()
If txtInput <> "" Then
    MousePointer = vbHourglass: DoEvents
    txtOutput = EvaluateInput(txtInput)
    MousePointer = vbNormal: DoEvents
End If
FocusAndSetCursor
End Sub

Private Sub cmdExp_Click()
AddTextAndFocus "Exp("
End Sub

Private Sub cmdFactorial_Click()
AddTextAndFocus "!"
End Sub

Private Sub cmdGetBase_Click()
AddTextAndFocus "^(1/("
End Sub

Private Sub cmdInt_Click()
AddTextAndFocus "Int("
End Sub

Private Sub cmdLn_Click()
AddTextAndFocus "Ln("
End Sub

Private Sub cmdLog_Click()
AddTextAndFocus "Log("
End Sub

Private Sub cmdMinusOrNeg_Click()
AddTextAndFocus "-"
End Sub

Private Sub cmdMod_Click()
AddTextAndFocus " Mod "
End Sub

Private Sub cmdNeg1Power_Click()
AddTextAndFocus "^-1"
End Sub

Private Sub cmdNum0_Click()
AddTextAndFocus "0"
End Sub

Private Sub cmdNum1_Click()
AddTextAndFocus "1"
End Sub

Private Sub cmdNum2_Click()
AddTextAndFocus "2"
End Sub

Private Sub cmdNum3_Click()
AddTextAndFocus "3"
End Sub

Private Sub cmdNum4_Click()
AddTextAndFocus "4"
End Sub

Private Sub cmdNum5_Click()
AddTextAndFocus "5"
End Sub

Private Sub cmdNum6_Click()
AddTextAndFocus "6"
End Sub

Private Sub cmdNum7_Click()
AddTextAndFocus "7"
End Sub

Private Sub cmdNum8_Click()
AddTextAndFocus "8"
End Sub

Private Sub cmdNum9_Click()
AddTextAndFocus "9"
End Sub

Private Sub cmdOpenPar_Click()
AddTextAndFocus "("
End Sub

Private Sub cmdPercent_Click()
AddTextAndFocus "%"
End Sub

Private Sub cmdPi_Click()
AddTextAndFocus "Pi"
End Sub

Private Sub cmdPlus_Click()
AddTextAndFocus "+"
End Sub

Private Sub cmdRecip_Click()
AddTextAndFocus "1/("
End Sub

Private Sub cmdRound_Click()
AddTextAndFocus ">RoundTo("
End Sub

Private Sub cmdSin_Click()
AddTextAndFocus "Sin("
End Sub

Private Sub cmdSqr_Click()
AddTextAndFocus "Sqr("
End Sub

Private Sub cmdSquared_Click()
AddTextAndFocus "^2"
End Sub

Private Sub cmdTan_Click()
AddTextAndFocus "Tan("
End Sub

Private Sub cmdTimes_Click()
AddTextAndFocus "*"
End Sub

Private Sub cmdToFraction_Click()
AddTextAndFocus ">Frac"
End Sub

Private Sub cmdXtoYpower_Click()
AddTextAndFocus "^("
End Sub

Private Sub txtInput_KeyPress(KeyAscii As Integer)
If txtInput <> "" Then
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmdEquals_Click
    End If
End If
End Sub
