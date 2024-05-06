Attribute VB_Name = "Module1"
Option Explicit
Global Const pi = 3.14159265358979
'Time are represented by date, and the unit for the period of time are years follows ACT365
'The unit for vol are in %, only in the final process in sigma_P would turns it into 100%

Function Dscnt(t As Long, curve As Range) As Double

    Dim idx As Integer: idx = 1
    While t > curve(idx, 3) And idx < 28
        idx = idx + 1
    Wend
    If idx = 1 Then
        Dscnt = curve(1, 5)
    ElseIf idx = 28 Then
        Dscnt = Exp(-curve(27, 6) / 100 * (t - curve(1, 1)) / 365)
    Else: Dscnt = Exp(-(curve(idx - 1, 6) * (curve(idx, 3) - t) + curve(idx, 6) * (t - curve(idx - 1, 3))) / (curve(idx, 3) - curve(idx - 1, 3)) / 100 * (t - curve(1, 1)) / 365)
    End If
    
    If t = curve(1, 1) Then
        Dscnt = 1
    End If
    
End Function

Function Hull_White_Value(sigma As Double, a As Double, curve As Range, SwapTerm As Integer, OptionTenorCase As String) As Double
    
    Dim freq As Integer: freq = 1
    Dim flows As Integer: flows = freq * SwapTerm
    Dim t() As Long, D() As Double
    ReDim t(0 To flows), D(0 To flows)
    If Right(OptionTenorCase, 1) = "m" Then
        t(0) = CInt(Left(OptionTenorCase, Len(OptionTenorCase) - 1)) * CInt(365 / 12) + curve(1, 1)
        Else: t(0) = CInt(Left(OptionTenorCase, Len(OptionTenorCase) - 1)) * 365 + curve(1, 1)
    End If
    
    Dim i As Integer
    For i = 1 To flows
        t(i) = t(0) + CInt(365 / freq * i)
    Next i
    
    For i = 0 To flows
        D(i) = Dscnt(t(i), curve)
    Next i
    
    Dim s0 As Double: s0 = FwdRate(curve, SwapTerm, OptionTenorCase)
    Dim sK As Double: sK = s0 ' At The Money
    
    Dim R As Double: R = zero_swap(t(0), t(flows), a, sigma, s0, curve, flows) 'Jamshidian's procedure
    
    Dim sigmaP As Double: sigmaP = sigma_P(sigma, a, t(0), t(flows), curve)
    Dim pp  As Double: pp = P(t(0), t(flows), a, sigma, R, curve)
    Dim hh As Double: hh = h(sigmaP, t(0), t(flows), curve, pp)
    Hull_White_Value = D(flows) * N(hh) - P(t(0), t(flows), a, sigma, R, curve) * D(0) * N(hh - sigmaP)
    
    For i = 1 To flows
        sigmaP = sigma_P(sigma, a, t(0), t(i), curve)
        hh = h(sigmaP, t(0), t(i), curve, P(t(0), t(i), a, sigma, R, curve))
        Hull_White_Value = Hull_White_Value + (D(i) * N(hh) - P(t(0), t(i), a, sigma, R, curve) * D(0) * N(hh - sigmaP)) * s0 * (t(i) - t(i - 1))
    Next i
    
End Function

Function Normal_Value(sigma As Double, curve As Range, SwapTerm As Integer, OptionTenorCase As String) As Double

    Dim freq As Integer: freq = 1
    Dim flows As Integer: flows = freq * SwapTerm
    Dim t() As Long, D() As Double
    ReDim t(0 To flows), D(0 To flows)
    If Right(OptionTenorCase, 1) = "m" Then
        t(0) = CInt(Left(OptionTenorCase, Len(OptionTenorCase) - 1)) * CInt(365 / 12) + curve(1, 1)
        Else: t(0) = CInt(Left(OptionTenorCase, Len(OptionTenorCase) - 1)) * 365 + curve(1, 1)
    End If
    
    Dim i As Integer
    For i = 1 To flows
        t(i) = t(i - 1) + CInt(365 / freq)
    Next i
    
    For i = 0 To flows
        D(i) = Dscnt(t(i), curve)
    Next i
    
    For i = 1 To flows
        Normal_Value = Normal_Value + D(i)
    Next i
    
    Normal_Value = sigma / 100 / freq * ((t(0) - curve(1, 1)) / 365 / 2 / pi) ^ 0.5 * Normal_Value

End Function

Function F(t As Long, curve As Range) As Double
    F = Log(Dscnt(t, curve) / Dscnt(t + 1, curve)) * 365
End Function

Function F_T(t As Long, curve As Range) As Double
    F_T = (F(t + 1, curve) - F(t, curve)) * 365
End Function

Function N(x As Double) As Double
    N = WorksheetFunction.Norm_S_Dist(x, True)
End Function

Function B_bond(t As Long, s As Long, a As Double) As Double
    B_bond = (1 - Exp(-a * (s - t) / 365)) / a
End Function

Function LOG_A_bond(t As Long, s As Long, a As Double, sigma As Double, curve As Range) As Double
    LOG_A_bond = Log(Dscnt(s, curve) / Dscnt(t, curve)) + B_bond(t, s, a) * F(t, curve) - (sigma / 100) ^ 2 * (Exp(-a * (s - curve(1, 1)) / 365) - Exp(-a * (t - curve(1, 1)) / 365)) * (Exp(2 * a * (t - curve(1, 1)) / 365) - 1) / (4 * a ^ 3)
End Function

Function P(t As Long, s As Long, a As Double, sigma As Double, R As Double, curve As Range) As Double
    P = Exp(LOG_A_bond(t, s, a, sigma, curve) - B_bond(t, s, a) * R)
End Function

Function sigma_P(sigma As Double, a As Double, t0 As Long, te As Long, curve As Range) As Double
    
    sigma_P = sigma / 100 / a * (1 - Exp(-a * (te - t0) / 365)) * ((1 - Exp(-2 * a * CDbl(t0 - curve(1, 1)) / 365)) / 2 / a) ^ 0.5

End Function

Function h(sigmaP As Double, t0 As Long, te As Long, curve As Range, k As Double) As Double

    h = Log(Dscnt(te, curve) / Dscnt(t0, curve) / k) / sigmaP + sigmaP / 2
    
End Function

Function swap_value(t0 As Long, te As Long, a As Double, sigma As Double, swaprate As Double, curve As Range, Nflow As Integer, R As Double) As Double
    Dim i As Integer
    swap_value = 0
    Dim temp  As Double: temp = 0
    For i = 1 To Nflow
        temp = temp + P(t0, t0 + CLng(i * (te - t0) / Nflow), a, sigma, R, curve)
    Next i
    swap_value = 1 - P(t0, te, a, sigma, R, curve) - swaprate * temp * (te - t0) / Nflow / 365
End Function

Function zero_swap(t As Long, s As Long, a As Double, sigma As Double, swaprate As Double, curve As Range, Nflow As Integer) As Double
    Dim i As Integer: i = 0
    Dim R As Double: R = 1
    For i = 0 To 20
        If swap_value(t, s, a, sigma, swaprate, curve, Nflow, R) > 0 Then
            R = R - 0.5 ^ i
        Else: R = R + 0.5 ^ i
        End If
    Next i
    zero_swap = R
End Function

Function FwdRate(curve As Range, SwapTerm As Integer, OptionTenorCase As String) As Double

    Dim freq As Integer: freq = 1
    Dim t() As Long, D() As Double
    ReDim t(0 To freq * SwapTerm), D(0 To freq * SwapTerm)
    If Right(OptionTenorCase, 1) = "m" Then
        t(0) = CInt(Left(OptionTenorCase, Len(OptionTenorCase) - 1)) * CInt(365 / 12) + curve(1, 1)
        Else: t(0) = CInt(Left(OptionTenorCase, Len(OptionTenorCase) - 1)) * 365 + curve(1, 1)
    End If
    
    Dim i As Integer
    For i = 1 To freq * SwapTerm
        t(i) = t(i - 1) + CInt(365 / freq)
    Next i
        
    For i = 0 To freq * SwapTerm
        D(i) = Dscnt(t(i), curve)
    Next i
    
    Dim s0 As Double: s0 = 0
    For i = 1 To freq * SwapTerm
        s0 = s0 + (t(i) - t(i - 1)) / 365 * D(i)
    Next i
    s0 = (D(0) - D(freq * SwapTerm)) / s0
    
    FwdRate = s0
       
End Function


Sub calibration()

        SolverAdd CellRef:="$Q$17", Relation:=3, FormulaText:="0.01"
        SolverAdd CellRef:="$Q$18", Relation:=3, FormulaText:="0.01"
        SolverOk SetCell:="$Q$25", MaxMinVal:=2, ValueOf:=0, ByChange:="$Q$17:$Q$18", Engine:=1, EngineDesc:="GRG Nonlinear"
        SolverSolve UserFinish:=True
        SolverFinish KeepFinal:=1
        SolverReset
        
End Sub
