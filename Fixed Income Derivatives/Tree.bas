Attribute VB_Name = "Tree"
Option Explicit
Sub Tree_Calc()

    Dim Nb As Integer: Nb = Worksheets("TREE").Range("F5")    'N_branchs_between_flows
    Dim a As Double: a = Worksheets("TREE").Range("C5")
    Dim sigma As Double: sigma = Worksheets("TREE").Range("C6")
    Dim swap_tenor As Double: swap_tenor = Worksheets("TREE").Range("F4")
    Dim Delta_t As Double: Delta_t = swap_tenor / Nb
    Dim curve As Range: Set curve = Worksheets("DATA").Range("C18:H44")
    Dim s_K As Double: s_K = Worksheets("TREE").Range("F6") ' pay fixed rate

    Dim Delta_R As Double: Delta_R = sigma * (3 * Delta_t) ^ 0.5
    Dim M As Integer: M = -WorksheetFunction.Floor_Math(-0.184 / a / Delta_t)   'MaxMove
    Dim N_options As Integer: N_options = Worksheets("TREE").Range("F7") ' Say "1*5 / 2*4 / 3*3 / 4*2 / 5*1" is 5
    Dim N_trees As Integer: N_trees = Nb * (N_options + 1) + 1 'Total number of nodes
    
    Dim Prob() As Double, i As Integer, j As Integer, k As Integer
    ReDim Prob(-M To M, -1 To 1)    'Pd, Pm, Pu probabilities of going down, middle, and down
    For j = 1 - M To M - 1
        Prob(j, 1) = 1 / 6 + ((a * j * Delta_t) ^ 2 - a * j * Delta_t) / 2
        Prob(j, 0) = 2 / 3 - (a * j * Delta_t) ^ 2
        Prob(j, -1) = 1 / 6 + ((a * j * Delta_t) ^ 2 + a * j * Delta_t) / 2
    Next j
    
    Prob(-M, 1) = 1 / 6 + ((a * M * Delta_t) ^ 2 - a * M * Delta_t) / 2 'Nonstandard Branching
    Prob(-M, 0) = -1 / 3 - (a * M * Delta_t) ^ 2 + 2 * a * M * Delta_t
    Prob(-M, -1) = 7 / 6 + ((a * M * Delta_t) ^ 2 - 3 * a * M * Delta_t) / 2
    Prob(M, 1) = Prob(-M, -1)
    Prob(M, 0) = Prob(-M, 0)
    Prob(M, -1) = Prob(-M, 1)
    
    Dim t() As Long, D() As Double: ReDim t(0 To N_trees + 1), D(0 To N_trees + 1) 'Date and their discount
    
    For i = 0 To N_trees + 1
        t(i) = curve(1, 1) + WorksheetFunction.Floor_Math(365.25 * Delta_t * i) 'Real calendar is not considered
        D(i) = Dscnt(t(i), curve)
    Next i
        
    Dim Q() As Double, alpha() As Double: ReDim Q(0 To N_trees, -M To M), alpha(0 To N_trees)
    
    'Q(i,j) is present value of a security that pays off $1 if node (i, j) is reached and zero otherwise.
    'alpha(t) = R(t)-R*(t)
    
    alpha(0) = -Log(D(1)) / Delta_t
    Q(0, 0) = 1
    If M > 2 Then
        For i = 1 To N_trees
            For j = 3 - M To M - 3
                For k = -1 To 1
                    Q(i, j) = Q(i, j) + Q(i - 1, j + k) * Prob(j + k, -k) * Exp(-(alpha(i - 1) + (j + k) * Delta_R) * Delta_t)
                Next k
            Next j
            Q(i, M) = Q(i - 1, M) * Prob(M, 1) * Exp(-(alpha(i - 1) + M * Delta_R) * Delta_t) _
                                     + Q(i - 1, M - 1) * Prob(M - 1, 1) * Exp(-(alpha(i - 1) + (M - 1) * Delta_R) * Delta_t)
            Q(i, M - 1) = Q(i - 1, M) * Prob(M, 0) * Exp(-(alpha(i - 1) + M * Delta_R) * Delta_t) _
                                           + Q(i - 1, M - 1) * Prob(M - 1, 0) * Exp(-(alpha(i - 1) + (M - 1) * Delta_R) * Delta_t) _
                                           + Q(i - 1, M - 2) * Prob(M - 2, 1) * Exp(-(alpha(i - 1) + (M - 2) * Delta_R) * Delta_t)
            Q(i, M - 2) = Q(i - 1, M) * Prob(M, -1) * Exp(-(alpha(i - 1) + M * Delta_R) * Delta_t) _
                                           + Q(i - 1, M - 1) * Prob(M - 1, -1) * Exp(-(alpha(i - 1) + (M - 1) * Delta_R) * Delta_t) _
                                           + Q(i - 1, M - 2) * Prob(M - 2, 0) * Exp(-(alpha(i - 1) + (M - 2) * Delta_R) * Delta_t) _
                                           + Q(i - 1, M - 3) * Prob(M - 3, 1) * Exp(-(alpha(i - 1) + (M - 3) * Delta_R) * Delta_t)
            Q(i, 2 - M) = Q(i - 1, -M) * Prob(-M, 1) * Exp(-(alpha(i - 1) - M * Delta_R) * Delta_t) _
                                           + Q(i - 1, -M + 1) * Prob(-M + 1, 1) * Exp(-(alpha(i - 1) - (M - 1) * Delta_R) * Delta_t) _
                                           + Q(i - 1, -M + 2) * Prob(-M + 2, 0) * Exp(-(alpha(i - 1) - (M - 2) * Delta_R) * Delta_t) _
                                           + Q(i - 1, -M + 3) * Prob(-M + 3, -1) * Exp(-(alpha(i - 1) - (M - 3) * Delta_R) * Delta_t)
            Q(i, 1 - M) = Q(i - 1, -M) * Prob(-M, 0) * Exp(-(alpha(i - 1) - M * Delta_R) * Delta_t) _
                                           + Q(i - 1, -M + 1) * Prob(-M + 1, 0) * Exp(-(alpha(i - 1) - (M - 1) * Delta_R) * Delta_t) _
                                           + Q(i - 1, -M + 2) * Prob(-M + 2, -1) * Exp(-(alpha(i - 1) - (M - 2) * Delta_R) * Delta_t)
            Q(i, -M) = Q(i - 1, -M) * Prob(-M, -1) * Exp(-(alpha(i - 1) - M * Delta_R) * Delta_t) _
                                       + Q(i - 1, -M + 1) * Prob(-M + 1, -1) * Exp(-(alpha(i - 1) - (M - 1) * Delta_R) * Delta_t)
            alpha(i) = 0
            For j = -M To M
                alpha(i) = alpha(i) + Q(i, j) * Exp(-j * Delta_R * Delta_t)
            Next j
            alpha(i) = Log(alpha(i) / D(i + 1)) / Delta_t
        Next i
    End If
    
    If M = 1 Then
        For i = 1 To N_trees
            Q(i, 1) = Q(i - 1, 1) * Prob(1, 1) * Exp(-(alpha(i - 1) + Delta_R) * Delta_t) _
                       + Q(i - 1, 0) * Prob(0, 1) * Exp(-(alpha(i - 1)) * Delta_t) _
                       + Q(i - 1, -1) * Prob(-1, 1) * Exp(-(alpha(i - 1) - Delta_R) * Delta_t)
            Q(i, 0) = Q(i - 1, 1) * Prob(1, 0) * Exp(-(alpha(i - 1) + Delta_R) * Delta_t) _
                       + Q(i - 1, 0) * Prob(0, 0) * Exp(-alpha(i - 1) * Delta_t) _
                       + Q(i - 1, -1) * Prob(-1, 0) * Exp(-(alpha(i - 1) - Delta_R) * Delta_t)
            Q(i, -1) = Q(i - 1, 1) * Prob(1, -1) * Exp(-(alpha(i - 1) + Delta_R) * Delta_t) _
                         + Q(i - 1, 0) * Prob(0, -1) * Exp(-alpha(i - 1) * Delta_t) _
                         + Q(i - 1, -1) * Prob(-1, -1) * Exp(-(alpha(i - 1) - Delta_R) * Delta_t)
            alpha(i) = 0
            For j = -1 To 1
                alpha(i) = alpha(i) + Q(i, j) * Exp(-j * Delta_R * Delta_t)
            Next j
            alpha(i) = Log(alpha(i) / D(i + 1)) / Delta_t
        Next i
    End If
    
    If M = 2 Then
        For i = 1 To N_trees
            Q(i, 2) = Q(i - 1, 2) * Prob(2, 1) * Exp(-(alpha(i - 1) + 2 * Delta_R) * Delta_t) _
                       + Q(i - 1, 1) * Prob(1, 1) * Exp(-(alpha(i - 1) + 1 * Delta_R) * Delta_t)
            Q(i, 1) = Q(i - 1, 2) * Prob(2, 0) * Exp(-(alpha(i - 1) + 2 * Delta_R) * Delta_t) _
                       + Q(i - 1, 1) * Prob(1, 0) * Exp(-(alpha(i - 1) + 1 * Delta_R) * Delta_t) _
                       + Q(i - 1, 0) * Prob(0, 1) * Exp(-(alpha(i - 1) + 0 * Delta_R) * Delta_t)
            Q(i, 0) = Q(i - 1, 2) * Prob(2, -1) * Exp(-(alpha(i - 1) + 2 * Delta_R) * Delta_t) _
                       + Q(i - 1, 1) * Prob(1, -1) * Exp(-(alpha(i - 1) + 1 * Delta_R) * Delta_t) _
                       + Q(i - 1, 0) * Prob(0, 0) * Exp(-(alpha(i - 1)) * Delta_t) _
                       + Q(i - 1, -1) * Prob(-1, 1) * Exp(-(alpha(i - 1) - Delta_R) * Delta_t) _
                       + Q(i - 1, -2) * Prob(-2, 1) * Exp(-(alpha(i - 1) - 2 * Delta_R) * Delta_t)
            Q(i, -1) = Q(i - 1, -2) * Prob(-2, 0) * Exp(-(alpha(i - 1) - 2 * Delta_R) * Delta_t) _
                         + Q(i - 1, -1) * Prob(-1, 0) * Exp(-(alpha(i - 1) - Delta_R) * Delta_t) _
                         + Q(i - 1, 0) * Prob(0, -1) * Exp(-(alpha(i - 1)) * Delta_t)
            Q(i, -2) = Q(i - 1, -2) * Prob(-2, -1) * Exp(-(alpha(i - 1) - 2 * Delta_R) * Delta_t) _
                         + Q(i - 1, -1) * Prob(-1, -1) * Exp(-(alpha(i - 1) - Delta_R) * Delta_t)
            alpha(i) = 0
            For j = -2 To 2
                alpha(i) = alpha(i) + Q(i, j) * Exp(-j * Delta_R * Delta_t)
            Next j
            alpha(i) = Log(alpha(i) / D(i + 1)) / Delta_t
        Next i
    End If
    
    Dim R() As Double: ReDim R(0 To N_trees, -M To M) 'delta_t period rate
    For i = 0 To N_trees
        For j = -WorksheetFunction.Min(M, i) To WorksheetFunction.Min(M, i)
            R(i, j) = alpha(i) + j * Delta_R
        Next j
    Next i
    
    Dim B() As Double: ReDim B(0 To N_trees, 0 To N_trees) 'Bond value P(t,T) = A(t,T)*exp(B(t,T)*r(t)) using instantanius rate r
    Dim B_() As Double: ReDim B_(0 To N_trees, 0 To N_trees) 'Bond_AB(t,T) = A_(t,T)*exp(B_(t,T)*R(t)) using delta_t period rate R
    Dim logA_() As Double: ReDim logA_(0 To N_trees, 0 To N_trees)
    For i = 0 To N_trees
        For j = i + 1 To N_trees
            B(i, j) = (1 - Exp(-a * (j - i) * Delta_t)) / a
        Next j
    Next i
    For i = 0 To N_trees
        For j = i + 1 To N_trees
            B_(i, j) = B(i, j) / B(i, i + 1) * Delta_t
            logA_(i, j) = Log(D(j) / D(i)) - B(i, j) / B(i, i + 1) * Log(D(i + 1) / D(i)) - sigma ^ 2 / 4 / a * (1 - Exp(-2 * a * i * Delta_t)) * B(i, j) * (B(i, j) - B(i, i + 1))
        Next j
    Next i
    
    '======================================================================================================
    '======================================================================================================
    
    'Dim Bond() As Double: ReDim Bond(0 To N_options + 1, 0 To (N_options + 1) * Nb, -M To M)
    'For k = 0 To N_options + 1 'The Bond pays 1 dollar at time k * Nb * delta_t
    '    For i = 0 To k ' The value of this bond at time i * Nb * delta_t
    '        For j = -WorksheetFunction.Min(M, Nb * i) To WorksheetFunction.Min(M, Nb * i)
    '            Bond(k, k * Nb, j) = 1
    '        Next j
    '    Next i
    'Next k
    
    'For k = 1 To N_options + 1
    '    For i = 1 To k * Nb
    '        Bond(k, k * Nb - i, M) = (Bond(k, k * Nb - i + 1, M) * Prob(M, 1) + Bond(k, k * Nb - i + 1, M - 1) * Prob(M, 0) + Bond(k, k * Nb - i + 1, M - 2) * Prob(M, -1)) * Exp(-R(k * Nb - i, M) / Nb)
    '        Bond(k, k * Nb - i, -M) = (Bond(k, k * Nb - i + 1, -M) * Prob(-M, -1) + Bond(k, k * Nb - i + 1, 1 - M) * Prob(-M, 0) + Bond(k, k * Nb - i + 1, 2 - M) * Prob(-M, 1)) * Exp(-R(k * Nb - i, -M) / Nb)
    '        If M > (N_options + 1) * Nb - i Then
    '            Bond(k, k * Nb - i, M) = 0
    '            Bond(k, k * Nb - i, -M) = 0
    '        End If
    '        For j = 1 - M To M - 1
    '            Bond(k, k * Nb - i, j) = (Bond(k, k * Nb - i + 1, j + 1) * Prob(j, 1) + Bond(k, k * Nb - i + 1, j) * Prob(j, 0) + Bond(k, k * Nb - i + 1, j - 1) * Prob(j, -1)) * Exp(-R(k * Nb - i, j) / Nb)
    '            If j > k * Nb - i Then
    '                 Bond(k, k * Nb - i, j) = 0
    '                Bond(k, k * Nb - i, -j) = 0
    '             End If
    '        Next j
    '    Next i
    'Next k
    
    '======================================================================================================
    '======================================================================================================
    
    Dim Bond_AB() As Double: ReDim Bond_AB(0 To N_options + 1, 0 To (N_options + 1) * Nb, -M To M)
    For k = 0 To N_options + 1 'The Bond pays 1 dollar at time k * Nb * delta_t
        For i = 0 To k * Nb ' The value of this bond at time i  * delta_t
            For j = -WorksheetFunction.Min(M, i) To WorksheetFunction.Min(M, i)
                Bond_AB(k, i, j) = Exp(logA_(i, k * Nb) - B_(i, k * Nb) * R(i, j))
            Next j
        Next i
    Next k
    
    Dim V_swap() As Double: ReDim V_swap(0 To N_options, -M To M) ' Swap Value
    For i = 0 To N_options
        For j = -WorksheetFunction.Min(M, Nb * i) To WorksheetFunction.Min(M, Nb * i)
            V_swap(i, j) = 0
            For k = i + 1 To N_options + 1
                V_swap(i, j) = V_swap(i, j) + Bond_AB(k, i * Nb, j)
            Next k
            V_swap(i, j) = 1 - Bond_AB(N_options + 1, i * Nb, j) - s_K * swap_tenor * V_swap(i, j)
        Next j
    Next i

    Dim V_swaption() As Double: ReDim V_swaption(0 To N_options * Nb, -M To M) 'Calculate swaption value
    For i = 1 To N_options
        For j = -WorksheetFunction.Min(M, Nb * N_options) To WorksheetFunction.Min(M, Nb * N_options)
            V_swaption(i * Nb, j) = WorksheetFunction.Max(V_swap(i, j), 0)
        Next j
    Next i
    
    For i = 1 To N_options * Nb
        V_swaption(N_options * Nb - i, M) = WorksheetFunction.Max(V_swaption(N_options * Nb - i, M), (V_swaption(N_options * Nb - i + 1, M) * Prob(M, 1) + V_swaption(N_options * Nb - i + 1, M - 1) * Prob(M, 0) + V_swaption(N_options * Nb - i + 1, M - 2) * Prob(M, -1)) * Exp(-R(N_options * Nb - i, M) / Nb))
        V_swaption(N_options * Nb - i, -M) = WorksheetFunction.Max(V_swaption(N_options * Nb - i, -M), (V_swaption(N_options * Nb - i + 1, -M) * Prob(-M, -1) + V_swaption(N_options * Nb - i + 1, 1 - M) * Prob(-M, 0) + V_swaption(N_options * Nb - i + 1, 2 - M) * Prob(-M, 1)) * Exp(-R(N_options * Nb - i, -M) / Nb))
        If M > N_options * Nb - i Then
            V_swaption(N_options * Nb - i, M) = 0
            V_swaption(N_options * Nb - i, -M) = 0
        End If
        For j = 1 - M To M - 1
            V_swaption(N_options * Nb - i, j) = WorksheetFunction.Max(V_swaption(N_options * Nb - i, j), (V_swaption(N_options * Nb - i + 1, j + 1) * Prob(j, 1) + V_swaption(N_options * Nb - i + 1, j) * Prob(j, 0) + V_swaption(N_options * Nb - i + 1, j - 1) * Prob(j, -1)) * Exp(-R(N_options * Nb - i, j) / Nb))
            If j > N_options * Nb - i Then
                V_swaption(N_options * Nb - i, j) = 0
                V_swaption(N_options * Nb - i, -j) = 0
            End If
        Next j
    Next i
    
    Worksheets("TREE").Range("I5") = V_swaption(0, 0) * 1000 'Swaption Value (Notional = 1000)
    
    Worksheets("TREE").Range("B16: BBB1123").Clear
    
    If Worksheets("TREE").Range("C15") = "YES" Then 'Print details or not
        For i = 0 To N_trees
        For j = -WorksheetFunction.Min(M, i) To WorksheetFunction.Min(M, i)
            Worksheets("TREE").Range("B16").Offset(M - j, i) = R(i, j)
            Worksheets("TREE").Range("B16").Offset(M - j, i).Font.Name = "Arial"
            If i Mod Nb = 0 Then: Worksheets("TREE").Range("B16").Offset(M - j, i).Interior.ColorIndex = 8
        Next j
        Next i
        Worksheets("TREE").Range("B16").Offset(M - 2, 0) = "Tree for R"
        Worksheets("TREE").Range("B16").Offset(M - 2, 0).Font.Name = "Arial"
    End If
End Sub


Sub Print_Details()
    If Worksheets("TREE").Range("C15") = "YES" Then: Worksheets("TREE").Range("C15") = "NO": Else Worksheets("TREE").Range("C15") = "YES"
End Sub

