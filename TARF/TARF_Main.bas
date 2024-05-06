Attribute VB_Name = "Main"
Option Explicit
Sub TARF()

    Const Nsimu As Long = 1000
    Dim simu(1 To Nsimu) As Variant, TTM(1 To Nsimu) As Double, PV(1 To Nsimu) As Double, Duration(1 To Nsimu) As Double
    
    Dim P1 As Double, P2 As Double
    P1 = Cells(4, 2)
    P2 = Cells(7, 2)
    Dim temp() As Double, rates() As Double
    ReDim temp(0 To P1 + P2) As Double
    ReDim rates(1 To P1 + P2) As Double
    
    Dim SPOT As Double, Target As Double, sum As Double
    SPOT = Range("F9").Value
    Target = Range("C9").Value
    
    Dim OVER As Boolean, KO As Boolean
    If Cells(2, 15) = "YES" Then
        OVER = True
    Else:
        OVER = False
    End If
    
    Dim Data As Range
    Set Data = Range("C12:L126")
    Dim i As Long, j As Integer
    
    For i = 1 To Nsimu
        TTM(i) = 0
        PV(i) = 0
        Duration(i) = P1 + P2
    Next i
    rates(1) = Data(2, 4) - Data(2, 6)
    For j = 2 To P1 + P2
        rates(j) = (Data(1 + j, 4) * (Data(1 + j, 2) - Data(1, 1)) - Data(j, 4) * (Data(j, 2) - Data(1, 1))) / (Data(1 + j, 2) - Data(j, 2)) _
                      - (Data(1 + j, 6) * (Data(1 + j, 2) - Data(1, 1)) - Data(j, 6) * (Data(j, 2) - Data(1, 1))) / (Data(1 + j, 2) - Data(j, 2))
    Next j
    
    temp(0) = SPOT
    Dim p As Double, index As Integer
    p = Data(1, 8)
    For i = 1 To Nsimu
        If Rnd > p Then
            index = 10
        Else: index = 8
        End If
        
        For j = 1 To P1 + P2
            temp(j) = temp(j - 1) * Exp((rates(j) / 100 - Data(1 + j, index) ^ 2 / 20000) * (Data(j + 1, 1) - Data(j, 1)) / 365 + Data(1 + j, index) / 100 * ((Data(j + 1, 1) - Data(j, 1)) / 365) ^ 0.5 * Rnd_normal())
        Next j
        simu(i) = temp
    Next i
    
    For i = 1 To Nsimu
        KO = False
        For j = 1 To P1
            If simu(i)(j) >= Cells(3, 3) And Not KO Then
                PV(i) = PV(i) + (simu(i)(j) - Cells(3, 8)) * Cells(3, 6) * Data(1 + j, 3)
                TTM(i) = TTM(i) + (simu(i)(j) - Cells(3, 8)) * Cells(3, 6)
            End If
            If Cells(4, 3) <= simu(i)(j) And simu(i)(j) < Cells(4, 5) And Not KO Then
                PV(i) = PV(i) + (simu(i)(j) - Cells(4, 8)) * Cells(4, 6) * Data(1 + j, 3)
                TTM(i) = TTM(i) + (simu(i)(j) - Cells(4, 8)) * Cells(4, 6)
            End If
            If simu(i)(j) < Cells(5, 5) And Not KO Then
                PV(i) = PV(i) + (simu(i)(j) - Cells(5, 8)) * Cells(5, 6) * Data(1 + j, 3)
                TTM(i) = TTM(i) + (simu(i)(j) - Cells(5, 8)) * Cells(5, 6)
            End If
            If TTM(i) >= Target And Not KO Then
                If Not OVER Then
                    PV(i) = PV(i) + (Target - TTM(i)) * Data(1 + j, 3)
                    TTM(i) = Target
                End If
                Duration(i) = j
                KO = True
            End If
        Next j
        For j = P1 + 1 To P1 + P2
            If simu(i)(j) >= Cells(6, 3) And Not KO Then
                PV(i) = PV(i) + (simu(i)(j) - Cells(6, 8)) * Cells(6, 6) * Data(1 + j, 3)
                TTM(i) = TTM(i) + (simu(i)(j) - Cells(6, 8)) * Cells(6, 6)
            End If
            If Cells(7, 3) <= simu(i)(j) And simu(i)(j) < Cells(7, 5) And Not KO Then
                PV(i) = PV(i) + (simu(i)(j) - Cells(7, 8)) * Cells(7, 6) * Data(1 + j, 3)
                TTM(i) = TTM(i) + (simu(i)(j) - Cells(7, 8)) * Cells(7, 6)
            End If
            If simu(i)(j) < Cells(8, 5) And Not KO Then
                PV(i) = PV(i) + (simu(i)(j) - Cells(8, 8)) * Cells(8, 6) * Data(1 + j, 3)
                TTM(i) = TTM(i) + (simu(i)(j) - Cells(8, 8)) * Cells(8, 6)
            End If
            If TTM(i) >= Target And Not KO Then
                If Not OVER Then
                    PV(i) = PV(i) + (Target - TTM(i)) * Data(1 + j, 3)
                    TTM(i) = Target
                End If
                Duration(i) = j
                KO = True
            End If
         Next j
    Next i
    
    sum = 0
    For i = 1 To Nsimu
        sum = sum + PV(i)
    Next i
    Cells(5, 15) = sum / Nsimu / SPOT
    
    sum = 0
    For i = 1 To Nsimu
        sum = sum + Duration(i)
    Next i
    Cells(5, 18) = sum / Nsimu
    
    ''''''''''''''''''''''''''''''''''''' Analytical ''''''''''''''''''''''''''''''''''''''''''
    For j = 1 To P1
        Cells(12 + j, 14) = (Cells(3, 6) * Cells(9, 6) * Cells(12 + j, 7) * Nd(1, 1, SPOT, Cells(3, 3), Cells(12 + j, 6), Cells(12 + j, 8), Cells(12 + j, 9), (Cells(12 + j, 4) - Cells(12, 3)) / 365) _
                                   - Cells(3, 6) * Cells(3, 8) * Cells(12 + j, 5) * Nd(2, 1, SPOT, Cells(3, 3), Cells(12 + j, 6), Cells(12 + j, 8), Cells(12 + j, 9), (Cells(12 + j, 4) - Cells(12, 3)) / 365)) * p + _
                                    (Cells(3, 6) * Cells(9, 6) * Cells(12 + j, 7) * Nd(1, 1, SPOT, Cells(3, 3), Cells(12 + j, 6), Cells(12 + j, 8), Cells(12 + j, 11), (Cells(12 + j, 4) - Cells(12, 3)) / 365) _
                                   - Cells(3, 6) * Cells(3, 8) * Cells(12 + j, 5) * Nd(2, 1, SPOT, Cells(3, 3), Cells(12 + j, 6), Cells(12 + j, 8), Cells(12 + j, 11), (Cells(12 + j, 4) - Cells(12, 3)) / 365)) * (1 - p)
        Cells(12 + j, 15) = (Cells(4, 6) * Cells(9, 6) * Cells(12 + j, 7) * Nd(1, 1, SPOT, Cells(4, 3), Cells(12 + j, 6), Cells(12 + j, 8), Cells(12 + j, 9), (Cells(12 + j, 4) - Cells(12, 3)) / 365) _
                                   - Cells(4, 6) * Cells(4, 8) * Cells(12 + j, 5) * Nd(2, 1, SPOT, Cells(4, 3), Cells(12 + j, 6), Cells(12 + j, 8), Cells(12 + j, 9), (Cells(12 + j, 4) - Cells(12, 3)) / 365)) * p + _
                                   (Cells(4, 6) * Cells(9, 6) * Cells(12 + j, 7) * Nd(1, 1, SPOT, Cells(4, 3), Cells(12 + j, 6), Cells(12 + j, 8), Cells(12 + j, 11), (Cells(12 + j, 4) - Cells(12, 3)) / 365) _
                                   - Cells(4, 6) * Cells(4, 8) * Cells(12 + j, 5) * Nd(2, 1, SPOT, Cells(4, 3), Cells(12 + j, 6), Cells(12 + j, 8), Cells(12 + j, 11), (Cells(12 + j, 4) - Cells(12, 3)) / 365)) * (1 - p)
        Cells(12 + j, 16) = (-Cells(4, 6) * Cells(9, 6) * Cells(12 + j, 7) * Nd(1, 1, SPOT, Cells(3, 3), Cells(12 + j, 6), Cells(12 + j, 8), Cells(12 + j, 9), (Cells(12 + j, 4) - Cells(12, 3)) / 365) _
                                   + Cells(4, 6) * Cells(4, 8) * Cells(12 + j, 5) * Nd(2, 1, SPOT, Cells(3, 3), Cells(12 + j, 6), Cells(12 + j, 8), Cells(12 + j, 9), (Cells(12 + j, 4) - Cells(12, 3)) / 365)) * p + _
                                   (-Cells(4, 6) * Cells(9, 6) * Cells(12 + j, 7) * Nd(1, 1, SPOT, Cells(3, 3), Cells(12 + j, 6), Cells(12 + j, 8), Cells(12 + j, 11), (Cells(12 + j, 4) - Cells(12, 3)) / 365) _
                                   + Cells(4, 6) * Cells(4, 8) * Cells(12 + j, 5) * Nd(2, 1, SPOT, Cells(3, 3), Cells(12 + j, 6), Cells(12 + j, 8), Cells(12 + j, 11), (Cells(12 + j, 4) - Cells(12, 3)) / 365)) * (1 - p)
        Cells(12 + j, 17) = (Cells(5, 6) * Cells(9, 6) * Cells(12 + j, 7) * Nd(1, -1, SPOT, Cells(4, 3), Cells(12 + j, 6), Cells(12 + j, 8), Cells(12 + j, 9), (Cells(12 + j, 4) - Cells(12, 3)) / 365) _
                                   - Cells(5, 6) * Cells(5, 8) * Cells(12 + j, 5) * Nd(2, -1, SPOT, Cells(4, 3), Cells(12 + j, 6), Cells(12 + j, 8), Cells(12 + j, 9), (Cells(12 + j, 4) - Cells(12, 3)) / 365)) * p + _
                                   (Cells(5, 6) * Cells(9, 6) * Cells(12 + j, 7) * Nd(1, -1, SPOT, Cells(4, 3), Cells(12 + j, 6), Cells(12 + j, 8), Cells(12 + j, 11), (Cells(12 + j, 4) - Cells(12, 3)) / 365) _
                                   - Cells(5, 6) * Cells(5, 8) * Cells(12 + j, 5) * Nd(2, -1, SPOT, Cells(4, 3), Cells(12 + j, 6), Cells(12 + j, 8), Cells(12 + j, 11), (Cells(12 + j, 4) - Cells(12, 3)) / 365)) * (1 - p)
    Next j
    For j = P1 + 1 To P1 + P2
        Cells(12 + j, 14) = (Cells(6, 6) * Cells(9, 6) * Cells(12 + j, 7) * Nd(1, 1, SPOT, Cells(6, 3), Cells(12 + j, 6), Cells(12 + j, 8), Cells(12 + j, 9), (Cells(12 + j, 4) - Cells(12, 3)) / 365) _
                                   - Cells(6, 6) * Cells(6, 8) * Cells(12 + j, 5) * Nd(2, 1, SPOT, Cells(6, 3), Cells(12 + j, 6), Cells(12 + j, 8), Cells(12 + j, 9), (Cells(12 + j, 4) - Cells(12, 3)) / 365)) * p + _
                                   (Cells(6, 6) * Cells(9, 6) * Cells(12 + j, 7) * Nd(1, 1, SPOT, Cells(6, 3), Cells(12 + j, 6), Cells(12 + j, 8), Cells(12 + j, 11), (Cells(12 + j, 4) - Cells(12, 3)) / 365) _
                                   - Cells(6, 6) * Cells(6, 8) * Cells(12 + j, 5) * Nd(2, 1, SPOT, Cells(6, 3), Cells(12 + j, 6), Cells(12 + j, 8), Cells(12 + j, 11), (Cells(12 + j, 4) - Cells(12, 3)) / 365)) * (1 - p)
        Cells(12 + j, 15) = (Cells(7, 6) * Cells(9, 6) * Cells(12 + j, 7) * Nd(1, 1, SPOT, Cells(7, 3), Cells(12 + j, 6), Cells(12 + j, 8), Cells(12 + j, 9), (Cells(12 + j, 4) - Cells(12, 3)) / 365) _
                                   - Cells(7, 6) * Cells(7, 8) * Cells(12 + j, 5) * Nd(2, 1, SPOT, Cells(7, 3), Cells(12 + j, 6), Cells(12 + j, 8), Cells(12 + j, 9), (Cells(12 + j, 4) - Cells(12, 3)) / 365)) * p + _
                                   (Cells(7, 6) * Cells(9, 6) * Cells(12 + j, 7) * Nd(1, 1, SPOT, Cells(7, 3), Cells(12 + j, 6), Cells(12 + j, 8), Cells(12 + j, 11), (Cells(12 + j, 4) - Cells(12, 3)) / 365) _
                                   - Cells(7, 6) * Cells(7, 8) * Cells(12 + j, 5) * Nd(2, 1, SPOT, Cells(7, 3), Cells(12 + j, 6), Cells(12 + j, 8), Cells(12 + j, 11), (Cells(12 + j, 4) - Cells(12, 3)) / 365)) * (1 - p)
        Cells(12 + j, 16) = (-Cells(7, 6) * Cells(9, 6) * Cells(12 + j, 7) * Nd(1, 1, SPOT, Cells(6, 3), Cells(12 + j, 6), Cells(12 + j, 8), Cells(12 + j, 9), (Cells(12 + j, 4) - Cells(12, 3)) / 365) _
                                   + Cells(7, 6) * Cells(7, 8) * Cells(12 + j, 5) * Nd(2, 1, SPOT, Cells(6, 3), Cells(12 + j, 6), Cells(12 + j, 8), Cells(12 + j, 9), (Cells(12 + j, 4) - Cells(12, 3)) / 365)) * p + _
                                   (-Cells(7, 6) * Cells(9, 6) * Cells(12 + j, 7) * Nd(1, 1, SPOT, Cells(6, 3), Cells(12 + j, 6), Cells(12 + j, 8), Cells(12 + j, 11), (Cells(12 + j, 4) - Cells(12, 3)) / 365) _
                                   + Cells(7, 6) * Cells(7, 8) * Cells(12 + j, 5) * Nd(2, 1, SPOT, Cells(6, 3), Cells(12 + j, 6), Cells(12 + j, 8), Cells(12 + j, 11), (Cells(12 + j, 4) - Cells(12, 3)) / 365)) * (1 - p)
        Cells(12 + j, 17) = (Cells(8, 6) * Cells(9, 6) * Cells(12 + j, 7) * Nd(1, -1, SPOT, Cells(7, 3), Cells(12 + j, 6), Cells(12 + j, 8), Cells(12 + j, 9), (Cells(12 + j, 4) - Cells(12, 3)) / 365) _
                                   - Cells(8, 6) * Cells(8, 8) * Cells(12 + j, 5) * Nd(2, -1, SPOT, Cells(7, 3), Cells(12 + j, 6), Cells(12 + j, 8), Cells(12 + j, 9), (Cells(12 + j, 4) - Cells(12, 3)) / 365)) * p + _
                                   (Cells(8, 6) * Cells(9, 6) * Cells(12 + j, 7) * Nd(1, -1, SPOT, Cells(7, 3), Cells(12 + j, 6), Cells(12 + j, 8), Cells(12 + j, 11), (Cells(12 + j, 4) - Cells(12, 3)) / 365) _
                                   - Cells(8, 6) * Cells(8, 8) * Cells(12 + j, 5) * Nd(2, -1, SPOT, Cells(7, 3), Cells(12 + j, 6), Cells(12 + j, 8), Cells(12 + j, 11), (Cells(12 + j, 4) - Cells(12, 3)) / 365)) * (1 - p)
    Next j
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''' Condition I '''''''''''''''''''''''''''''''''''''''''
    For j = 1 To P1
        sum = 0
        For i = 1 To Nsimu
            If simu(i)(j) >= Cells(3, 3) Then
                sum = sum + (simu(i)(j) - Cells(3, 8)) * Cells(3, 6) * Data(1 + j, 3)
            End If
        Next i
        Cells(12 + j, 18) = sum / Nsimu
    Next j
    For j = P1 + 1 To P1 + P2
        sum = 0
        For i = 1 To Nsimu
            If simu(i)(j) >= Cells(6, 3) Then
                sum = sum + (simu(i)(j) - Cells(6, 8)) * Cells(6, 6) * Data(1 + j, 3)
            End If
        Next i
        Cells(12 + j, 18) = sum / Nsimu
    Next j
    
    ''''''''''''''''''''''''''''''''''''' Condition II '''''''''''''''''''''''''''''''''''''''''
    For j = 1 To P1
        sum = 0
        For i = 1 To Nsimu
            If simu(i)(j) >= Cells(4, 3) Then
                sum = sum + (simu(i)(j) - Cells(4, 8)) * Cells(4, 6) * Data(1 + j, 3)
            End If
        Next i
        Cells(12 + j, 19) = sum / Nsimu
    Next j
    For j = P1 + 1 To P1 + P2
        sum = 0
        For i = 1 To Nsimu
            If simu(i)(j) >= Cells(7, 3) Then
                sum = sum + (simu(i)(j) - Cells(7, 8)) * Cells(7, 6) * Data(1 + j, 3)
            End If
        Next i
        Cells(12 + j, 19) = sum / Nsimu
    Next j
    
    ''''''''''''''''''''''''''''''''''''' Condition III  ''''''''''''''''''''''''''''''''''''''''''
    For j = 1 To P1
        sum = 0
        For i = 1 To Nsimu
            If simu(i)(j) >= Cells(3, 3) Then
                sum = sum - (simu(i)(j) - Cells(4, 8)) * Cells(4, 6) * Data(1 + j, 3)
            End If
        Next i
        Cells(12 + j, 20) = sum / Nsimu
    Next j
    For j = P1 + 1 To P1 + P2
        sum = 0
        For i = 1 To Nsimu
            If simu(i)(j) >= Cells(6, 3) Then
                sum = sum - (simu(i)(j) - Cells(7, 8)) * Cells(7, 6) * Data(1 + j, 3)
            End If
        Next i
        Cells(12 + j, 20) = sum / Nsimu
    Next j
    
    ''''''''''''''''''''''''''''''''''''' Condition IV ''''''''''''''''''''''''''''''''''''''''''
    For j = 1 To P1
        sum = 0
        For i = 1 To Nsimu
            If simu(i)(j) < Cells(4, 3) Then
                sum = sum + (simu(i)(j) - Cells(5, 8)) * Cells(5, 6) * Data(1 + j, 3)
            End If
        Next i
        Cells(12 + j, 21) = sum / Nsimu
    Next j
    For j = P1 + 1 To P1 + P2
        sum = 0
        For i = 1 To Nsimu
            If simu(i)(j) < Cells(7, 3) Then
                sum = sum + (simu(i)(j) - Cells(8, 8)) * Cells(8, 6) * Data(1 + j, 3)
            End If
        Next i
        Cells(12 + j, 21) = sum / Nsimu
    Next j
    
End Sub

Function Nd(O, sign, S, K, rd, rf, sigma, t)
    
    If K <> 0 Then
            Nd = Application.WorksheetFunction.Norm_S_Dist(sign * (Log(S / K) + (rd / 100 - rf / 100 + (3 - 2 * O) * sigma ^ 2 / 20000) * t) / (sigma / 100 * t ^ 0.5), True)
    Else: Nd = (sign + 1) * 0.5
    End If
    
End Function

Function Rnd_normal()

    Dim r As Double
    r = Rnd
    If r = 0 Then r = 1E-300
    Rnd_normal = Application.WorksheetFunction.NormInv(r, 0, 1)

End Function

