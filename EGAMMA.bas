Attribute VB_Name = "EGAMMA"
Option Explicit

Function EGAMMA_DIST(x As Double, alpha As Double, beta As Double, delta As Double, cumulative As Boolean) As Double
    If cumulative Then
        If beta > 0 Then
            EGAMMA_DIST = WorksheetFunction.Gamma_Dist(Abs(x - delta), alpha, beta, cumulative)
        Else
            EGAMMA_DIST = 1 - WorksheetFunction.Gamma_Dist(Abs(x - delta), alpha, Abs(beta), cumulative)
        End If
    Else
        If (x - delta) / beta < 0 Then
            EGAMMA_DIST = CVErr(xlErrNA)
        Else
            EGAMMA_DIST = WorksheetFunction.Gamma_Dist(Abs(x - delta), alpha, Abs(beta), cumulative)
        End If
    End If
End Function

Function EGAMMA_INV(probability As Double, alpha As Double, beta As Double, delta As Double) As Double
     If beta > 0 Then
        EGAMMA_INV = delta + WorksheetFunction.Gamma_Inv(probability, alpha, beta)
    Else
        EGAMMA_INV = delta - WorksheetFunction.Gamma_Inv(1 - probability, alpha, Abs(beta))
    End If
End Function

Function EGAMMA_MEAN(alpha As Double, beta As Double, delta As Double) As Double
    EGAMMA_MEAN = alpha * beta + delta
End Function
Function EGAMMA_MODE(alpha As Double, beta As Double, delta As Double) As Double
    EGAMMA_MODE = (alpha - 1) * beta + delta
End Function
Function EGAMMA_MEDIAN(alpha As Double, beta As Double, delta As Double) As Double
    EGAMMA_MEDIAN = EGAMMA_INV(0.5, alpha, beta, delta)
End Function

Function EGAMMA_VAR(alpha As Double, beta As Double) As Double
    EGAMMA_VAR = alpha * beta * beta
End Function

Function EGAMMA_STDDEV(alpha As Double, beta As Double) As Double
    EGAMMA_STDDEV = Sqr(alpha) * beta
End Function

Function EGAMMA_SKEW(alpha As Double, beta As Double) As Double
    EGAMMA_SKEW = 2 / Sqr(alpha)
End Function

Function EGAMMA_KURT(alpha As Double, beta As Double) As Double
    EGAMMA_KURT = 6 / alpha
End Function

Function EGAMMA_TPE_TO_PARAMS(low As Double, likely As Double, high As Double, Optional low_probability As Double = 0.1)
    Dim returnVal(1 To 3) As Double
    
    If (low = likely And high = likely) Or low > likely Or high < likely Then
        EGAMMA_TPE_TO_PARAMS = CVErr(xlErrNA)
     End If
    
    If low = likely Or high = likely Then
        returnVal(1) = FindAlphaAtModeEqualsProbability(low_probability)
    Else
        returnVal(1) = FindAlpha(low, likely, high, low_probability)
    End If
    returnVal(2) = (high - low) / (WorksheetFunction.Gamma_Inv(1 - low_probability, returnVal(1), 1) - WorksheetFunction.Gamma_Inv(low_probability, returnVal(1), 1))
    If high - likely < likely - low Then returnVal(2) = -returnVal(2)
    returnVal(3) = likely - (returnVal(1) - 1) * returnVal(2)
    EGAMMA_TPE_TO_PARAMS = returnVal
End Function

Private Function FindAlphaAtModeEqualsProbability(probability As Double, Optional decimals As Integer = 10) As Double
    Dim skew_low As Double
    Dim skew_high As Double
    Dim skew_mid As Double
    Dim alpha_candidate As Double
    Dim mode As Double
    Dim mode_candidate As Double
    
    If probability > 0.5 Then
    probability = 1 - probability
    End If
    
    skew_low = 2 / Sqr(1000000000#)
    skew_high = 2
    
    While skew_low <= skew_high
        skew_mid = (skew_low + skew_high) / 2
        alpha_candidate = 4 / (skew_mid ^ 2)
        mode = alpha_candidate - 1
        mode_candidate = WorksheetFunction.Gamma_Inv(probability, alpha_candidate, 1)

        If Round(mode_candidate, decimals) = Round(mode, decimals) Then
            FindAlphaAtModeEqualsProbability = alpha_candidate
            Exit Function
        ElseIf mode_candidate > mode Then
            skew_high = skew_mid
        Else:
            skew_low = skew_mid
        End If
    Wend
End Function

Private Function FindAlpha(low As Double, mode As Double, high As Double, Optional low_prob As Double = 0.1, Optional threshold As Double = 0.0000000001) As Double
    Dim skew_low As Double
    Dim skew_high As Double
    Dim skew_mid As Double
    Dim alpha_candidate As Double
    Dim target_ratio As Double
    Dim current_ratio As Double
    
 target_ratio = (mode - low) / (high - mode)
    If Abs(target_ratio) > 1 Then target_ratio = 1 / target_ratio

    If target_ratio > 0.99999 Then
        FindAlpha = 1000000000#
        Exit Function
    End If
    

    skew_low = 2 / (Sqr(1000000000#))
    skew_high = 2
    While skew_low <= skew_high
        skew_mid = (skew_low + skew_high) / 2
        alpha_candidate = 4 / (skew_mid ^ 2)
        current_ratio = ((alpha_candidate - 1) - WorksheetFunction.Gamma_Inv(low_prob, alpha_candidate, 1)) / (WorksheetFunction.Gamma_Inv(1 - low_prob, alpha_candidate, 1) - (alpha_candidate - 1))

        If Abs((current_ratio / target_ratio) - 1) < threshold Then
             FindAlpha = alpha_candidate
             Exit Function
        ElseIf current_ratio < target_ratio Then
            skew_high = skew_mid
        Else:
            skew_low = skew_mid
        End If
    Wend
End Function

Function EGAMMA_FIT_TO_PARAMS(ParamArray args() As Variant)
    Dim alpha As Double
    Dim beta As Double
    Dim delta As Double
    Dim returnVal(1 To 3) As Double
    Dim result() As Variant
    Dim arg As Variant
    Dim cell As Range
    Dim itemCount As Long
    Dim i As Long
    Dim skew As Double
    

    itemCount = 0

    For Each arg In args
        If TypeName(arg) = "Range" Then
            itemCount = itemCount + arg.Cells.Count
        Else
            itemCount = itemCount + 1
        End If
    Next arg

    ReDim result(1 To itemCount)

    i = 1

    For Each arg In args
        If TypeName(arg) = "Range" Then
            For Each cell In arg
                result(i) = cell.Value
                i = i + 1
            Next cell
        Else
            result(i) = arg
            i = i + 1
        End If
    Next arg
       
    skew = WorksheetFunction.skew(result)
    alpha = 4 / skew ^ 2
    beta = WorksheetFunction.StDev(result) / Sqr(alpha)
    delta = WorksheetFunction.Average(result) - alpha * beta


    returnVal(1) = alpha
    returnVal(2) = beta
    returnVal(3) = delta
    
    EGAMMA_FIT_TO_PARAMS = returnVal
    
End Function





