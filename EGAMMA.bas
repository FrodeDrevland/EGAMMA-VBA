Option Explicit

Function EGAMMA_DIST(x As Double, alpha As Double, beta As Double, delta As Double, cumulative As Boolean) As Double
    If cumulative Then
        ' CDF
        If beta > 0 Then
            ' Support: x >= delta
            If x <= delta Then
                EGAMMA_DIST = 0
            Else
                EGAMMA_DIST = WorksheetFunction.Gamma_Dist(x - delta, alpha, beta, True)
            End If
        ElseIf beta < 0 Then
            ' Support: x <= delta
            If x >= delta Then
                EGAMMA_DIST = 1
            Else
                EGAMMA_DIST = 1 - WorksheetFunction.Gamma_Dist(delta - x, alpha, Abs(beta), True)
            End If
        Else
            EGAMMA_DIST = CVErr(xlErrNum)   ' beta = 0 is invalid
        End If
    Else
        ' PDF
        If beta = 0 Then
            EGAMMA_DIST = CVErr(xlErrNum)
        ElseIf (x - delta) / beta < 0 Then
            EGAMMA_DIST = CVErr(xlErrNA)    ' outside support
        Else
            EGAMMA_DIST = WorksheetFunction.Gamma_Dist(Abs(x - delta), alpha, Abs(beta), False)
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
    EGAMMA_STDDEV = Sqr(alpha) * Abs(beta)
End Function

Function EGAMMA_SKEW(alpha As Double, beta As Double) As Double
    EGAMMA_SKEW = (2 / Sqr(alpha)) * Sgn(beta)
End Function

Function EGAMMA_KURT(alpha As Double) As Double
    EGAMMA_KURT = 6 / alpha
End Function

Function EGAMMA_TPE_TO_PARAMS(low As Double, likely As Double, high As Double, Optional low_probability As Double = 0.1)
    Dim returnVal(1 To 3) As Double
    
    If (low = likely And high = likely) Or low > likely Or high < likely Then
        EGAMMA_TPE_TO_PARAMS = CVErr(xlErrNA)
 
    Else
        If low = likely Or high = likely Then
            returnVal(1) = FindAlphaAtModeEqualsProbability(low_probability)
        Else
            returnVal(1) = FindAlpha(low, likely, high, low_probability)
        End If
        returnVal(2) = (high - low) / (WorksheetFunction.Gamma_Inv(1 - low_probability, returnVal(1), 1) - WorksheetFunction.Gamma_Inv(low_probability, returnVal(1), 1))
        If high - likely < likely - low Then returnVal(2) = -returnVal(2)
        returnVal(3) = likely - (returnVal(1) - 1) * returnVal(2)
        EGAMMA_TPE_TO_PARAMS = returnVal
    End If
End Function

Private Function FindAlphaAtModeEqualsProbability(probability As Double, Optional decimals As Integer = 10) As Double
    Dim skew_low As Double
    Dim skew_high As Double
    Dim skew_mid As Double
    Dim alpha_candidate As Double
    Dim mode As Double
    Dim mode_candidate As Double
    Dim iter As Long
    
    If probability > 0.5 Then
    probability = 1 - probability
    End If
    
    skew_low = 2 / Sqr(1000000000#)
    skew_high = 2
    iter = 0
    
    While skew_low <= skew_high And iter < 200
        iter = iter + 1
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
    If iter >= 200 Then
        FindAlphaAtModeEqualsProbability = alpha_candidate 'best effort
    End If
End Function

Private Function FindAlpha(low As Double, mode As Double, high As Double, Optional low_prob As Double = 0.1, Optional threshold As Double = 0.0000000001) As Double
    Dim skew_low As Double
    Dim skew_high As Double
    Dim skew_mid As Double
    Dim alpha_candidate As Double
    Dim target_ratio As Double
    Dim current_ratio As Double
    Dim iter As Long

    
 target_ratio = (mode - low) / (high - mode)
    If Abs(target_ratio) > 1 Then target_ratio = 1 / target_ratio

    If target_ratio > 0.99999 Then
        FindAlpha = 1000000000#
        Exit Function
    End If
    

    skew_low = 2 / (Sqr(1000000000#))
    skew_high = 2
    iter = 0
    
    While skew_low <= skew_high And iter < 200
        iter = iter + 1
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
    If iter >= 200 Then
        FindAlpha = alpha_candidate 'best effort
    End If
End Function

Function EGAMMA_FIT_TO_PARAMS(ParamArray args() As Variant)
    Dim alpha As Double
    Dim beta As Double
    Dim delta As Double
    Dim returnVal(1 To 3) As Double
    Dim result() As Double
    Dim arg As Variant
    Dim cell As Range
    Dim itemCount As Long
    Dim i As Long
    Dim skew As Double
    
    ' Flatten args into a 1D array of Doubles
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
    
    ' Method-of-moments fit
    skew = WorksheetFunction.skew(result)
    If skew = 0 Then
        ' effectively symmetric; use very large alpha, small beta
        alpha = 1000000000#
        beta = WorksheetFunction.StDev(result) / Sqr(alpha)
    Else
        alpha = 4 / (skew * skew)              ' from |skew| = 2 / sqrt(alpha)
        beta = WorksheetFunction.StDev(result) / Sqr(alpha)
        If skew < 0 Then beta = -beta          ' left-skew => beta < 0
    End If
    
    delta = WorksheetFunction.Average(result) - alpha * beta
    
    returnVal(1) = alpha
    returnVal(2) = beta
    returnVal(3) = delta
    
    EGAMMA_FIT_TO_PARAMS = returnVal
End Function




'=============================
'  REGISTRATION MACRO
'=============================
Public Sub RegisterEGammaFunctions()
    With Application
        .MacroOptions _
            Macro:="EGAMMA_DIST", _
            Description:="Expanded gamma distribution (shifted and signed). Returns PDF or CDF.", _
            Category:="User Defined", _
            ArgumentDescriptions:=Array( _
                "x: Value at which to evaluate the distribution.", _
                "alpha: Shape parameter (>0).", _
                "beta: Scale parameter (sign controls tail direction).", _
                "delta: Location (shift) parameter.", _
                "cumulative: TRUE for cumulative distribution, FALSE for density." _
            )

        .MacroOptions _
            Macro:="EGAMMA_INV", _
            Description:="Inverse expanded gamma distribution (quantile function).", _
            Category:="User Defined", _
            ArgumentDescriptions:=Array( _
                "probability: Cumulative probability (0–1).", _
                "alpha: Shape parameter (>0).", _
                "beta: Scale parameter (sign controls tail direction).", _
                "delta: Location (shift) parameter." _
            )

        .MacroOptions _
            Macro:="EGAMMA_MEAN", _
            Description:="Mean of the expanded gamma distribution.", _
            Category:="User Defined", _
            ArgumentDescriptions:=Array( _
                "alpha: Shape parameter (>0).", _
                "beta: Scale parameter (can be signed).", _
                "delta: Location (shift) parameter." _
            )

        .MacroOptions _
            Macro:="EGAMMA_MODE", _
            Description:="Mode of the expanded gamma distribution.", _
            Category:="User Defined", _
            ArgumentDescriptions:=Array( _
                "alpha: Shape parameter (>1 for a proper mode).", _
                "beta: Scale parameter (can be signed).", _
                "delta: Location (shift) parameter." _
            )

        .MacroOptions _
            Macro:="EGAMMA_MEDIAN", _
            Description:="Median of the expanded gamma distribution (via EGAMMA_INV at p = 0.5).", _
            Category:="User Defined", _
            ArgumentDescriptions:=Array( _
                "alpha: Shape parameter (>0).", _
                "beta: Scale parameter (can be signed).", _
                "delta: Location (shift) parameter." _
            )

        .MacroOptions _
            Macro:="EGAMMA_VAR", _
            Description:="Variance of the expanded gamma distribution.", _
            Category:="User Defined", _
            ArgumentDescriptions:=Array( _
                "alpha: Shape parameter (>0).", _
                "beta: Scale parameter." _
            )

        .MacroOptions _
            Macro:="EGAMMA_STDDEV", _
            Description:="Standard deviation of the expanded gamma distribution.", _
            Category:="User Defined", _
            ArgumentDescriptions:=Array( _
                "alpha: Shape parameter (>0).", _
                "beta: Scale parameter." _
            )

        .MacroOptions _
            Macro:="EGAMMA_SKEW", _
            Description:="Skewness of the expanded gamma distribution.", _
            Category:="User Defined", _
            ArgumentDescriptions:=Array( _
                "alpha: Shape parameter (>0).", _
                "beta: Scale parameter." _
            )

        .MacroOptions _
            Macro:="EGAMMA_KURT", _
            Description:="Excess kurtosis of the expanded gamma distribution.", _
            Category:="User Defined", _
            ArgumentDescriptions:=Array( _
                "alpha: Shape parameter (>0)." _
            )

        .MacroOptions _
            Macro:="EGAMMA_TPE_TO_PARAMS", _
            Description:="Fits expanded gamma parameters from three-point estimate (low, likely, high).", _
            Category:="User Defined", _
            ArgumentDescriptions:=Array( _
                "low: Lower bound of three-point estimate.", _
                "likely: Most likely (mode) value.", _
                "high: Upper bound of three-point estimate.", _
                "low_probability: Cumulative probability at low/high (default 0.1)." _
            )

        .MacroOptions _
            Macro:="EGAMMA_FIT_TO_PARAMS", _
            Description:="Fits expanded gamma parameters (alpha, beta, delta) to sample data.", _
            Category:="User Defined", _
            ArgumentDescriptions:=Array( _
                "args: Sample values and/or ranges to fit the distribution to." _
            )
    End With
End Sub


