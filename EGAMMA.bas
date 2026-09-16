Option Explicit

' Largest shape parameter the fitting routines will return. A symmetric
' three-point estimate is only reproducible in the limit as alpha tends to
' infinity, so a finite stand-in is needed. Excel's GAMMA.INV also loses
' reliability well before this value, which is why it is not set higher.
'
' This value doubles as the signal that an estimate was too near symmetry to
' resolve: EGAMMA_TPE_TO_PARAMS returns exactly this shape in that case. A
' caller that cares can compare the returned shape with 1E9. No separate
' function reports it, because the shape is already in the cell the user typed
' the formula in.
'
' At this ceiling, and at the default 10th/90th percentile convention, the
' elicited values are reproduced to about 1.5E-5 of the elicited range rather
' than to TOLERANCE. That figure belongs to the convention: the ceiling error
' grows as the elicited percentiles approach the median, reaching about 4.2E-2
' at a low probability of 0.4999. Raising the ceiling does reduce it; it is not
' raised because the shape becomes increasingly ill-conditioned near symmetry
' and GAMMA.INV loses reliability at large shapes.
Private Const ALPHA_MAX As Double = 1000000000#

' Acceptance tolerance on the normalised mode position. The search stops when
' the fitted mode position differs from the elicited one by less than this, and
' that difference is itself the maximum normalised error in the three reproduced
' values, so the tolerance is stated directly in the quantity a user cares about.
' It replaces the relative threshold on the half-range ratio used up to version
' 1.1.1, which demanded ever finer absolute agreement as the mode approached an
' outer value. The value preserves the reproduction accuracy that the old
' relative threshold of 1E-10 implied.
Private Const TOLERANCE As Double = 0.000000000025

Private Const MAX_ITER As Long = 200

' Library version. Reported by EGAMMA_VERSION() so a workbook can record which
' build produced its numbers.
Private Const EGAMMA_LIB_VERSION As String = "1.2.1"


Function EGAMMA_VERSION() As String
    EGAMMA_VERSION = EGAMMA_LIB_VERSION
End Function


' Position of the standard Gamma mode within the interval spanned by the two
' percentiles, as a fraction of that interval. This is the quantity the
' three-point fit matches: the elicited target is the mode's position within
' the elicited range, so the two are directly comparable.
Private Function ModePosition(alpha As Double, low_prob As Double) As Double
    Dim q_low As Double
    Dim q_high As Double
    q_low = WorksheetFunction.Gamma_Inv(low_prob, alpha, 1)
    q_high = WorksheetFunction.Gamma_Inv(1 - low_prob, alpha, 1)
    ModePosition = (alpha - 1 - q_low) / (q_high - q_low)
End Function

' Returns Variant so that an error value can be returned; a Double-typed
' function cannot hold CVErr and raises a type mismatch instead.
Function EGAMMA_DIST(x As Double, alpha As Double, beta As Double, delta As Double, cumulative As Boolean) As Variant
    If alpha <= 0 Then
        EGAMMA_DIST = CVErr(xlErrNum)
        Exit Function
    End If

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
        ElseIf (x - delta) / beta <= 0 Then
            EGAMMA_DIST = 0                 ' outside support: density is zero
        Else
            EGAMMA_DIST = WorksheetFunction.Gamma_Dist(Abs(x - delta), alpha, Abs(beta), False)
        End If
    End If
End Function


Function EGAMMA_INV(probability As Double, alpha As Double, beta As Double, delta As Double) As Variant
    If alpha <= 0 Or beta = 0 Or probability <= 0 Or probability >= 1 Then
        EGAMMA_INV = CVErr(xlErrNum)
        Exit Function
    End If

    If beta > 0 Then
        EGAMMA_INV = delta + WorksheetFunction.Gamma_Inv(probability, alpha, beta)
    Else
        EGAMMA_INV = delta - WorksheetFunction.Gamma_Inv(1 - probability, alpha, Abs(beta))
    End If
End Function

Function EGAMMA_MEAN(alpha As Double, beta As Double, delta As Double) As Double
    EGAMMA_MEAN = alpha * beta + delta
End Function

' For alpha > 1 the density has an interior maximum at (alpha - 1) * beta +
' delta. For 0 < alpha <= 1 it is monotone on its support and the mode is at
' the support boundary, delta; the interior expression would place it outside
' the support. A three-point fit never returns a shape at or below 1, so this
' distinction arises only for parameters entered directly.
Function EGAMMA_MODE(alpha As Double, beta As Double, delta As Double) As Variant
    If alpha <= 0 Then
        EGAMMA_MODE = CVErr(xlErrNum)
    ElseIf alpha <= 1 Then
        EGAMMA_MODE = delta
    Else
        EGAMMA_MODE = (alpha - 1) * beta + delta
    End If
End Function
' Returns Variant because EGAMMA_INV returns an error value for invalid
' parameters, and a Double-typed function cannot hold one: it raises a type
' mismatch at run time instead of putting #NUM! in the cell.
Function EGAMMA_MEDIAN(alpha As Double, beta As Double, delta As Double) As Variant
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
    
    If low_probability <= 0 Or low_probability >= 0.5 Then
        EGAMMA_TPE_TO_PARAMS = CVErr(xlErrNum)

    ElseIf low >= high Or low > likely Or high < likely Then
        EGAMMA_TPE_TO_PARAMS = CVErr(xlErrNA)

    Else
        ' A mode at an outer value gives a target position of zero and goes
        ' through the ordinary search; there is no separate endpoint routine.
        returnVal(1) = FindAlpha(low, likely, high, low_probability)

        ' The shape search reports failure as a non-positive value. Propagate
        ' it rather than using it as a shape parameter.
        If returnVal(1) <= 0 Then
            EGAMMA_TPE_TO_PARAMS = CVErr(xlErrNum)
            Exit Function
        End If

        returnVal(2) = (high - low) / (WorksheetFunction.Gamma_Inv(1 - low_probability, returnVal(1), 1) - WorksheetFunction.Gamma_Inv(low_probability, returnVal(1), 1))
        If high - likely < likely - low Then returnVal(2) = -returnVal(2)
        returnVal(3) = likely - (returnVal(1) - 1) * returnVal(2)
        EGAMMA_TPE_TO_PARAMS = returnVal
    End If
End Function

Private Function FindAlpha(low As Double, mode As Double, high As Double, _
                           Optional low_prob As Double = 0.1) As Double
    ' Finds the shape parameter by bisection on the magnitude of skewness,
    ' matching the mode's position within the elicited range:
    '
    '     target_position = min(mode - low, high - mode) / (high - low)
    '     candidate       = (alpha - 1 - q_low) / (q_high - q_low)
    '
    ' and stopping when the two differ by less than the tolerance. That
    ' difference is the maximum normalised error in the three reproduced
    ' values, so the stopping rule is expressed in the quantity being promised.
    '
    ' Matching positions rather than the half-range ratio matters near an outer
    ' value. The ratio tends to zero there, so a relative test on it would
    ' demand progressively finer absolute agreement even though the accuracy
    ' required of the elicited values had not changed. Both skew directions and
    ' the mode-at-an-outer-value cases reduce to the same target, so no separate
    ' endpoint routine is needed.
    '
    ' Returns 0 to signal that the tolerance could not be met; callers must
    ' check for this rather than using the value.
    Dim skew_low As Double
    Dim skew_high As Double
    Dim skew_mid As Double
    Dim alpha_candidate As Double
    Dim target_position As Double
    Dim candidate_position As Double
    Dim max_position As Double
    Dim iter As Long

    If high - mode < mode - low Then
        target_position = (high - mode) / (high - low)
    Else
        target_position = (mode - low) / (high - low)
    End If

    ' The greatest position the search can reach is the one ALPHA_MAX produces,
    ' and it depends on the percentile convention: about 0.499985 at P_L = 0.10
    ' but 0.457948 at P_L = 0.4999. Deriving the cut from ALPHA_MAX rather than
    ' fixing it keeps the root bracketed under any convention.
    max_position = ModePosition(ALPHA_MAX, low_prob)

    If max_position <= 0 Then
        FindAlpha = 0                           ' ceiling below the admissible range
        Exit Function
    End If

    If target_position >= max_position Then
        FindAlpha = ALPHA_MAX                   ' ceiling approximation
        Exit Function
    End If

    skew_low = 2 / Sqr(ALPHA_MAX)
    skew_high = 2

    For iter = 1 To MAX_ITER
        skew_mid = (skew_low + skew_high) / 2
        If skew_mid = skew_low Or skew_mid = skew_high Then
            FindAlpha = 0                       ' interval collapsed
            Exit Function
        End If
        alpha_candidate = 4 / (skew_mid ^ 2)
        candidate_position = ModePosition(alpha_candidate, low_prob)

        If Abs(candidate_position - target_position) < TOLERANCE Then
            FindAlpha = alpha_candidate
            Exit Function
        ElseIf candidate_position < target_position Then
            skew_high = skew_mid
        Else
            skew_low = skew_mid
        End If
    Next iter

    FindAlpha = 0                               ' iteration limit reached
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
    If Abs(skew) < 0.000000001 Then
        ' effectively symmetric; use very large alpha, small beta
        alpha = ALPHA_MAX
        beta = WorksheetFunction.StDev(result) / Sqr(alpha)
    Else
        alpha = 4 / (skew * skew)              ' from |skew| = 2 / sqrt(alpha)
        If alpha > ALPHA_MAX Then alpha = ALPHA_MAX
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
                "probability: Cumulative probability (0 to 1).", _
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
            Macro:="EGAMMA_VERSION", _
            Description:="Version of the EGAMMA-VBA library.", _
            Category:="User Defined"

        .MacroOptions _
            Macro:="EGAMMA_FIT_TO_PARAMS", _
            Description:="Fits expanded gamma parameters (alpha, beta, delta) to sample data.", _
            Category:="User Defined", _
            ArgumentDescriptions:=Array( _
                "args: Sample values and/or ranges to fit the distribution to." _
            )
    End With
End Sub


