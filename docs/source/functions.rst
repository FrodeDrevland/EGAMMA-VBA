.. _functions:

Functions
=========
.. py:function:: EGAMMA_DIST(x, alpha, beta, delta, cumulative):

    Calculate the expanded gamma distribution.

    :param x: Required. The value at which the function is evaluated.
    :param alpha: Required. The shape parameter of the distribution.
    :param beta: Required. The scale parameter of the distribution. Negative beta entails a left skewed distribution.
    :param delta: Required. The location  parameter of the distribution.
    :param cumulative: Required. A logical value that determines the form of the function. If cumulative is TRUE, GAMMA.DIST returns the cumulative distribution function; if FALSE, it returns the probability density function.
    :return: The value of the (cumulative) gamma distribution.
    :rtype: float

.. py:function:: EGAMMA_INV(probability, alpha, beta, delta):

    Calculate the inverse of the expanded gamma distribution.

    :param probability: The probability for which the inverse is calculated.
    :param alpha: The shape parameter of the expanded gamma distribution.
    :param beta: The scale parameter of the expanded gamma distribution.
    :param delta: The location  parameter of the expanded gamma distribution.
    :return: The inverse value of the expanded gamma distribution for the given probability.

.. py:function:: EGAMMA_MEAN(alpha, beta, delta):
    
    Calculate the mean of the expanded gamma distribution.

    :param alpha: The shape parameter of the expanded gamma distribution.
    :param beta: The scale parameter of the expanded gamma distribution.
    :param delta: The location  parameter of the expanded gamma distribution.
    :return: The mean of the expanded gamma distribution.
    

.. py:function:: EGAMMA_MODE(alpha, beta, delta):
    
    Calculate the mode of the expanded gamma distribution.

    For :math:`\alpha > 1` the density has an interior maximum at
    :math:`(\alpha - 1)\beta + \delta`. For :math:`0 < \alpha \le 1` it is
    monotone on its support and the mode lies at the support boundary,
    :math:`\delta`; the interior expression would place it outside the support
    altogether. A three-point fit never returns a shape at or below 1, so this
    distinction arises only for parameters entered directly.

    :param alpha: The shape parameter of the expanded gamma distribution.
    :param beta: The scale parameter of the expanded gamma distribution.
    :param delta: The location  parameter of the expanded gamma distribution.
    :return: The mode of the expanded gamma distribution, or ``#NUM!`` if the
        shape parameter is not positive.
    


.. py:function:: EGAMMA_MEDIAN(alpha, beta, delta):
    
    Calculate the median of the expanded gamma distribution.

    :param alpha: The shape parameter of the expanded gamma distribution.
    :param beta: The scale parameter of the expanded gamma distribution.
    :param delta: The location  parameter of the expanded gamma distribution.
    :return: The median of the expanded gamma distribution.
    


.. py:function:: EGAMMA_VAR(alpha, beta):
    
    Calculate the variance of the expanded gamma distribution.

    :param alpha: The shape parameter of the expanded gamma distribution.
    :param beta: The scale parameter of the expanded gamma distribution.
    :return: The variance of the expanded gamma distribution.
    


.. py:function:: EGAMMA_STDDEV(alpha, beta):

    Calculate the standard deviation of the expanded gamma distribution.

    :param alpha: The shape parameter of the expanded gamma distribution.
    :param beta: The scale parameter of the expanded gamma distribution.
    :return: The standard deviation of the expanded gamma distribution.
    


.. py:function:: EGAMMA_SKEW(alpha, beta):
    
    Calculate the skewness of the expanded gamma distribution.

    :param alpha: The shape parameter of the expanded gamma distribution.
    :param beta: The scale parameter of the expanded gamma distribution.
    :return: The skewness of the expanded gamma distribution.
    
.. py:function:: EGAMMA_KURT(alpha):

    Calculate the excess kurtosis of the expanded gamma distribution.

    :param alpha: The shape parameter of the expanded gamma distribution.
    :return: The excess kurtosis of the expanded gamma distribution.

.. py:function:: EGAMMA_FIT_TO_PARAMS(*args):
    
    Fit the parameters of the expanded gamma distribution to a given set of data points.

    This function estimates the alpha, beta, and delta parameters of the expanded  gamma distribution,
    based on the provided data using the method of moments. It uses the skewness of the data to estimate alpha, then calculates beta
    using the standard deviation, and finally computes delta from the mean, alpha, and beta.

    :param args: A variable number of arrays or lists containing numerical data points.
                 These are combined into a single dataset for the analysis.
    :return: A list containing the estimated parameters [alpha, beta, delta] of the extended gamma distribution.
    :raises ValueError: If the input data does not allow for a proper fit, e.g., if it leads to divide-by-zero errors.
    

.. py:function:: EGAMMA_TPE_TO_PARAMS(low, likely, high, low_probability=0.1, fit_tolerance=2.5E-11, ceiling_tolerance=2E-5, require_nonnegative=FALSE):
    
    Estimate the parameters of the expanded gamma distribution based on a three-point estimate.

    This function uses a low, likely, and high estimate to determine the parameters of the extended gamma distribution.

    The shape parameter is found by bisection on the magnitude of the
    skewness. The quantity matched is the mode's position within the elicited
    range,

    .. math::
        t = \dfrac{\min(\text{likely} - \text{low},\ \text{high} - \text{likely})}
                  {\text{high} - \text{low}},

    against the position the standard Gamma with that shape produces between
    its two percentiles,

    .. math::
        t_\alpha = \dfrac{\alpha - 1 - q_L}{q_H - q_L},

    and the search stops when the two differ by less than
    :math:`2.5\times10^{-11}`. That difference is itself the largest error, as
    a fraction of the elicited range, with which the three elicited values are
    recovered from the fitted parameters, so the tolerance is stated directly
    in the quantity of interest. A most likely value equal to an outer value
    gives :math:`t = 0` and goes through the same search; there is no separate
    case for it.

    An estimate too near symmetry to resolve cannot be fitted to that
    tolerance, because a perfectly symmetric estimate is only reproducible in
    the limit as the shape tends to infinity. Such an estimate is fitted at
    the library's finite ceiling instead, and the returned shape is then
    exactly :math:`10^9`, which is how a caller recognises the case. The
    elicited values are reproduced to about :math:`1.5\times10^{-5}` of the
    elicited range rather than to the tolerance, and
    :func:`EGAMMA_TPE_DIAGNOSTICS` reports the case as ``"approximation"``.

    Once the parameters are found, the result is accepted only if three rules
    hold. The tolerance is ``fit_tolerance`` for an ordinary fit and
    ``ceiling_tolerance`` for a fit at the ceiling.

    1. The position residual :math:`|t_\alpha - t|`, which is the normalised
       reproduction error in exact arithmetic, is within the tolerance. No
       allowance is added.
    2. The estimated rounding uncertainty of the reconstruction check,
       :math:`4\,\epsilon_{\text{mach}}\,m/(\text{high}-\text{low})`, where
       :math:`m` is the largest magnitude among the elicited values, the
       location and the scaled quantiles, is no larger than the tolerance.
       Otherwise the check cannot verify the accuracy asked for: in double
       precision this happens at the default tolerance once the values exceed
       about 28,000 times their range. The factor 4 is an estimate from a
       first-order count of rounded operations, not a proven bound.
    3. The low, likely and high values reconstructed from the parameters agree
       with those supplied to within the tolerance plus that uncertainty.

    A result that fails any rule returns ``#NUM!``, and
    :func:`EGAMMA_TPE_DIAGNOSTICS` names the rule.

    :param low: The low estimate of the distribution.
    :param likely: The most likely estimate of the distribution.
    :param high: The high estimate of the distribution.
    :param low_probability: The probability of the low  estimates (default is 0.1).
    :param fit_tolerance: The position residual accepted for an ordinary fit,
        as a fraction of the range (default 2.5E-11). Loosen it for an
        estimate refused as unverifiable, or fit shifted values instead.
    :param ceiling_tolerance: The position residual accepted at the shape
        ceiling, as a fraction of the range (default 2E-5). All five supported
        percentile conventions meet the default for a symmetric estimate; a
        convention near the median, such as 0.4999, does not.
    :param require_nonnegative: TRUE to refuse, with ``#NUM!``, a fit whose
        distribution has support below zero (default FALSE).
    :return: A list containing the estimated parameters [alpha, beta, delta] of the extended gamma distribution.
    :raises ValueError: Returns ``#N/A`` if the input values are inconsistent (low > likely, high < likely, or low >= high), and ``#NUM!`` if an argument is out of range or the fit fails any of the three rules or the non-negative support requirement.


.. py:function:: EGAMMA_TPE_DIAGNOSTICS(low, likely, high, low_probability=0.1, fit_tolerance=2.5E-11, ceiling_tolerance=2E-5, require_nonnegative=FALSE):

    Report on the fit :func:`EGAMMA_TPE_TO_PARAMS` makes for the same
    arguments, including why it failed.

    :return: A list of five values. First the status: ``"converged"`` or
        ``"approximation"`` (a fit at the shape ceiling) for an accepted fit,
        and otherwise ``"ceiling tolerance not met"``,
        ``"precision insufficient"``, ``"reconstruction inconsistent"``,
        ``"nonnegative support not met"``, ``"search failed"`` or
        ``"numerical failure"``. Then the position residual, the reconstructed
        error as computed, and the verification uncertainty, each as a
        fraction of the range; and TRUE if the fitted distribution has support
        below zero. That is always so for a negative scale, whose support is
        unbounded below, and for a positive scale with a negative location.
        Values not computed before a failure are ``#N/A``.
    :raises ValueError: Returns the same error value as
        :func:`EGAMMA_TPE_TO_PARAMS` when an argument is invalid.




