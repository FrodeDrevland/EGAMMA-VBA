.. _usage:

Usage
=====

Once the EGAMMA plugin is installed, its functions become accessible within your Excel workbook(s). This section offers a general overview of how to use these functions.

EGAMMA_DIST and EGAMMA_INV Functions
------------------------------------
The EGAMMA_DIST and EGAMMA_INV functions are analogous to Excel's native GAMMA.DIST and GAMMA.INV functions, respectively. The key enhancements in the EGAMMA versions include an additional `delta` parameter for location adjustment, and the ability for the `beta` parameter to accept negative values, accommodating negatively skewed distributions.

For instance, to calculate the cumulative probability for a value of 200 in an expanded distribution characterized by alpha = 10, beta = 20, and delta = 100, use:

.. code-block:: none

    =EGAMMA_DIST(200, 10, 20, 100, TRUE)

Statistical Measures
--------------------
The EGAMMA library includes a range of functions designed to compute various statistical measures based on the parameters of the expanded gamma distribution (alpha, beta, and delta). These functions operate similarly to standard Excel functions:

- ``EGAMMA_MEAN(alpha, beta, delta)``: Calculates the mean.
- ``EGAMMA_MODE(alpha, beta, delta)``: Calculates the mode.
- ``EGAMMA_MEDIAN(alpha, beta, delta)``: Calculates the median.
- ``EGAMMA_VAR(alpha, beta)``: Calculates the variance.
- ``EGAMMA_STDDEV(alpha, beta)``: Calculates the standard deviation.
- ``EGAMMA_SKEW(alpha, beta)``: Calculates the skewness.
- ``EGAMMA_KURT(alpha)``: Calculates the *excess* kurtosis.

Parameter estimation
--------------------

The EGAMMA library offers two distinct functions for parameter estimation: EGAMMA_TPE_TO_PARAMS and EGAMMA_FIT_TO_PARAMS. These functions are designed to calculate the parameters of the Expanded Gamma Distribution either from a three-point estimate or by fitting to a data set.

EGAMMA_TPE_TO_PARAMS
"""""""""""""""""""""
The EGAMMA_TPE_TO_PARAMS function requires three key values as arguments: the low, most likely, and high values of a three-point estimate. Additionally, it accepts an optional low probability value, which is the probability of the low estimate (defaulting to 0.1). It's important to note that the probability of the high estimate is always the complement of the low probability (1 - low probability).

.. code-block:: none

    =EGAMMA_TPE_TO_PARAMS(low, most-likely, high, [low-probability])
    Example: =EGAMMA_TPE_TO_PARAMS(100, 200, 400, 0.1)

The function outputs a 3x1 array: alpha in the function's cell, and beta and delta in the two adjacent cells to the right. If these adjacent cells are not empty, the function will return ``#SPILL!``, indicating it cannot display the results.

Not every three-point estimate can be fitted. The mode of the fitted
distribution cannot fall outside the interval the two elicited percentiles
enclose, which limits how lopsided an estimate may be: at the default
``low_probability`` of 0.1 the skewness cannot exceed about 1.86 in magnitude.
An estimate that violates ``low <= most-likely <= high`` returns ``#N/A``; a
``low-probability`` outside the open interval (0, 0.5), or a fit that cannot
converge, returns ``#NUM!``.

A perfectly symmetric estimate, where the mode sits exactly midway, is a
special case: no finite shape parameter matches it exactly, so the function
returns a very large one. The elicited values are then reproduced to about 15
parts per million of the elicited range rather than exactly. For any other
admissible estimate they are reproduced to within about 2.5e-11 of the range.

EGAMMA_FIT_TO_PARAMS
"""""""""""""""""""""
EGAMMA_FIT_TO_PARAMS is designed to estimate the alpha, beta, and delta parameters of the Expanded Gamma Distribution from a given data set. This function uses the method of moments, employing the skewness of the data to estimate alpha, the standard deviation for beta, and the mean, alpha, and beta to calculate delta.

The ``egamma`` Python library implements the same estimator as
``fit(data, method='mom')``, so results from the two agree. Its default,
``method='mle'``, uses maximum likelihood instead and will generally give
slightly different parameters from the same sample.

.. code-block:: none

    =EGAMMA_FIT_TO_PARAMS(range)
    Example: =EGAMMA_FIT_TO_PARAMS(A1:A100)

Similar to EGAMMA_TPE_TO_PARAMS, the results are returned as an array. The alpha parameter appears in the cell where the function is entered, with beta and delta in the adjacent right cells.



For further details, consult the :ref:`functions` documentation.


