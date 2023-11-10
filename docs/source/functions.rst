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

    :param alpha: The shape parameter of the expanded gamma distribution.
    :param beta: The scale parameter of the expanded gamma distribution.
    :param delta: The location  parameter of the expanded gamma distribution.
    :return: The mode of the expanded gamma distribution.
    


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
    


    :param alpha: The shape parameter of the expanded gamma distribution.
    :param beta: The scale parameter of the expanded gamma distribution.
    :return: The standard deviation of the expanded gamma distribution.
    


.. py:function:: EGAMMA_SKEW(alpha, beta):
    
    Calculate the skewness of the expanded gamma distribution.

    :param alpha: The shape parameter of the expanded gamma distribution.
    :param beta: The scale parameter of the expanded gamma distribution.
    :return: The skewness of the expanded gamma distribution.
    
.. py:function:: EGAMMA_KURT(alpha):

    Calculate the skewness of the expanded gamma distribution.

    :param alpha: The shape parameter of the expanded gamma distribution.
    :param beta: The scale parameter of the expanded gamma distribution.
    :return: The skewness of the expanded gamma distribution.

# The functions EGAMMA_TPE_TO_PARAMS, FindAlphaAtModeEqualsProbability, FindAlpha, and EGAMMA_FIT_TO_PARAMS are complex and would require similarly detailed docstrings. Each function should have a description of its purpose, parameters, return value, and any exceptions.

.. py:function:: EGAMMA_FIT_TO_PARAMS(*args):
    
    Fit the parameters of the expanded gamma distribution to a given set of data points.

    This function estimates the alpha, beta, and delta parameters of the expanded  gamma distribution,
    based on the provided data using the method of moments. It uses the skewness of the data to estimate alpha, then calculates beta
    using the standard deviation, and finally computes delta from the mean, alpha, and beta.

    :param args: A variable number of arrays or lists containing numerical data points.
                 These are combined into a single dataset for the analysis.
    :return: A list containing the estimated parameters [alpha, beta, delta] of the extended gamma distribution.
    :raises ValueError: If the input data does not allow for a proper fit, e.g., if it leads to divide-by-zero errors.
    

.. py:function:: EGAMMA_TPE_TO_PARAMS(low, likely, high, low_probability=0.1):
    
    Estimate the parameters of the expanded gamma distribution based on a three-point estimate.

    This function uses a low, likely, and high estimate to determine the parameters of the extended gamma distribution.

    :param low: The low estimate of the distribution.
    :param likely: The most likely estimate of the distribution.
    :param high: The high estimate of the distribution.
    :param low_probability: The probability of the low  estimates (default is 0.1).
    :return: A list containing the estimated parameters [alpha, beta, delta] of the extended gamma distribution.
    :raises ValueError: If the input values are inconsistent (e.g., low > likely, high < likely, or all values are equal).
    



