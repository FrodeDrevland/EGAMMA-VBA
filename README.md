# EGAMMA-VBA: Expanded Gamma Distribution Library for Excel

Welcome to EGAMMA-VBA, an open-source Excel library offering a VBA implementation of the expanded gamma distribution. This library is designed to enhance Excel's capabilities in statistical analysis by incorporating a more versatile gamma distribution model.

Turn an expert's **low / most likely / high** estimate into a probability
distribution you can use in a spreadsheet model — including the left-skewed
estimates that Excel's own `GAMMA.DIST` cannot represent.

```
=EGAMMA_TPE_TO_PARAMS(100, 140, 300)
```

returns the shape, scale and location parameters of the distribution whose
mode is exactly 140, whose 10th percentile is exactly 100, and whose 90th
percentile is exactly 300. The numbers the estimator gave you come back out.

## Overview

Excel's `GAMMA.DIST` has two limitations for this kind of work. It has no
location parameter, so the distribution always starts at zero, and its scale
parameter must be positive, so the distribution is always right-skewed. An
estimate whose most likely value sits closer to the high end than the low end
cannot be represented at all.

EGAMMA-VBA adds both: a location parameter, and a scale parameter whose sign
sets the direction of skew. One family then covers estimates leaning either
way. It also adds what Excel has no equivalent of, fitting to a three-point
estimate where the outer values are elicited as percentiles rather than as
absolute bounds, with the percentile convention configurable.

The family itself is not new. It is the Pearson Type III distribution, long
established in hydrology and elsewhere. What this library provides is that
family in Excel, in the shape–scale–location parameters already familiar from
`GAMMA.DIST`.

For Python users, EGAMMA-VBA complements the `egamma` Python library. More
information on the Python version can be found at the
[egamma](https://frodedrevland.github.io/egamma) repository. The two share the same three-point fitting procedure. This library fits to
data by the method of moments; the Python library offers that estimator too,
as `fit(data, method='mom')`, so results from the two agree.

## Features
- **Enhanced Distribution Functions**: EGAMMA_DIST and EGAMMA_INV, improved versions of Excel's GAMMA.DIST and GAMMA.INV.
- **Statistical Measures**: Functions for calculating mean, mode, median, variance, standard deviation, skewness, and kurtosis.
- **Parameter Estimation**: Functions for data fitting and three-point estimation.
  - **Data Fitting**: EGAMMA_FIT_TO_PARAMS fits the expanded gamma distribution to your data.
  - **Three-Point Estimation**: EGAMMA_TPE_TO_PARAMS calculates distribution parameters from a three-point estimate.
- **Comprehensive Documentation**: Detailed guides for each function, including examples and parameter descriptions.

### Things worth knowing

Not every three-point estimate can be fitted. The fitted mode cannot fall
outside the interval your two percentiles enclose, which limits how lopsided
an estimate can be: at the default 10th/90th convention the skewness cannot
exceed about 1.86 in magnitude. Anything more extreme returns `#N/A`.

A perfectly symmetric estimate is a special case, since no finite
gamma-shaped distribution is exactly symmetric. The function returns a very
large shape parameter; the elicited values then come back to within about 15
parts per million of the range rather than exactly. Otherwise they are
reproduced to within about 2.5e-11 of the range.

## Installation
Install EGAMMA-VBA by downloading the EGAMMA.xlam add-in or the EGAMMA.bas file for direct code module import into your VBA project. Detailed installation instructions are available at [EGAMMA-VBA Installation Guide](https://frodedrevland.github.io/EGAMMA-VBA/installation/index.html).

## Usage
Upon installation, EGAMMA-VBA provides a suite of new functions in Excel. For detailed usage instructions, visit [EGAMMA-VBA Usage Guide](https://frodedrevland.github.io/EGAMMA-VBA/usage.html#).

## Documentation
Find in-depth documentation on the expanded gamma distribution and library functions at [EGAMMA-VBA Documentation](https://frodedrevland.github.io/EGAMMA-VBA).

The same documentation is included as a PDF,
[EGAMMA-VBA-manual.pdf](EGAMMA-VBA-manual.pdf), so that it travels with the
add-in: it is present in every release archive and in the archived deposit made
from it, and can be read offline. It is generated from the sources under
`docs/`, and its title page reports the same version as `EGAMMA_VERSION()`.

## Support
For assistance or to report issues, please use the [GitHub Issue Tracker](https://github.com/FrodeDrevland/EGAMMA-VBA/issues).

## Contact Information
For inquiries or collaboration, contact:

- **Associate Professor Frode Drevland**
- **Affiliation**: Norwegian University of Science and Technology (NTNU)
- **Email**: [frode.drevland@ntnu.no](mailto:frode.drevland@ntnu.no)

Dr. Drevland is committed to the ongoing development of the EGAMMA library and welcomes community engagement.

## License
EGAMMA-VBA is available under the MIT License. For more details, see the [LICENSE](https://github.com/FrodeDrevland/EGAMMA-VBA/blob/main/LICENSE).
