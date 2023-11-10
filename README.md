# EGAMMA-VBA: Expanded Gamma Distribution Library for Excel

Welcome to EGAMMA-VBA, an open-source Excel library offering a VBA implementation of the expanded gamma distribution. This library is designed to enhance Excel's capabilities in statistical analysis by incorporating a more versatile gamma distribution model.

## Overview
EGAMMA-VBA implements the expanded gamma distribution,  a variant of the traditional three-parameter gamma distribution also allowing for left-skewed data. This library extends Excel's native gamma distribution functions, providing additional flexibility and precision in statistical modeling. It is especially useful for scenarios involving expert judgment for three-point estimation, allowing users to define custom percentile ranges for optimistic and pessimistic scenarios.

For Python users, EGAMMA-VBA complements the `egamma` Python library. More information on the Python version can be found see the [egamma](https://frodedrevland.github.io/egamma) github repository.

## Features
- **Enhanced Distribution Functions**: EGAMMA_DIST and EGAMMA_INV, improved versions of Excel's GAMMA.DIST and GAMMA.INV.
- **Statistical Measures**: Functions for calculating mean, mode, median, variance, standard deviation, skewness, and kurtosis.
- **Parameter Estimation**: Functions for data fitting and three-point estimation.
  - **Data Fitting**: EGAMMA_FIT_TO_PARAMS fits the expanded gamma distribution to your data.
  - **Three-Point Estimation**: EGAMMA_TPE_TO_PARAMS calculates distribution parameters from a three-point estimate.
- **Comprehensive Documentation**: Detailed guides for each function, including examples and parameter descriptions.

## Installation
Install EGAMMA-VBA by downloading the EGAMMA.xlam add-in or the EGAMMA.bas file for direct code module import into your VBA project. Detailed installation instructions are available at [EGAMMA-VBA Installation Guide](https://frodedrevland.github.io/EGAMMA-VBA/installation/index.html).

## Usage
Upon installation, EGAMMA-VBA provides a suite of new functions in Excel. For detailed usage instructions, visit [EGAMMA-VBA Usage Guide](https://frodedrevland.github.io/EGAMMA-VBA/usage.html#).

## Documentation
Find in-depth documentation on the expanded gamma distribution and library functions at [EGAMMA-VBA Documentation](https://frodedrevland.github.io/EGAMMA-VBA).

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
