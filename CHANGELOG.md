# Changelog

## 1.1.0

Corrections to the fitting and distribution functions. The three-point fit
already produced correct parameters in 1.0; the changes below concern error
handling, robustness and one runtime fault.

### Fixed

- **`EGAMMA_DIST` raised a type mismatch instead of returning an error
  value.** The function was declared `As Double` but assigned `CVErr(...)` for
  a zero scale parameter and outside the support. A Double-typed function
  cannot hold an error value. `EGAMMA_DIST` and `EGAMMA_INV` are now
  `As Variant`.
- **The density returned `#N/A` outside the support** rather than zero, which
  was also inconsistent with the cumulative function, which correctly returned
  0 or 1 there.
- **Both shape searches returned a "best effort" value on non-convergence.**
  After exhausting their iterations they returned the last candidate, so a fit
  that had not converged was indistinguishable from one that had. They now
  report failure, and `EGAMMA_TPE_TO_PARAMS` propagates it as `#NUM!`.
- **The symmetry shortcut left a band of unreachable targets.** It fired at a
  fixed ratio of 0.99999, while the search ceiling of 1e9 produces a maximum
  ratio of 0.99994. Estimates falling between the two passed the shortcut and
  could not be bracketed, so they ran the full 200 iterations of `GAMMA.INV`
  before returning an unconverged value. The shortcut is now derived from the
  ceiling, so those estimates return immediately.

### Changed

- The endpoint search stops on the residual normalised by the percentile span
  rather than comparing values rounded to ten decimals.
- `EGAMMA_TPE_TO_PARAMS` validates `low_probability`, returning `#NUM!` unless
  `0 < low_probability < 0.5`.
- `EGAMMA_DIST` and `EGAMMA_INV` validate the shape and scale parameters.
- `EGAMMA_FIT_TO_PARAMS` treats a near-zero sample skewness as symmetric
  rather than testing for exact floating-point zero, and caps the shape
  parameter at the library ceiling.
- `ALPHA_MAX`, `THRESHOLD` and `MAX_ITER` are named constants rather than
  literals repeated through the code.

### Added

- `EGAMMA_VERSION()`, so a workbook can record which build produced its
  numbers.

### Documentation

- `EGAMMA_STDEV` corrected to `EGAMMA_STDDEV`; the documented name did not
  exist, so anyone copying it got `#NAME?`.
- `EGAMMA_KURT` documented as excess kurtosis.
- The mathematical documentation gains the derivation of the three-point fit,
  the admissible range of estimates, and the accuracy bound. It was previously
  a copy of the Python library's page, and carried two malformed math roles
  that rendered as literal text.
- The overview explains what Excel's own `GAMMA.DIST` cannot do and why, and
  states what inputs the fit will reject.

## 1.0

Initial release.
