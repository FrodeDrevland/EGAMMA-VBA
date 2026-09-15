# Changelog

## 1.2.0

The three-point fit now searches on the mode's **position** within the elicited
range rather than on the half-range ratio, and stops on an absolute tolerance
rather than a relative one. **Fitted parameters change** in the last few digits
for most inputs, and estimates with a mode very close to an outer value are now
fitted where 1.1.1 could refuse them. Workbooks whose numbers came from 1.1.1
should be recalculated if either matters.

This release matches version 1.2.0 of the companion Python library; the two
produce the same parameters on the published conformance vectors.

Only the fitting routine changes. The distribution functions are untouched, and
so are the moments: `EGAMMA_MEAN` remains `alpha * beta + delta` and
`EGAMMA_STDDEV` remains `SQRT(alpha) * ABS(beta)`, defined from the parameters
alone, as they must be for a distribution whose parameters did not come from a
three-point fit.

### Changed

- **The shape search matches mode position.** The target is
  `min(likely - low, high - likely) / (high - low)` and the model quantity is
  `(alpha - 1 - q_low) / (q_high - q_low)`. The search stops when the two differ
  by less than the tolerance. That difference is itself the maximum normalised
  error in the three reproduced values, so the tolerance is now stated directly
  in the quantity a user cares about rather than implying it through a bound on
  the ratio.
- **`TOLERANCE = 2.5E-11` replaces `THRESHOLD = 1E-10`.** This is not a rename:
  the old value bounded the relative error in the half-range ratio, the new one
  bounds the position error directly. 2.5E-11 is the reproduction accuracy the
  old relative threshold implied, so the intended accuracy is unchanged.

### Fixed

- **A mode very close to an outer value is no longer refused.** The relative
  stopping rule demanded an absolute agreement proportional to the ratio, which
  tends to zero as the mode approaches an outer value, so an estimate such as
  (100, 100.00001, 300) could exhaust the search and return `#NUM!` although it
  is admissible. The absolute rule imposes the same reproduction requirement
  near an outer value as at it.

### Added

- **`EGAMMA_TPE_AT_CEILING(low, likely, high, [low_probability])`**, which is
  `TRUE` when the fit returned the shape ceiling rather than a shape meeting the
  tolerance. An estimate too near symmetry to resolve reproduces the elicited
  values to about 1.5E-5 of the range rather than to the tolerance; returning
  the ceiling silently left no way to tell the two apart.

### Removed

- **The separate endpoint solver.** A mode at an outer value gives a target
  position of zero and goes through the ordinary search, so
  `FindAlphaAtModeEqualsProbability` is gone. There is no longer a different
  code path, or a different tolerance, for the endpoint cases, and
  `EGAMMA_TPE_TO_PARAMS` no longer branches on them.

## 1.1.1

Packaging only. The library is unchanged from 1.1.0; the worksheet functions
and the results they produce are identical.

### Fixed

- **Generated Sphinx output was tracked in the repository.** `docs/build/` was
  listed in `.gitignore`, but had been committed before that rule was added, so
  it stayed tracked and bloated the repository. It is no longer tracked.
- **The release archive carried documentation sources but no readable
  documentation.** It contained the reStructuredText under `docs/source`, which
  is of no use to someone installing an Excel add-in, and nothing rendered. The
  archive now excludes `docs` entirely and ships `EGAMMA-VBA-manual.pdf` at the
  top level instead, alongside the add-in.

### Changed

- The manual's version is now read from `EGAMMA_LIB_VERSION` in `EGAMMA.bas`
  rather than repeated in the Sphinx configuration, so the title page cannot
  claim a version the library does not report.

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
