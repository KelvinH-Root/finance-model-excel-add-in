# Monte Carlo and one-way cash flow proof (Phase 0)

Proves the spec's simulation design before any add-in code: a seeded Monte Carlo that runs in plain Excel, with no add-in, macros, custom functions or Python in Excel, over a one-way cash flow model driven by a scenario events register. The EXL Cloud demos had both (a scenario manager with Monte Carlo bounds and a distribution report; a one-way cash flow model with spread, percentage, repeat, bullet and start-and-end events). This is HFG's own version of the pattern.

```
python prototypes/montecarlo/simulation.py            # build/montecarlo/mc_demo.xlsx, 2,000 trials
python prototypes/montecarlo/simulation.py --trials 500 out.xlsx
pytest tests/test_montecarlo.py                       # 8 tests; the LibreOffice ones skip without it
```

Open the workbook in Excel and press F9. Calculation is set to automatic except tables, so the simulation runs only when asked. The Checks sheet compares the results with the Python reference written into the file at build time.

## What is in the workbook

| Sheet | Holds |
|---|---|
| Control | Seed, trials (fixed at build), inspect trial, cash settings |
| Risk | Risk register: PERT, triangular, normal, discrete and Bernoulli drivers, base values and live values |
| Correlation | Target correlation and its Cholesky factor as plain formulas, with a validity check |
| Draws | Per trial and driver: two uniforms from the HDR generator, Latin hypercube rank, normal score, correlated score, sampled value |
| Events | Scenario events register: bullet, repeat, spread, start and end, percentage of an earlier event with a lag; scale, shift and only-if links to risk drivers |
| Cash | Monthly engine: base cash, one row per event, a facility that draws to hold the cash floor and repays from surplus, results (peak debt and month, lowest cash and month, interest, breach) |
| Run | One data table over the trial number: 8 results plus closing debt for each month |
| Stats | Mean, SD, percentiles, breach probability, histogram, fan chart, convergence checkpoints |
| Sensitivity | Rank correlation with peak debt and approximate contribution to variance |
| Checks | Error checks, alerts, and the comparison with the build's reference |

## How it works

- **Seeded draws.** Each uniform is a pure function of (trial, variable, seed) using Hubbard Decision Research's published counter-based generator (2019 form). No RAND, so the same seed gives the same numbers in any Excel, and any trial can be replayed on its own.
- **Latin hypercube.** A second stream orders the trials; each draw lands in its own 1/N slice of the distribution.
- **Correlation.** Normal scores are multiplied by the Cholesky factor (a Gaussian copula). If the matrix is not valid the factor fails, an error check fires and the draws fall back to uncorrelated.
- **One data table.** Each row reruns the whole model with that trial's draws. The draws sit outside the loop, so only the model recalculates per trial.
- **Inspect trial.** Type a trial number on Control and the whole visible model shows that trial, row by row.
- **Self-check.** Expected values from the Python reference are stored in Checks; a mismatch is an error check.

## Verified

- In LibreOffice, every one of 200 trials matches the Python reference to 1e-7 for all 32 table columns; the workbook's checks pass; rank sensitivity matches; inspect trial replays a trial exactly.
- LibreOffice imports only the first two columns of a wide one-variable data table, so `hfgmodels.verify.recalculate` rebuilds the full table operation before calculating. Excel does not need this.
- Not yet verified in Excel: speed per trial, whether Excel for the web calculates an existing data table, whether a data table survives sheet insertion from a file. The probe add-in tests all three (Monte Carlo data table card).

## Open points

- The HDR generator is published by Hubbard Decision Research; confirm its terms before Phase 1. The design only needs some pure function of (trial, variable, entity, seed), so a different counter-based generator can replace it.
- The Gaussian copula slightly disturbs the Latin hypercube strata for correlated drivers; Iman-Conover keeps them exact and is a later option.
