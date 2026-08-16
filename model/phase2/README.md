# Phase 2 discrete MPPT control

`PV_MPPT_v2.slx` extends the phase-1 averaged PV-battery plant with a
discrete perturb-and-observe (P&O) controller for the PV converter.

The battery converter duty ratio remains fixed at its phase-1 operating
point. DC-bus regulation is intentionally deferred to a later phase.

## Signal flow

1. Select `Vpv = x(1)` and `Ipv = aux(1)`.
2. Sample both measurements every 2 ms.
3. Calculate PV power and compare it with the previous sample.
4. Reverse the duty perturbation direction when power decreases by more
   than the configured threshold.
5. Apply the duty perturbation and limit the command to `[0.05, 0.90]`.

The phase-1 state and auxiliary logs are retained. Phase 2 adds
`phase2_duty` and `phase2_power` workspace logs.

## Tuned parameters

- MPPT sample time: `2e-3 s`
- Duty perturbation: `5e-4`
- Initial duty: `0.5656125`
- Duty limits: `[0.05, 0.90]`

## Initial verification targets

- All state, power, and duty signals remain finite.
- Duty remains inside its configured limits.
- At 1000 W/m2 and 25 degC, PV operation remains close to
  `Vmp = 261 V`, `Imp = 36.75 A`, and `Pmp = 9591.75 W`.

## Validation result

An 8 s simulation with the existing load-resistance step from 50 ohm to
25 ohm at 4 s completed with finite signals.

- Pre-step mean PV power: `9603.75 W`
- Final mean PV power: `9603.78 W`
- Final mean DC-bus voltage: `597.18 V`
- Overall duty range: `[0.5471125, 0.5666125]`

The DC-bus transient and final offset are expected because the battery
converter duty ratio is still fixed in this phase.
