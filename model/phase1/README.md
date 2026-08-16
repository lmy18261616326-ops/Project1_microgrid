# Phase 1 averaged plant component

`Plant_Derivatives.m` implements the six-state averaged PV-battery system
from the paper. It is intended to be called from a Simulink MATLAB Function
block:

```matlab
[dx, aux] = Plant_Derivatives(x, u, w, Phase1Params);
```

Initialize parameters and an electrically consistent operating point with:

```matlab
Phase1Params = phase1_parameters;
[x0, u0, w0, info] = phase1_operating_point( ...
    Phase1Params, 7.2e3, 0.8, 25, 1000);
```

The first four and sixth entries of `info.electricalDerivative` should be
near zero. Its SoC derivative is intentionally nonzero when the battery is
charging or discharging.

The 9-series, 5-parallel PV array, environmental coefficients, simplified
battery resistance, and constant open-circuit voltage are explicit phase-1
assumptions. They should be replaced when the missing datasheet or author
parameters become available.

For a self-contained version that can be pasted directly into a Simulink
MATLAB Function block, use `Plant_Derivatives_Block.m`. It has only three
inputs (`x`, `u`, and `w`) and contains no external function calls.
