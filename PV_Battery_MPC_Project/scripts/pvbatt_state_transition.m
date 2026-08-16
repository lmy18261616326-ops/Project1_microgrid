function xNext = pvbatt_state_transition(x, combinedInput)
%PVBATT_STATE_TRANSITION Discrete nonlinear model for the nlmpc object.
% combinedInput = [dpv; db; temperature; irradiance; loadResistance].

global PVBATT_P %#ok<GVMIS> Shared read-only parameters for the nlmpc callback signature.
p = PVBATT_P;

duty = combinedInput(1:2);
disturbance = combinedInput(3:5);
step = p.controller.interval / p.controller.predictionSubsteps;
xNext = x;

for k = 1:p.controller.predictionSubsteps
    [xStage, k1] = localReducedDerivative( ...
        xNext, duty, disturbance, p);
    [~, k2] = localReducedDerivative( ...
        xStage + 0.5 * step * k1, duty, disturbance, p);
    [~, k3] = localReducedDerivative( ...
        xStage + 0.5 * step * k2, duty, disturbance, p);
    [~, k4] = localReducedDerivative( ...
        xStage + step * k3, duty, disturbance, p);
    xNext = xNext + step / 6 * (k1 + 2 * k2 + 2 * k3 + k4);
    xNext(3) = pvbatt_battery_ocv(xNext(5)) ...
        - p.battery.internalResistance * xNext(4);
end

xNext(5) = min(max(xNext(5), 0), 1);
end

function [conditionedState, dx] = localReducedDerivative( ...
        x, duty, disturbance, p)
% The battery terminal capacitor settles much faster than the 10 ms control
% interval. Enforcing its algebraic quasi-steady value keeps the horizon-one
% predictor numerically well-conditioned while the plant retains all six
% published continuous states at the 10 us simulation step.

conditionedState = x;
conditionedState(3) = pvbatt_battery_ocv(x(5)) ...
    - p.battery.internalResistance * x(4);
dx = pvbatt_state_derivatives(conditionedState, duty, disturbance);
dx(3) = 0;
end
