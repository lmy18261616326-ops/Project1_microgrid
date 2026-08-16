function [dx, aux] = Plant_Derivatives(x, u, w, p)
%PLANT_DERIVATIVES Six-state averaged PV-battery plant dynamics.
%   [DX, AUX] = PLANT_DERIVATIVES(X, U, W, P) evaluates the continuous-time
%   model used in the first reproduction phase of Batiyah et al. (2022).
%
%   States X:
%     X(1) = PV terminal voltage, V
%     X(2) = PV boost-inductor current, A
%     X(3) = battery terminal voltage, V
%     X(4) = battery-converter inductor current, A
%     X(5) = battery state of charge, 0...1
%     X(6) = DC-bus voltage, V
%
%   Control inputs U:
%     U(1) = PV boost duty ratio
%     U(2) = battery-converter duty ratio
%
%   Disturbances W:
%     W(1) = cell temperature, degC
%     W(2) = solar irradiance, W/m^2
%     W(3) = load resistance, ohm
%
%   AUX channels:
%     1 PV array current, A
%     2 battery current, A (positive while discharging)
%     3 battery internal voltage, V
%     4 applied PV duty ratio
%     5 applied battery duty ratio
%     6 PV source power, W
%     7 PV power injected into the DC bus, W
%     8 battery power injected into the DC bus, W
%     9 load power, W
%    10 DC-bus power-balance residual, W
%
%   The battery model is intentionally simplified for phase 1. Replace
%   BATTERY.OCVSLOPE and related parameters when validated cell data is
%   available. The corrected SoC equation uses battery terminal voltage
%   X(3), not inductor current X(4) as printed in Equation (12).
%
%#codegen

assert(numel(x) == 6, "Plant_Derivatives:InvalidStateSize");
assert(numel(u) == 2, "Plant_Derivatives:InvalidInputSize");
assert(numel(w) == 3, "Plant_Derivatives:InvalidDisturbanceSize");

pvVoltage = x(1);
pvInductorCurrent = x(2);
batteryTerminalVoltage = x(3);
batteryInductorCurrent = x(4);
stateOfCharge = x(5);
dcBusVoltage = x(6);

pvDuty = min(max(u(1), p.limits.minimumDuty), ...
    p.limits.maximumDuty);
batteryDuty = min(max(u(2), p.limits.minimumDuty), ...
    p.limits.maximumDuty);

temperatureC = w(1);
irradiance = max(w(2), 0);
loadResistance = max(w(3), p.limits.minimumLoadResistance);

pvCurrent = phase1_pv_current( ...
    pvVoltage, temperatureC, irradiance, p);

boundedStateOfCharge = min(max(stateOfCharge, 0), 1);
batteryInternalVoltage = p.battery.nominalVoltage ...
    + p.battery.ocvSlope ...
    * (boundedStateOfCharge - p.battery.nominalStateOfCharge);
batteryCurrent = (batteryInternalVoltage ...
    - batteryTerminalVoltage) / p.battery.internalResistance;

dx = zeros(6, 1);
dx(1) = (pvCurrent - pvInductorCurrent) ...
    / p.pvConverter.capacitance;
dx(2) = (-p.pvConverter.inductorResistance ...
    * pvInductorCurrent + pvVoltage ...
    - (1 - pvDuty) * dcBusVoltage) ...
    / p.pvConverter.inductance;
dx(3) = (batteryCurrent - batteryInductorCurrent) ...
    / p.batteryConverter.capacitance;
dx(4) = (-p.batteryConverter.inductorResistance ...
    * batteryInductorCurrent + batteryTerminalVoltage ...
    - (1 - batteryDuty) * dcBusVoltage) ...
    / p.batteryConverter.inductance;
dx(5) = -batteryCurrent / (3600 * p.battery.capacityAh);
dx(6) = ((1 - pvDuty) * pvInductorCurrent ...
    + (1 - batteryDuty) * batteryInductorCurrent ...
    - dcBusVoltage / loadResistance) ...
    / p.dcBus.capacitance;

pvSourcePower = pvVoltage * pvCurrent;
pvBusPower = dcBusVoltage * (1 - pvDuty) * pvInductorCurrent;
batteryBusPower = dcBusVoltage ...
    * (1 - batteryDuty) * batteryInductorCurrent;
loadPower = dcBusVoltage^2 / loadResistance;
dcBusStoredPower = p.dcBus.capacitance * dcBusVoltage * dx(6);
powerBalanceResidual = pvBusPower + batteryBusPower ...
    - loadPower - dcBusStoredPower;

aux = zeros(10, 1);
aux(1) = pvCurrent;
aux(2) = batteryCurrent;
aux(3) = batteryInternalVoltage;
aux(4) = pvDuty;
aux(5) = batteryDuty;
aux(6) = pvSourcePower;
aux(7) = pvBusPower;
aux(8) = batteryBusPower;
aux(9) = loadPower;
aux(10) = powerBalanceResidual;
end
