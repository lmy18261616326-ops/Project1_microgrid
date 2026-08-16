function p = phase1_parameters()
%PHASE1_PARAMETERS Parameters for the averaged phase-1 plant model.
%   Values marked as assumptions fill gaps in the paper and should remain
%   tunable until module and battery datasheets are available.

p.pvConverter.capacitance = 300e-6;
p.pvConverter.inductance = 10e-3;
p.pvConverter.inductorResistance = 10e-3;

p.batteryConverter.capacitance = 300e-6;
p.batteryConverter.inductance = 10e-3;
p.batteryConverter.inductorResistance = 10e-3;

p.dcBus.capacitance = 1500e-6;
p.dcBus.referenceVoltage = 600;

% Published PV module data.
p.pv.moduleMaximumPower = 213.15;
p.pv.moduleMaximumPowerVoltage = 29;
p.pv.moduleMaximumPowerCurrent = 7.35;
p.pv.moduleOpenCircuitVoltage = 36.3;
p.pv.moduleShortCircuitCurrent = 7.84;
p.pv.referenceTemperatureC = 25;
p.pv.referenceIrradiance = 1000;

% Phase-1 array assumption: 9 series modules and 5 parallel strings.
p.pv.seriesModules = 9;
p.pv.parallelStrings = 5;

% Phase-1 environmental assumptions. The temperature coefficients are
% typical placeholder values, not values reported in the paper.
p.pv.shortCircuitCurrentTempCoefficient = ...
    0.0005 * p.pv.moduleShortCircuitCurrent;
p.pv.openCircuitVoltageTempCoefficient = ...
    -0.0032 * p.pv.moduleOpenCircuitVoltage;
p.pv.openCircuitVoltageLogCoefficient = 2.0;
p.pv.minimumIrradianceRatio = 1e-6;
p.pv.minimumModuleOpenCircuitVoltage = 1;

currentRatioAtMpp = p.pv.moduleMaximumPowerCurrent ...
    / p.pv.moduleShortCircuitCurrent;
voltageRatioAtMpp = p.pv.moduleMaximumPowerVoltage ...
    / p.pv.moduleOpenCircuitVoltage;
p.pv.curveExponent = log(1 - currentRatioAtMpp) ...
    / log(voltageRatioAtMpp);

% Published battery ratings plus phase-1 assumptions.
p.battery.nominalVoltage = 300;
p.battery.capacityAh = 20;
p.battery.nominalStateOfCharge = 0.8;
p.battery.internalResistance = 0.05;
p.battery.ocvSlope = 0;

p.limits.minimumDuty = 0;
p.limits.maximumDuty = 1;
p.limits.minimumLoadResistance = 1;
p.limits.minimumActiveIrradiance = 1e-6;
end
