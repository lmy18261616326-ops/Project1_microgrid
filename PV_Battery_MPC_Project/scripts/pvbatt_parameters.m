function p = pvbatt_parameters()
%PVBATT_PARAMETERS Parameters for the PV-battery MPC reproduction.
% Values explicitly published in Batiyah et al. (2022) are marked
% "published". Values not disclosed by the paper are isolated under the
% assumption fields so that they can be replaced without editing the model.

%% Published converter and simulation parameters
p.pvConverter.capacitance = 300e-6;
p.pvConverter.inductance = 10e-3;
p.pvConverter.inductorResistance = 10e-3;

p.batteryConverter.capacitance = 300e-6;
p.batteryConverter.inductance = 10e-3;
p.batteryConverter.inductorResistance = 10e-3;

p.dcBus.capacitance = 1500e-6;
p.dcBus.referenceVoltage = 600;

p.sim.sampleTime = 10e-6;
p.controller.interval = 10e-3;
p.controller.predictionHorizon = 1;
p.controller.predictionSubsteps = 10;
p.converter.switchingFrequency = 10e3;

%% Published PV module and array ratings
p.pv.moduleMaximumPower = 213.15;
p.pv.moduleMaximumPowerVoltage = 29;
p.pv.moduleMaximumPowerCurrent = 7.35;
p.pv.moduleOpenCircuitVoltage = 36.3;
p.pv.moduleShortCircuitCurrent = 7.84;
p.pv.referenceTemperatureC = 25;
p.pv.referenceIrradiance = 1000;

% The paper states a 9.5 kW array but does not publish the module layout.
% Nine modules in series and five strings in parallel give 9.59175 kW.
p.pv.seriesModules = 9;
p.pv.parallelStrings = 5;
p.pv.cellsPerModule = 60;

% Datasheet-calibrated one-diode parameters at STC. These values are fitted
% from the four published module points plus the MPP slope condition.
p.pv.diodeIdeality = 1.13950820159814;
p.pv.seriesResistance = 0.327685084019587;
p.pv.shuntResistance = 4972.06961578068;
p.pv.photocurrentSTC = 7.84000084543376;
p.pv.saturationCurrentSTC = 8.30470768095633e-9;

% Not reported by the paper. Typical crystalline-silicon assumptions.
p.pv.shortCircuitTempCoefficient = 0.0005 ...
    * p.pv.moduleShortCircuitCurrent;
p.pv.bandGapElectronVolt = 1.12;

%% Published battery ratings and isolated reproduction assumptions
p.battery.nominalVoltage = 300;
p.battery.capacityAh = 20;
p.battery.minimumStateOfCharge = 0.20;
p.battery.maximumStateOfCharge = 0.90;

% E0, K, A, B, and Rbat are not reported in the paper. The plant therefore
% uses a replaceable OCV-versus-SoC table plus an internal resistance.
p.battery.internalResistance = 0.08;
p.battery.socBreakpoints = [0, 0.10, 0.20, 0.50, 0.80, 0.90, 1.00];
p.battery.ocvTable = [270, 282, 290, 296, 300, 303, 312];

%% Published MPC formulation
p.controller.weights.dcBusVoltage = 0.75;
p.controller.weights.pvVoltage = 0.15;
p.controller.weights.dutyRate = [0.10, 0.10];
p.controller.weights.batteryCurrent = 0.15;
p.controller.minimumDuty = 0;
p.controller.maximumDuty = 1;

% Practical numerical limits used only to keep the reproduction robust.
p.controller.appliedMinimumDuty = 0.02;
p.controller.appliedMaximumDuty = 0.95;

% Standard-block Mode III constraint-enforcement gains. This outer layer
% removes the local-solution chatter of a horizon-one nonlinear optimizer.
p.controller.curtailBatteryCurrentGain = 2.0e-4;
p.controller.curtailPvVoltageGain = 2.0e-4;

%% P&O voltage-reference generator
p.mppt.sampleTime = p.controller.interval;
p.mppt.initialVoltage = p.pv.seriesModules ...
    * p.pv.moduleMaximumPowerVoltage;
p.mppt.voltageStep = 0.25;
p.mppt.minimumVoltage = 180;
p.mppt.maximumVoltage = 315;
p.mppt.powerDeadband = 1.0;

%% ARIMA one-step predictor assumptions
% The paper identifies ARIMA forecasting but does not publish orders or
% coefficients. Zero AR coefficients implement the robust ARIMA(0,1,0)
% persistence baseline. Nonzero values enable the visible ARIMA(1,1,0)
% branches in the model for later calibration against measured histories.
p.arima.temperaturePhi = 0;
p.arima.irradiancePhi = 0;
p.arima.loadResistancePhi = 0;

%% PI baseline settings
% The paper publishes state-feedback gains but omits the full realization.
% These gains implement a transparent two-loop PI comparison model.
p.pi.pvKp = 2.0e-3;
p.pi.pvKi = 0.20;
p.pi.dcKp = 1.0e-3;
p.pi.dcKi = 0.05;

p.limits.minimumLoadResistance = 1;
p.limits.minimumIrradiance = 1e-3;
end
