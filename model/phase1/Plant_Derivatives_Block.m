function [dx, aux] = Plant_Derivatives_Block(x, u, w)
%PLANT_DERIVATIVES_BLOCK Paste-ready code for a MATLAB Function block.
%   Inputs:
%     x = [vpv; iLpv; vb; iLb; SoC; vdc]
%     u = [dpv; db]
%     w = [temperature_degC; irradiance_Wm2; Rload_ohm]
%
%   Outputs:
%     dx  = six continuous-time state derivatives
%     aux = [Ipv; Ib; Ebat; dpv; db; Ppv; PpvBus; ...
%            PbatBus; Pload; powerBalanceResidual]
%
%   Current signs:
%     Ib > 0 and PbatBus > 0 mean battery discharge.
%     Ib < 0 and PbatBus < 0 mean battery charge.
%
%#codegen

%% Published converter and DC-bus parameters
Cpv = 300e-6;
Lpv = 10e-3;
rLpv = 10e-3;

Cb = 300e-6;
Lb = 10e-3;
rLb = 10e-3;

Cdc = 1500e-6;

%% Published battery ratings and phase-1 assumptions
batteryCapacityAh = 20;
batteryNominalVoltage = 300;
batteryInternalResistance = 0.05;  % Phase-1 assumption
batteryNominalSoC = 0.8;
batteryOcvSlope = 0;               % Constant OCV in phase 1

%% Published PV-module data
moduleVmp = 29;
moduleImp = 7.35;
moduleVoc = 36.3;
moduleIsc = 7.84;

%% Phase-1 PV-array and environmental assumptions
seriesModules = 9;
parallelStrings = 5;
referenceTemperatureC = 25;
referenceIrradiance = 1000;
iscTempCoefficient = 0.0005 * moduleIsc;
vocTempCoefficient = -0.0032 * moduleVoc;
vocLogCoefficient = 2.0;
minimumIrradianceRatio = 1e-6;

currentRatioAtMpp = moduleImp / moduleIsc;
voltageRatioAtMpp = moduleVmp / moduleVoc;
curveExponent = log(1 - currentRatioAtMpp) ...
    / log(voltageRatioAtMpp);

%% Unpack states, commands, and disturbances
vpv = x(1);
iLpv = x(2);
vb = x(3);
iLb = x(4);
soc = x(5);
vdc = x(6);

dpv = min(max(u(1), 0), 1);
db = min(max(u(2), 0), 1);

temperatureC = w(1);
irradiance = max(w(2), 0);
Rload = max(w(3), 1);

%% Empirical PV I-V curve
if irradiance <= 1e-6
    Ipv = 0;
else
    temperatureOffset = temperatureC - referenceTemperatureC;
    irradianceRatio = irradiance / referenceIrradiance;

    moduleIscNow = moduleIsc ...
        + iscTempCoefficient * temperatureOffset;
    arrayIsc = parallelStrings * irradianceRatio ...
        * max(moduleIscNow, 0);

    moduleVocNow = moduleVoc ...
        + vocTempCoefficient * temperatureOffset ...
        + vocLogCoefficient ...
        * log(max(irradianceRatio, minimumIrradianceRatio));
    moduleVocNow = max(moduleVocNow, 1);
    arrayVoc = seriesModules * moduleVocNow;

    normalizedVoltage = max(vpv, 0) / arrayVoc;
    if normalizedVoltage >= 1
        Ipv = 0;
    else
        Ipv = arrayIsc ...
            * (1 - normalizedVoltage^curveExponent);
    end
end

%% Simplified phase-1 battery model
boundedSoC = min(max(soc, 0), 1);
Ebat = batteryNominalVoltage ...
    + batteryOcvSlope * (boundedSoC - batteryNominalSoC);
Ib = (Ebat - vb) / batteryInternalResistance;

%% Six-state averaged plant equations
dx = zeros(6, 1);
dx(1) = (Ipv - iLpv) / Cpv;
dx(2) = (-rLpv * iLpv + vpv ...
    - (1 - dpv) * vdc) / Lpv;
dx(3) = (Ib - iLb) / Cb;
dx(4) = (-rLb * iLb + vb ...
    - (1 - db) * vdc) / Lb;

% Corrected form of the paper's Equation (12): SoC depends on battery
% current, equivalently on Ebat - vb, not on inductor current x(4).
dx(5) = -Ib / (3600 * batteryCapacityAh);

dx(6) = ((1 - dpv) * iLpv ...
    + (1 - db) * iLb ...
    - vdc / Rload) / Cdc;

%% Diagnostic outputs
Ppv = vpv * Ipv;
PpvBus = vdc * (1 - dpv) * iLpv;
PbatBus = vdc * (1 - db) * iLb;
Pload = vdc^2 / Rload;
dcBusStoredPower = Cdc * vdc * dx(6);
powerBalanceResidual = PpvBus + PbatBus ...
    - Pload - dcBusStoredPower;

aux = zeros(10, 1);
aux(1) = Ipv;
aux(2) = Ib;
aux(3) = Ebat;
aux(4) = dpv;
aux(5) = db;
aux(6) = Ppv;
aux(7) = PpvBus;
aux(8) = PbatBus;
aux(9) = Pload;
aux(10) = powerBalanceResidual;
end
