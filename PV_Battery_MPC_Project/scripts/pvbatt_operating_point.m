function [x0, u0, w0, info] = pvbatt_operating_point( ...
        p, initialStateOfCharge, temperatureC, irradiance, loadPower)
%PVBATT_OPERATING_POINT Construct an electrically consistent initial state.

global PVBATT_P %#ok<GVMIS> Initialize shared parameters used by helper callbacks.
PVBATT_P = p;

vdc = p.dcBus.referenceVoltage;
loadResistance = vdc^2 / loadPower;

voltageGrid = linspace(0, 1.02 ...
    * p.pv.seriesModules * p.pv.moduleOpenCircuitVoltage, 600);
currentGrid = arrayfun(@(v) pvbatt_pv_current( ...
    v, temperatureC, irradiance), voltageGrid);
[pvPower, maximumIndex] = max(voltageGrid .* currentGrid);
vpv = voltageGrid(maximumIndex);
ipv = currentGrid(maximumIndex);
iLpv = ipv;

oneMinusDpv = (vpv ...
    - p.pvConverter.inductorResistance * iLpv) / vdc;
dpv = 1 - oneMinusDpv;

requiredBatteryPower = loadPower - pvPower;
ebat = pvbatt_battery_ocv(initialStateOfCharge);
combinedResistance = p.battery.internalResistance ...
    + p.batteryConverter.inductorResistance;
discriminant = ebat^2 ...
    - 4 * combinedResistance * requiredBatteryPower;
if discriminant <= 0
    error("PVBATT:OperatingPointUnavailable", ...
        "The requested load exceeds the assumed battery capability.");
end

iLb = (ebat - sqrt(discriminant)) / (2 * combinedResistance);
vb = ebat - p.battery.internalResistance * iLb;
oneMinusDb = (vb ...
    - p.batteryConverter.inductorResistance * iLb) / vdc;
db = 1 - oneMinusDb;

x0 = [vpv; iLpv; vb; iLb; initialStateOfCharge; vdc];
u0 = [dpv; db];
w0 = [temperatureC; irradiance; loadResistance];

info.pvPower = pvPower;
info.loadPower = loadPower;
info.requiredBatteryPower = requiredBatteryPower;
info.arrayRatedPower = p.pv.seriesModules ...
    * p.pv.parallelStrings * p.pv.moduleMaximumPower;
end
