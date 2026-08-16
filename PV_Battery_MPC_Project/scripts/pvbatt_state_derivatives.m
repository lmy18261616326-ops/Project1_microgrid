function dx = pvbatt_state_derivatives(x, duty, disturbance)
%PVBATT_STATE_DERIVATIVES Six-state averaged model from paper Equation 12.
% The published SoC line contains x4 where the battery terminal voltage x3
% is dimensionally required. This implementation follows Equations 6 and 9.

global PVBATT_P %#ok<GVMIS> Shared read-only parameters for the plant callback.
p = PVBATT_P;

vpv = x(1);
iLpv = x(2);
vb = x(3);
iLb = x(4);
soc = x(5);
vdc = x(6);

dpv = min(max(duty(1), p.controller.minimumDuty), ...
    p.controller.maximumDuty);
db = min(max(duty(2), p.controller.minimumDuty), ...
    p.controller.maximumDuty);

temperatureC = disturbance(1);
irradiance = max(disturbance(2), 0);
loadResistance = max(disturbance(3), ...
    p.limits.minimumLoadResistance);

ipv = pvbatt_pv_current(vpv, temperatureC, irradiance);
ebat = pvbatt_battery_ocv(soc);
ib = (ebat - vb) / p.battery.internalResistance;

dx = zeros(6, 1);
dx(1) = (ipv - iLpv) / p.pvConverter.capacitance;
dx(2) = (-p.pvConverter.inductorResistance * iLpv ...
    + vpv - (1 - dpv) * vdc) ...
    / p.pvConverter.inductance;
dx(3) = (ib - iLb) / p.batteryConverter.capacitance;
dx(4) = (-p.batteryConverter.inductorResistance * iLb ...
    + vb - (1 - db) * vdc) ...
    / p.batteryConverter.inductance;
dx(5) = -ib / (3600 * p.battery.capacityAh);
dx(6) = ((1 - dpv) * iLpv ...
    + (1 - db) * iLb ...
    - vdc / loadResistance) / p.dcBus.capacitance;
end
