function voltage = pvbatt_battery_ocv(stateOfCharge)
%PVBATT_BATTERY_OCV Evaluate the replaceable lithium-ion OCV table.

global PVBATT_P %#ok<GVMIS> Shared read-only parameters for the plant callback.
stateOfCharge = min(max(stateOfCharge, 0), 1);
voltage = interp1( ...
    PVBATT_P.battery.socBreakpoints, ...
    PVBATT_P.battery.ocvTable, ...
    stateOfCharge, 'linear');
end
