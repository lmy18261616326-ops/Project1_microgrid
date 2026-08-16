function pvbatt_initialize()
%PVBATT_INITIALIZE Initialize model variables and nonlinear MPC object.

global PVBATT_P %#ok<GVMIS> Shared read-only parameters for nlmpc callbacks.
global PVBATT_PV_VOLTAGE_BP PVBATT_PV_TEMPERATURE_BP %#ok<GVMIS>
global PVBATT_PV_IRRADIANCE_BP PVBATT_PV_CURRENT_TABLE %#ok<GVMIS>

p = pvbatt_parameters();
PVBATT_P = p;

[PVBATT_PV_VOLTAGE_BP, PVBATT_PV_TEMPERATURE_BP, ...
    PVBATT_PV_IRRADIANCE_BP, PVBATT_PV_CURRENT_TABLE] = ...
    pvbatt_generate_pv_lookup(p);

if evalin('base', "exist('ScenarioName','var')")
    scenarioName = evalin('base', 'ScenarioName');
else
    scenarioName = "comparison";
end
if evalin('base', "exist('ControllerSelect','var')")
    controllerSelect = evalin('base', 'ControllerSelect');
else
    controllerSelect = 1;
end

[scenarioW, initialSoc, stopTime] = ...
    pvbatt_build_scenario(scenarioName, p);
wInitial = scenarioW(1, 2:4)';
loadPowerInitial = p.dcBus.referenceVoltage^2 / wInitial(3);
[x0, u0, w0, operatingPoint] = pvbatt_operating_point( ...
    p, initialSoc, wInitial(1), wInitial(2), loadPowerInitial);

nlmpcobjPVBatt = pvbatt_create_nlmpc(x0, u0, w0, p);

variables = {
    'PVBATT_P', p;
    'PV_V_BP', PVBATT_PV_VOLTAGE_BP;
    'PV_T_BP', PVBATT_PV_TEMPERATURE_BP;
    'PV_G_BP', PVBATT_PV_IRRADIANCE_BP;
    'PV_I_TABLE', PVBATT_PV_CURRENT_TABLE;
    'BAT_SOC_BP', p.battery.socBreakpoints;
    'BAT_OCV_TABLE', p.battery.ocvTable;
    'scenario_w', scenarioW;
    'ScenarioName', scenarioName;
    'ControllerSelect', controllerSelect;
    'ScenarioStopTime', stopTime;
    'x0', x0;
    'u0', u0;
    'w0', w0;
    'OperatingPointInfo', operatingPoint;
    'nlmpcobj_pvbatt', nlmpcobjPVBatt;
    'ARIMA_Phi_T', p.arima.temperaturePhi;
    'ARIMA_Phi_G', p.arima.irradiancePhi;
    'ARIMA_Phi_R', p.arima.loadResistancePhi;
    'MPPT_V_init', x0(1);
    'MPPT_P_init', operatingPoint.pvPower;
    'MPPT_V_step', p.mppt.voltageStep;
    'MPPT_V_min', p.mppt.minimumVoltage;
    'MPPT_V_max', p.mppt.maximumVoltage;
    'MPPT_P_deadband', p.mppt.powerDeadband;
    'MPC_OV_weights_normal', [ ...
        p.controller.weights.dcBusVoltage, ...
        p.controller.weights.pvVoltage, 0];
    'MPC_OV_weights_curtail', [ ...
        p.controller.weights.dcBusVoltage, 0, ...
        p.controller.weights.batteryCurrent];
    'PV_PI_Kp', p.pi.pvKp;
    'PV_PI_Ki', p.pi.pvKi;
    'DC_PI_Kp', p.pi.dcKp;
    'DC_PI_Ki', p.pi.dcKi
    };

for k = 1:size(variables, 1)
    assignin('base', variables{k, 1}, variables{k, 2});
end
end
