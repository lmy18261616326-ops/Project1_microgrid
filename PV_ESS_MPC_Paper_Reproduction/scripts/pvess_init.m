function cfg = pvess_init(assignToBase)
%PVESS_INIT Parameters for Hu et al. (Applied Energy 221, 2018).
%   CFG = PVESS_INIT(false) returns a structure without modifying the base
%   workspace. PVESS_INIT(true) also assigns every field to the base
%   workspace. Paper values and engineering assumptions are separated.

if nargin < 1
    assignToBase = true;
end

%% Paper values (Table 1 and Section 6)
cfg.Vdc_ref = 1200;
cfg.Vac_ll = 690;
cfg.f_grid = 50;
cfg.Cdc = 6e-3;
cfg.Lf = 0.6e-3;
cfg.Cf = 1338e-6;
cfg.Rf = 0.019;
cfg.P_bat_rated = 0.5e6;
cfg.V_bat_nom = 300;
cfg.Q_bat_Ah = 1300;
cfg.SOC0 = 50;
cfg.SOC_min = 10;
cfg.SOC_max = 90;
cfg.Fsw_pv = 5e3;
cfg.Fsw_bat = 5e3;
cfg.Ts_mpc = 1/20e3;
cfg.m_var = 4e6;
cfg.P_ac_base = 0.5e6;
cfg.P_ac_step = 0.5e6;
cfg.P_dc_1 = 0.6e6;
cfg.P_dc_2 = 1.0e6;
cfg.P_pv_stc = 2.0e6;

% Exact PV module named by the paper. 10 modules in series reproduce the
% approximately 550 V MPP / 640 V open-circuit voltage in Fig. 3.
cfg.PV_module = 'SunPower SPR-305E-WHT-D';
cfg.Ns_pv = 10;
cfg.Np_pv = 656;

%% Engineering assumptions (not reported in the paper)
% These values make the switching template numerically tractable. Replace
% them with measured hardware values before claiming quantitative identity.
cfg.Ts_power = 2e-6;
cfg.Lpv = 0.2e-3;
cfg.RLpv = 2e-3;
cfg.Cpv = 5e-3;
cfg.Lbat = 0.5e-3;
cfg.RLbat = 2e-3;
cfg.R_bat_internal = 0.02;
cfg.Ron_switch = 1e-3;
cfg.Vf_diode = 1.0;
cfg.R_grid_lv = 0.01;
cfg.L_grid_lv = 0.1e-3;
cfg.lambda_sw = 0.002;
cfg.P_base_mpc = 2e6;
cfg.Q_base_mpc = 1e6;

% PV MPPT implementation parameters.
cfg.Ts_mppt = 1e-3;
cfg.Dpv0 = 1 - 547/cfg.Vdc_ref;
cfg.Dpv_min = 0.35;
cfg.Dpv_max = 0.75;
cfg.Dpv_step = 2e-4;
cfg.Dpv_off_mppt = 0.46;

% Battery cascaded controller. Current is positive while discharging.
cfg.Ts_bat_ctrl = 100e-6;
cfg.Kpv_bat = 2.0;
cfg.Kiv_bat = 100.0;
cfg.Kpi_bat = 2e-4;
cfg.Kii_bat = 0.02;
cfg.Ibat_max = cfg.P_bat_rated/cfg.V_bat_nom;
cfg.Db0 = 1 - cfg.V_bat_nom/cfg.Vdc_ref;
cfg.Db_min = 0.05;
cfg.Db_max = 0.95;

% DC-bus outer PI for grid-tied MPPC.
cfg.Kp_dc = 2500;
cfg.Ki_dc = 2e5;
cfg.Pref_min = -2.2e6;
cfg.Pref_max = 2.2e6;

% Initial values used by Specialized Power Systems.
cfg.Vdc0 = cfg.Vdc_ref;
cfg.Vpv0 = 547;
% Zero-current start avoids opening/closing an initialized high-current
% inductor during staged commissioning. Enable steady-state initialization
% only after every source-side isolator is closed.
cfg.ILpv0 = 0;
cfg.Ibat0 = 0;
cfg.PV_blocking = 0;
cfg.Battery_blocking = 0;
cfg.Inverter_enable = 1;
cfg.PV_connect = 1;
cfg.Battery_connect = 1;
cfg.Inverter_connect = 1;

% Transformer and source equivalent used by the switching template.
cfg.S_tr = 2.5e6;
cfg.V_tr_lv = cfg.Vac_ll;
cfg.V_tr_hv = 25e3;
cfg.S_sc_grid = 47e6;
cfg.XR_grid = 10;

%% Paper event schedule (Table 2)
cfg.T_ac_step = 0.5;
cfg.T_pv_ramp_up = 1.0;
cfg.T_dc_1 = 1.0;
cfg.T_sync = 1.4;
cfg.T_grid_connect = 1.6;
cfg.T_dc_2_on = 2.5;
cfg.T_pv_ramp_down = 3.0;
cfg.T_dc_2_off = 3.0;
cfg.T_grid_dip = 3.5;

% Timeseries used by From Workspace blocks. Duplicate time stamps create
% ideal steps; the irradiance profile uses linear ramps.
cfg.G_profile = timeseries([400;400;800;800;400], ...
    [0;cfg.T_pv_ramp_up;2.0;cfg.T_pv_ramp_down;4.0]);
cfg.Rdc1_profile = timeseries([1e6;1e6;cfg.Vdc_ref^2/cfg.P_dc_1], ...
    [0;cfg.T_dc_1;cfg.T_dc_1 + cfg.Ts_power]);
cfg.Rdc2_profile = timeseries([1e6;1e6;cfg.Vdc_ref^2/cfg.P_dc_2; ...
    cfg.Vdc_ref^2/cfg.P_dc_2;1e6], ...
    [0;cfg.T_dc_2_on;cfg.T_dc_2_on + cfg.Ts_power; ...
    cfg.T_dc_2_off;cfg.T_dc_2_off + cfg.Ts_power]);
cfg.GridVpu_profile = timeseries([1;1;0.9], ...
    [0;cfg.T_grid_dip;cfg.T_grid_dip + cfg.Ts_power]);

if assignToBase
    names = fieldnames(cfg);
    for k = 1:numel(names)
        assignin('base',names{k},cfg.(names{k}));
    end
end
end
