%% PV-ESS MPC paper reproduction: minimum parameter initialization
% Run this script once before opening/running the model.
% It DOES NOT create or modify any Simulink model.
%
% Source classification:
%   [PAPER] values explicitly reported by Hu et al., Applied Energy 221.
%   [ASSUMPTION] values required by the switching model but not reported.

%% 1. Paper values
Vdc_ref = 1200;              % [V] [PAPER] DC-bus reference
Vac_ll = 690;                % [V rms] [PAPER] low-voltage AC bus
f_grid = 50;                 % [Hz] [PAPER]
Cdc = 6e-3;                  % [F] [PAPER]
Lf = 0.6e-3;                 % [H/phase] [PAPER]
Rf = 0.019;                  % [ohm/phase] [PAPER]
Cf = 1338e-6;                % [F/phase equivalent] [PAPER]

P_bat_rated = 0.5e6;         % [W] [PAPER]
V_bat_nom = 300;             % [V] [PAPER]
Q_bat_Ah = 1300;             % [Ah] [PAPER]
SOC0 = 50;                   % [%] [PAPER test initial condition]
SOC_min = 10;                % [%] [PAPER]
SOC_max = 90;                % [%] [PAPER]

Fsw_pv = 5e3;                % [Hz] [PAPER]
Fsw_bat = 5e3;               % [Hz] [PAPER]
Ts_mpc = 1/20e3;             % [s] [PAPER] 20 kHz controller update
m_var = 4e6;                 % [var/pu] [PAPER] Q-V coefficient

P_ac_base = 0.5e6;           % [W] [PAPER]
P_ac_step = 0.5e6;           % [W] [PAPER]
P_dc_1 = 0.6e6;              % [W] [PAPER]
P_dc_2 = 1.0e6;              % [W] [PAPER]

PV_module = 'SunPower SPR-305E-WHT-D'; % [PAPER]
Ns_pv = 10;                  % [ASSUMPTION] gives about 550 V at MPP
Np_pv = 656;                 % [ASSUMPTION] about 2 MW at STC

S_tr = 2.5e6;                % [VA] [PAPER]
V_tr_lv = Vac_ll;            % [V rms line-line]
V_tr_hv = 25e3;              % [V rms line-line] [PAPER]

%% 2. Switching-model engineering assumptions
Ts_power = 2e-6;             % [s] electrical-network discrete sample
Lpv = 0.2e-3;                % [H] PV boost inductor
RLpv = 2e-3;                 % [ohm] PV inductor series resistance
Cpv = 5e-3;                  % [F] PV terminal capacitor
Lbat = 0.5e-3;               % [H] battery converter inductor
RLbat = 2e-3;                % [ohm]
R_bat_internal = 0.02;       % [ohm]
Ron_switch = 1e-3;           % [ohm]
Vf_diode = 1.0;              % [V]

S_sc_grid = 47e6;            % [VA] utility short-circuit level
XR_grid = 10;                % [-] utility X/R ratio

%% 3. PV MPPT values (MPPT itself is built from standard blocks)
Ts_mppt = 1e-3;              % [s] actual MPPT decision period
Dpv0 = 1 - 547/Vdc_ref;      % [-] initial boost duty, about 0.544
Dpv_min = 0.35;
Dpv_max = 0.75;
Dpv_step = 2e-4;
Dpv_off_mppt = 0.46;         % [-] curtailment starting point; tune experimentally
eps_dV = 1e-3;               % [V] zero-change dead band
eps_dI = 1e-3;               % [A] zero-change dead band
eps_inc = 1e-4;              % [A/V] conductance dead band

%% 4. Battery cascaded PI values
% Positive Ibat means battery discharging into the DC bus.
Ts_bat_ctrl = 100e-6;        % [s]
Ibat_max = P_bat_rated/V_bat_nom;
Db0 = 1 - V_bat_nom/Vdc_ref;
Db_min = 0.05;
Db_max = 0.95;
Kpv_bat = 2.0;               % DC-voltage outer loop
Kiv_bat = 100.0;
Kpi_bat = 2e-4;              % battery-current inner loop
Kii_bat = 0.02;

%% 5. Grid-tied DC-bus outer PI and FCS normalization
Kp_dc = 2500;
Ki_dc = 2e5;
Pref_min = -2.2e6;
Pref_max = 2.2e6;
P_base_mpc = 2e6;
Q_base_mpc = 1e6;
lambda_sw = 0.002;           % optional switching-change penalty
omega_grid = 2*pi*f_grid;

%% 6. Initial conditions
Vdc0 = Vdc_ref;
Vpv0 = 547;
ILpv0 = 0;
Ibat0 = 0;

%% 7. Exact discrete MPVC matrices from the paper state model
% x = [Vc; If], y = [Vi; Iload], x(k+1) = Ad*x(k) + Bd*y(k)
A_mpvc = [0, 1/Cf; -1/Lf, -Rf/Lf];
B_mpvc = [0, -1/Cf; 1/Lf, 0];
Ad_mpvc = expm(A_mpvc*Ts_mpc);
Bd_mpvc = A_mpvc\((Ad_mpvc-eye(2))*B_mpvc);

%% 8. Paper event times and From Workspace profiles
T_ac_step = 0.5;
T_pv_ramp_up = 1.0;
T_sync = 1.4;
T_grid_connect = 1.6;
T_dc_2_on = 2.5;
T_pv_ramp_down = 3.0;
T_grid_dip = 3.5;

G_profile = timeseries([400;400;800;800;400], ...
    [0;T_pv_ramp_up;2.0;T_pv_ramp_down;4.0]);

% Controlled resistor values: R = Vdc^2/P. 1e6 ohm means nearly open.
Rdc1_profile = timeseries([1e6;1e6;Vdc_ref^2/P_dc_1], ...
    [0;T_pv_ramp_up;T_pv_ramp_up+Ts_power]);
Rdc2_profile = timeseries([1e6;1e6;Vdc_ref^2/P_dc_2; ...
    Vdc_ref^2/P_dc_2;1e6], ...
    [0;T_dc_2_on;T_dc_2_on+Ts_power;T_pv_ramp_down;T_pv_ramp_down+Ts_power]);
GridVpu_profile = timeseries([1;1;0.9], ...
    [0;T_grid_dip;T_grid_dip+Ts_power]);

%% 9. Convenient derived values used directly in block dialogs
Vac_phase_peak = sqrt(2)*Vac_ll/sqrt(3);
Qc_filter = 2*pi*f_grid*Cf*Vac_ll^2;
Rdc_0p6MW = Vdc_ref^2/P_dc_1;
Rdc_1MW = Vdc_ref^2/P_dc_2;

disp('pvess_beginner_init: parameters and profiles are ready.');
