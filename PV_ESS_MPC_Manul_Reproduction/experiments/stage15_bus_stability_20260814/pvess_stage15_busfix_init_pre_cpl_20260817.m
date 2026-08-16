%% PV-ESS MPC paper reproduction: minimum parameter initialization
% Run this script once before opening/running the model.
% It DOES NOT create or modify any Simulink model.
%
% Source classification:
%   [PAPER] values explicitly reported by Hu et al., Applied Energy 221.
%   [ASSUMPTION] values required by the switching model but not reported.

%% 1. initial values
Vdc_ref = 1200;              % [V] [PAPER] DC-bus reference
Vac_ll = 690;                % [V rms] [PAPER] low-voltage AC bus
f_grid = 50;                 % [Hz] [PAPER]
Cdc_paper = 6e-3;            % [F] [PAPER] reported DC-bus capacitance
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

P_ac_base_standalone = 0.5e6;% [W] standalone 0.5 -> 1.0 MW regression case
P_ac_base = 1;               % [W] full-paper case starts unloaded; SPS load requires nonzero P
P_ac_step = 0.5e6;           % [W] [PAPER] switched in at 0.5 s
P_dc_1 = 0.6e6;              % [W] [PAPER]
P_dc_2 = 1.0e6;              % [W] [PAPER]

PV_module = 'SunPower SPR-305E-WHT-D'; % [PAPER]
Ns_pv = 10;                  % [ASSUMPTION] gives about 550 V at MPP
Np_pv = 656;                 % [ASSUMPTION] about 2 MW at STC

S_tr = 2.5e6;                % [VA] [PAPER]
V_tr_lv = Vac_ll;            % [V rms line-line]
V_tr_hv = 25e3;              % [V rms line-line] [PAPER]
Iac_rated_lv = S_tr/(sqrt(3)*V_tr_lv); % [A rms] converter/transformer LV side
Igrid_rated_hv = S_tr/(sqrt(3)*V_tr_hv); % [A rms] retained for any future HV-side measurement
% The existing grid_VI/Vabc_grid signal is on the 690 V side, so the
% executable limiter must use Iac_rated_lv and an LV-side voltage floor.
Vgrid_limit_floor = 0.1*V_tr_lv; % [V rms LL] numerical floor at measured LV bus
lambda_ilim = 100;           % optional MPPC predicted-current penalty; disabled initially
PQLimit_eps = 1e3;           % [VA] limit-active detection tolerance

%% 2. Switching-model engineering assumptions
Ts_power = 2e-6;             % [s] electrical-network discrete sample
Cdc = 18e-3;                 % [F] tuned switching-model DC-bus capacitor
Lpv = 0.8e-3;                % [H] PV boost inductor
RLpv = 2e-3;                 % [ohm] PV inductor series resistance
Cpv = 10e-3;                 % [F] tuned PV terminal capacitor
% [STAGE15 FIX] 11.25 mH limited the charge-to-discharge current reversal
% to about 27-39 ms, while Vdc lost inverter voltage margin in about 14 ms.
% The 0.5 mH value is the original switching-model engineering assumption
% and passed the 0.5 MW AC load-step validation in the experiment copy.
Lbat = 0.5e-3;                % [H] battery converter inductor
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
Dpv_curtail_bias = 0.506;    % [-] smooth curtailment duty at nominal DC-bus voltage
Kp_pv_curtail = 5e-5;       % [duty/V] low-gain correction avoids PV/DC-link limit cycling
Dpv_slew_up = 1.0;           % [duty/s] bumpless return from curtailment to MPPT
Dpv_slew_down = -5.0;        % [duty/s] faster protective power reduction
Vdc_curtail_on = 1220;       % [V] enable PV curtailment after battery charge saturation
Vdc_curtail_off = 1185;      % [V] release on a real load-induced bus dip
eps_dV = 1e-3;               % [V] zero-change dead band
eps_dI = 1e-3;               % [A] zero-change dead band
eps_inc = 1e-4;              % [A/V] conductance dead band

%% 4. Battery cascaded PI values
% Positive Ibat means battery discharging into the DC bus.
Ts_bat_ctrl = 100e-6;        % [s]
Ibat_max = P_bat_rated/V_bat_nom;
Ibat_curtail_release = -0.8*Ibat_max; % [A] release once charging demand leaves saturation margin
Db0 = 1 - V_bat_nom/Vdc_ref;
Db_min = 0.05;
Db_max = 0.95;
Kpv_bat = 4;             % DC-voltage outer loop, retuned for Cdc=18 mF
Kiv_bat = 650;
% [STAGE15 FIX] With Kp=0.1, a 2.5 A current error saturated the +/-0.25
% duty correction. These gains retain a fast inner loop without relay-like
% duty toggling under normal tracking errors.
Kpi_bat = 5e-3;               % battery-current inner-loop proportional gain
Kii_bat = 0.05;               % battery-current inner-loop integral gain

%% 4a. Engineering hardening: dynamic battery limits and cross-layer tracking
% [ENGINEERING] These are executable protection limits, not waveform clamps.
% Positive battery current discharges into the DC bus; negative current charges.
Pbat_dis_max = P_bat_rated;    % [W] continuous discharge terminal-power limit
Pbat_chg_max = P_bat_rated;    % [W] continuous charge terminal-power limit
Ibat_dis_hw = Ibat_max;        % [A] positive hardware-current limit
Ibat_chg_hw = Ibat_max;        % [A] charge-current magnitude limit
Vbat_div_floor = 270;          % [V] safe denominator; matches existing battery low-end parameter
fc_vbat_limit = 200;           % [Hz] Vbat filter used only by feedforward/limit calculation
a_vbat_limit = exp(-2*pi*fc_vbat_limit*Ts_bat_ctrl);

% Smooth SOC derating replaces the former abrupt SOC MATLAB Function.
% Discharge is blocked at/below 10% and fully restored at 15%.
% Charge is fully available through 85% and blocked at/above 90%.
SOC_dis_bp = [0 10 15 100];
SOC_dis_k  = [0  0  1   1];
SOC_chg_bp = [0 85 90 100];
SOC_chg_k  = [1  1  0   0];

Ilim_eps = 1;                 % [A] limit-active detection tolerance
Kt_bat = 143;                 % [1/s] tracking gain, aligned with existing outer-loop Kb
Dcorr_min = -0.25;            % [-] retained conservative negative duty correction
Dcorr_max = Db_max - Db0;     % [-] +0.20; matches the actual final duty ceiling

% Monitoring-only warning/trip thresholds. A trip is never used to make a
% voltage waveform look better; if it operates, the scenario is reported failed.
Ibat_warn_factor = 1.05;
Ibat_trip_factor = 1.20;

%% 5. Grid-tied DC-bus outer PI and FCS normalization
Kp_dc = 3836.125;
Ki_dc = 3.40748e5;
Pref_min = -2.2e6;
Pref_max = 2.2e6;
P_base_mpc = 2e6;
Q_base_mpc = 1e6;
lambda_mpvc = 3e-6;      % optional switching-change penalty
lambda_mppc = 0.001;
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
Ts_ems = 1e-3;
Ts_soc = 1e-4;
N_ems_fast_avg = 100;
K_grid_sign = -1;
K_bat_sign = 1;
Vsync_threshold = 10;
sync_hold       = 0.1;
phase_threshold = 0.99863;
freq_threshold  = 0.15;

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

stage15_fix_id = 'ff_sign_lbat_current_pi_ac_phase_pv_curtail_dynamic_limits_tracking_20260816';
disp('pvess_stage15_busfix_init: corrected parameters and profiles are ready.');
