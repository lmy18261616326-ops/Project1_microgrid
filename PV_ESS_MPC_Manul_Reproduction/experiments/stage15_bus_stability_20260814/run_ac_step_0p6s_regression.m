experimentDirectory = fileparts(mfilename('fullpath'));
modelFile = fullfile(experimentDirectory, 'manul_model_stage15_complete.slx');
[~, modelName] = fileparts(modelFile);
load_system(modelFile);

toWorkspaceBlocks = find_system(modelName, ...
    'LookUnderMasks', 'all', ...
    'FollowLinks', 'on', ...
    'BlockType', 'ToWorkspace');
originalDecimation = cell(size(toWorkspaceBlocks));
for blockIndex = 1:numel(toWorkspaceBlocks)
    originalDecimation{blockIndex} = ...
        get_param(toWorkspaceBlocks{blockIndex}, 'Decimation');
    set_param(toWorkspaceBlocks{blockIndex}, 'Decimation', '80');
end

simulationInput = Simulink.SimulationInput(modelName);
simulationInput = simulationInput.setVariable('P_ac_base', 0.5e6);
simulationInput = simulationInput.setVariable('P_ac_step', 0.5e6);
simulationInput = simulationInput.setVariable('T_ac_step', 0.6);
simulationInput = simulationInput.setModelParameter( ...
    'StartTime', '0', ...
    'StopTime', '0.9', ...
    'SignalLogging', 'off', ...
    'SaveOutput', 'off', ...
    'SaveState', 'off');

simulationTimer = tic;
simulationOutput = sim(simulationInput);
simulationElapsedSeconds = toc(simulationTimer);

for blockIndex = 1:numel(toWorkspaceBlocks)
    set_param(toWorkspaceBlocks{blockIndex}, 'Decimation', ...
        originalDecimation{blockIndex});
end

vdc = simulationOutput.get('V_dc');
vdcTime = vdc.Time(:);
vdcData = reshape(vdc.Data, [], 1);
preMask = vdcTime >= 0.52 & vdcTime < 0.59;
transientMask = vdcTime >= 0.60 & vdcTime < 0.75;
steadyMask = vdcTime >= 0.82 & vdcTime <= 0.90;

loadVoltage = simulationOutput.get('Vabc_load');
voltageSteadyMask = loadVoltage.Time >= 0.82 & loadVoltage.Time <= 0.90;
phaseRms = sqrt(mean(loadVoltage.Data(voltageSteadyMask, :).^2, 1));
phaseImbalancePercent = ...
    (max(phaseRms) - min(phaseRms))/mean(phaseRms)*100;

interlinkPower = simulationOutput.get('Pf_threephase');
powerSteadyMask = interlinkPower.Time >= 0.82 & ...
    interlinkPower.Time <= 0.90;
batteryPower = simulationOutput.get('P_bat');
batteryPowerSteadyMask = batteryPower.Time >= 0.82 & ...
    batteryPower.Time <= 0.90;

summary = struct( ...
    'simulationElapsedSeconds', simulationElapsedSeconds, ...
    'preMeanVdc_V', mean(vdcData(preMask)), ...
    'preStdVdc_V', std(vdcData(preMask)), ...
    'transientMinVdc_V', min(vdcData(transientMask)), ...
    'transientMaxVdc_V', max(vdcData(transientMask)), ...
    'steadyMeanVdc_V', mean(vdcData(steadyMask)), ...
    'steadyStdVdc_V', std(vdcData(steadyMask)), ...
    'steadyMinVdc_V', min(vdcData(steadyMask)), ...
    'steadyMaxVdc_V', max(vdcData(steadyMask)), ...
    'steadyPhaseRms_V', phaseRms, ...
    'phaseImbalancePercent', phaseImbalancePercent, ...
    'steadyInterlinkPower_W', mean(reshape( ...
        interlinkPower.Data(powerSteadyMask), [], 1)), ...
    'steadyBatteryPower_W', mean(reshape( ...
        batteryPower.Data(batteryPowerSteadyMask), [], 1)), ...
    'batteryPowerMinimum_W', min(reshape(batteryPower.Data, [], 1)), ...
    'batteryPowerMaximum_W', max(reshape(batteryPower.Data, [], 1)), ...
    'allFinite', all(isfinite(vdcData)));

regression = struct();
regression.summary = summary;
regression.signals.V_dc = vdc;
regression.signals.Vabc_load = loadVoltage;
regression.signals.Pf_threephase = interlinkPower;
regression.signals.P_bat = batteryPower;
regression.signals.P_pv = simulationOutput.get('P_pv');
regression.signals.I_ref_protected = simulationOutput.get('I_ref_protected');
regression.signals.I_bat_mean = simulationOutput.get('I_bat_mean');
regression.signals.I_dis_limit_dyn = simulationOutput.get('I_dis_limit_dyn');
regression.signals.I_chg_limit_dyn = simulationOutput.get('I_chg_limit_dyn');
regression.signals.bat_limit_active = simulationOutput.get('bat_limit_active');
regression.signals.bat_power_limit_active = simulationOutput.get('bat_power_limit_active');
regression.signals.Vbat_safe_limit = simulationOutput.get('Vbat_safe_limit');
regression.signals.I_PI_track = simulationOutput.get('I_PI_track');
regression.signals.P_ref_raw = simulationOutput.get('P_ref_raw');
regression.signals.Q_ref_raw = simulationOutput.get('Q_ref_raw');
regression.signals.S_grid_current_max = simulationOutput.get('S_grid_current_max');
regression.signals.grid_current_limit_active = simulationOutput.get('grid_current_limit_active');

batLimit = regression.signals.bat_limit_active;
gridLimit = regression.signals.grid_current_limit_active;
summary.batLimitDuration_s = trapz(batLimit.Time, double(batLimit.Data(:) > 0.5));
summary.gridLimitDuration_s = trapz(gridLimit.Time, double(gridLimit.Data(:) > 0.5));

save(fullfile(experimentDirectory, 'ac_step_0p6s_regression.mat'), ...
    'regression', '-v7.3');
summaryHandle = fopen(fullfile(experimentDirectory, ...
    'ac_step_0p6s_summary.json'), 'w');
assert(summaryHandle ~= -1, 'Unable to create AC-step summary JSON.');
fprintf(summaryHandle, '%s', jsonencode(summary, 'PrettyPrint', true));
fclose(summaryHandle);

set(groot, 'defaultFigureVisible', 'off');
regressionFigure = figure('Color', 'w', 'Position', [50 50 1200 720]);
regressionLayout = tiledlayout(2, 1, 'TileSpacing', 'compact', ...
    'Padding', 'compact');
nexttile(regressionLayout);
plot(vdcTime, vdcData, 'LineWidth', 1.1);
xline(0.6, ':r', '0.5 MW load step');
yline(1200, '--k');
grid on;
ylabel('V_{dc} (V)');
title('Standalone AC-load step regression');

nexttile(regressionLayout);
plot(interlinkPower.Time, reshape(interlinkPower.Data, [], 1)/1e6, ...
    'LineWidth', 1.0);
hold on;
plot(batteryPower.Time, reshape(batteryPower.Data, [], 1)/1e6, ...
    'LineWidth', 1.0);
hold off;
xline(0.6, ':r');
grid on;
xlabel('Time (s)');
ylabel('Power (MW)');
legend({'Interlink AC power', 'Battery power (+discharge)'}, ...
    'Location', 'best');

exportgraphics(regressionFigure, fullfile(experimentDirectory, ...
    'ac_step_0p6s_regression.png'), 'Resolution', 180);
close(regressionFigure);

try
    sdiRunId = Simulink.sdi.createRun( ...
        'stage15_ac_step_0p6s_regression', 'vars', ...
        vdc, interlinkPower, batteryPower, regression.signals.P_pv);
    sdiRun = Simulink.sdi.getRun(sdiRunId);
    fprintf('SDI_RUN_ID=%d SIGNAL_COUNT=%d\n', ...
        sdiRunId, sdiRun.SignalCount);
catch sdiException
    warning('Stage15:ACStepSDI', 'SDI creation failed: %s', ...
        sdiException.message);
end

fprintf('AC_STEP_0P6S_OK elapsed=%.1f s\n', simulationElapsedSeconds);
fprintf('VDC pre=%.3f +/- %.3f V\n', ...
    summary.preMeanVdc_V, summary.preStdVdc_V);
fprintf('VDC transient min=%.3f max=%.3f V\n', ...
    summary.transientMinVdc_V, summary.transientMaxVdc_V);
fprintf('VDC steady=%.3f +/- %.3f V\n', ...
    summary.steadyMeanVdc_V, summary.steadyStdVdc_V);
fprintf('PHASE_RMS=[%.3f %.3f %.3f] imbalance=%.5f%%\n', ...
    phaseRms, phaseImbalancePercent);
fprintf('POWER interlink=%.1f W battery=%.1f W\n', ...
    summary.steadyInterlinkPower_W, summary.steadyBatteryPower_W);
fprintf('PBAT_RANGE=[%.1f %.1f] W BAT_LIMIT_DURATION=%.6f s GRID_LIMIT_DURATION=%.6f s\n', ...
    summary.batteryPowerMinimum_W, summary.batteryPowerMaximum_W, ...
    summary.batLimitDuration_s, summary.gridLimitDuration_s);

close_system(modelName, 0);
