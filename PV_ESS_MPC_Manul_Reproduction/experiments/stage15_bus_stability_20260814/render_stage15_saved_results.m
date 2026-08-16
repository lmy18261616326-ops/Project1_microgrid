experimentDirectory = fileparts(mfilename('fullpath'));
load(fullfile(experimentDirectory, 'stage15_complete_results.mat'), 'results');
set(groot, 'defaultFigureVisible', 'off');

eventTimes = [0.5 1.0 1.4 1.6 2.5 3.0 3.5];
vdc = results.signals.V_dc;

overviewFigure = figure('Color', 'w', 'Position', [50 50 1400 900]);
overviewLayout = tiledlayout(3, 1, 'TileSpacing', 'compact', ...
    'Padding', 'compact');

nexttile(overviewLayout);
plot(vdc.Time, reshape(vdc.Data, [], 1), 'LineWidth', 1.0);
yline(1200, '--k', '1200 V');
for eventTime = eventTimes
    xline(eventTime, ':', 'Color', [0.5 0.5 0.5]);
end
grid on;
ylabel('V_{dc} (V)');
title('Stage 15 complete scenario: DC-link voltage');

nexttile(overviewLayout);
powerSignals = {'P_pv', 'P_bat', 'P_DC', ...
    'Pf_threephase', 'P_grid_threephase'};
powerLabels = {'PV', 'Battery (+discharge)', 'DC load', ...
    'Interlink AC', 'Grid (+export)'};
hold on;
for signalIndex = 1:numel(powerSignals)
    powerSignal = results.signals.(powerSignals{signalIndex});
    plot(powerSignal.Time, reshape(powerSignal.Data, [], 1)/1e6, ...
        'LineWidth', 0.9);
end
hold off;
for eventTime = eventTimes
    xline(eventTime, ':', 'Color', [0.75 0.75 0.75]);
end
grid on;
ylabel('Power (MW)');
legend(powerLabels, 'Location', 'eastoutside');
title('Active-power balance');

nexttile(overviewLayout);
reactiveSignals = {'Q_ref_cmd', 'Qf_threephase', 'Q_grid_threephase'};
reactiveLabels = {'Q reference', 'Interlink Q', 'Grid Q'};
hold on;
for signalIndex = 1:numel(reactiveSignals)
    reactiveSignal = results.signals.(reactiveSignals{signalIndex});
    plot(reactiveSignal.Time, reshape(reactiveSignal.Data, [], 1)/1e6, ...
        'LineWidth', 0.9);
end
hold off;
for eventTime = eventTimes
    xline(eventTime, ':', 'Color', [0.75 0.75 0.75]);
end
grid on;
xlabel('Time (s)');
ylabel('Reactive power (Mvar)');
legend(reactiveLabels, 'Location', 'eastoutside');
title('Q-V droop response');

exportgraphics(overviewFigure, fullfile(experimentDirectory, ...
    'stage15_complete_overview.png'), 'Resolution', 180);
close(overviewFigure);

gridVoltage = results.signals.Vabc_grid;
gridPreMask = gridVoltage.Time >= 3.35 & gridVoltage.Time < 3.45;
gridTimeStep = median(diff(gridVoltage.Time));
samplesPerCycle = max(2, round(0.02/gridTimeStep));
gridMovingRms = sqrt(movmean(gridVoltage.Data.^2, ...
    samplesPerCycle, 1));
gridRmsAverage = mean(gridMovingRms, 2);
gridNominalRms = mean(gridRmsAverage(gridPreMask));

dipFigure = figure('Color', 'w', 'Position', [50 50 1200 700]);
dipLayout = tiledlayout(2, 1, 'TileSpacing', 'compact', ...
    'Padding', 'compact');
nexttile(dipLayout);
plot(gridVoltage.Time, gridRmsAverage/gridNominalRms, 'LineWidth', 1.1);
xline(3.5, ':r', 'Voltage dip');
yline(0.9, '--k', '0.9 pu source command');
xlim([3.3 4.0]);
ylim([0.82 1.05]);
grid on;
ylabel('PCC voltage (pu RMS)');
title('Physical grid-voltage dip at the measured PCC');

nexttile(dipLayout);
plot(vdc.Time, reshape(vdc.Data, [], 1), 'LineWidth', 1.0);
xline(3.5, ':r', 'Voltage dip');
yline(1200, '--k');
xlim([3.3 4.0]);
grid on;
xlabel('Time (s)');
ylabel('V_{dc} (V)');
title('DC-link ride-through');

exportgraphics(dipFigure, fullfile(experimentDirectory, ...
    'stage15_grid_dip_zoom.png'), 'Resolution', 180);
close(dipFigure);

try
    sdiRunId = Simulink.sdi.createRun( ...
        'stage15_complete_actual_grid_dip_4s', 'vars', ...
        results.signals.V_dc, ...
        results.signals.P_pv, ...
        results.signals.P_bat, ...
        results.signals.P_DC, ...
        results.signals.Pf_threephase, ...
        results.signals.P_grid_threephase, ...
        results.signals.Q_ref_cmd, ...
        results.signals.Qf_threephase);
    sdiRun = Simulink.sdi.getRun(sdiRunId);
    fprintf('SDI_RUN_ID=%d SIGNAL_COUNT=%d\n', ...
        sdiRunId, sdiRun.SignalCount);
catch sdiException
    warning('Stage15:SDIRender', 'SDI creation failed: %s', ...
        sdiException.message);
end

fprintf('RENDER_OK\n');
