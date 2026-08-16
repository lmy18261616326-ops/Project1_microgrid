function report = run_validation()
%RUN_VALIDATION Execute the principal paper-reproduction checks.

projectRoot = fileparts(fileparts(mfilename('fullpath')));
resultsFolder = fullfile(projectRoot, 'results');
if ~isfolder(resultsFolder)
    mkdir(resultsFolder);
end

mpcResult = run_paper_scenario("comparison", "MPC");
piResult = run_paper_scenario("comparison", "PI");
mode3Result = run_paper_scenario("mode3_quick", "MPC");

report.mpcMeanVRI = mpcResult.meanVRI;
report.mpcMaximumVRI = mpcResult.maximumVRI;
report.piMeanVRI = piResult.meanVRI;
report.piMaximumVRI = piResult.maximumVRI;

mode3Index = find(mode3Result.mode(:, 2) == 3, 1, 'first');
report.mode3EntryTime = mode3Result.time(mode3Index);
report.mode3FinalPVPower = mode3Result.aux(end, 4);
report.mode3FinalBatteryPower = mode3Result.aux(end, 5);
report.mode3FinalDCVoltage = mode3Result.state(end, 6);

comparison = localReducedResult(mpcResult);
baseline = localReducedResult(piResult);
curtailment = localReducedResult(mode3Result);
save(fullfile(resultsFolder, 'validation_results.mat'), ...
    'comparison', 'baseline', 'curtailment', 'report');

comparisonFigure = figure('Color', 'w', 'Name', 'MPC vs PI');
tiledlayout(comparisonFigure, 2, 1, 'TileSpacing', 'compact');
nexttile;
plot(mpcResult.time, mpcResult.state(:, 6), 'LineWidth', 1.1);
hold on;
plot(piResult.time, piResult.state(:, 6), 'LineWidth', 1.0);
yline(600, '--k');
grid on;
ylabel('V_{dc} (V)');
legend('MPC', 'PI', 'Reference', 'Location', 'best');
title('Paper Section 5.5 Disturbance Comparison');
nexttile;
plot(mpcResult.time, ...
    abs(mpcResult.state(:, 6) - 600) / 600 * 100, ...
    'LineWidth', 1.1);
hold on;
plot(piResult.time, ...
    abs(piResult.state(:, 6) - 600) / 600 * 100, ...
    'LineWidth', 1.0);
grid on;
ylabel('VRI (%)');
xlabel('Time (s)');
legend('MPC', 'PI', 'Location', 'best');
exportgraphics(comparisonFigure, ...
    fullfile(resultsFolder, 'comparison_mpc_vs_pi.png'), ...
    'Resolution', 180);

modeFigure = figure('Color', 'w', 'Name', 'Mode III Curtailment');
tiledlayout(modeFigure, 3, 1, 'TileSpacing', 'compact');
nexttile;
plot(mode3Result.time, mode3Result.aux(:, 4:6) / 1000, ...
    'LineWidth', 1.0);
grid on;
ylabel('Power (kW)');
legend('P_{pv}', 'P_{bat}', 'P_{load}', 'Location', 'best');
title('Accelerated Mode III Verification');
nexttile;
plot(mode3Result.time, mode3Result.state(:, 6), 'LineWidth', 1.1);
hold on;
yline(600, '--k');
grid on;
ylabel('V_{dc} (V)');
nexttile;
yyaxis left;
plot(mode3Result.time, 100 * mode3Result.state(:, 5), ...
    'LineWidth', 1.1);
ylabel('SoC (%)');
yyaxis right;
stairs(mode3Result.time, mode3Result.mode(:, 2), ...
    'LineWidth', 1.0);
ylabel('Mode ID');
grid on;
xlabel('Time (s)');
exportgraphics(modeFigure, ...
    fullfile(resultsFolder, 'mode3_curtailment.png'), ...
    'Resolution', 180);
end

function reduced = localReducedResult(result)
reduced.scenario = result.scenario;
reduced.controller = result.controller;
reduced.time = result.time;
reduced.state = result.state;
reduced.aux = result.aux;
reduced.duty = result.duty;
reduced.mode = result.mode;
reduced.meanVRI = result.meanVRI;
reduced.maximumVRI = result.maximumVRI;
end
