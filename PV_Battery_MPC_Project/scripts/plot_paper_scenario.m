function figureHandle = plot_paper_scenario(result)
%PLOT_PAPER_SCENARIO Plot the paper's main power-management quantities.

tiledlayout(3, 1, 'TileSpacing', 'compact');
figureHandle = gcf;
figureHandle.Color = 'w';

nexttile;
plot(result.time, result.aux(:, 4:6) / 1000, 'LineWidth', 1.1);
grid on;
ylabel('Power (kW)');
legend('P_{pv}', 'P_{bat}', 'P_{load}', 'Location', 'best');
title(result.controller + " - " + result.scenario);

nexttile;
plot(result.time, result.state(:, 6), 'LineWidth', 1.1);
hold on;
yline(600, '--');
grid on;
ylabel('V_{dc} (V)');

nexttile;
plot(result.time, 100 * result.state(:, 5), 'LineWidth', 1.1);
grid on;
ylabel('SoC (%)');
xlabel('Time (s)');
end
