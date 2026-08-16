experimentDirectory = fileparts(mfilename('fullpath'));
load(fullfile(experimentDirectory, 'stage15_complete_results.mat'), 'results');

vdc = results.signals.V_dc;
vdcTime = vdc.Time(:);
vdcData = reshape(vdc.Data, [], 1);
[vdcMinimum, minimumIndex] = min(vdcData);
[vdcMaximum, maximumIndex] = max(vdcData);

eventNames = {'ac_step', 'dc_0p6MW_on', 'grid_connect', ...
    'dc_1MW_on', 'dc_1MW_off', 'grid_voltage_dip'};
eventTimes = [0.5 1.0 1.6 2.5 3.0 3.5];
nextEventTimes = [1.0 1.6 2.5 3.0 3.5 4.0];
recoveryBand = [0.98 1.02]*1200;
sampleTime = median(diff(vdcTime));
sustainedSamples = max(2, round(0.02/sampleTime));
recoverySeconds = nan(size(eventTimes));

insideBand = vdcData >= recoveryBand(1) & vdcData <= recoveryBand(2);
for eventIndex = 1:numel(eventTimes)
    candidateIndices = find(vdcTime >= eventTimes(eventIndex) & ...
        vdcTime < nextEventTimes(eventIndex));
    for candidateIndex = candidateIndices(:)'
        lastIndex = candidateIndex + sustainedSamples - 1;
        if lastIndex <= numel(vdcData) && all(insideBand(candidateIndex:lastIndex))
            recoverySeconds(eventIndex) = ...
                vdcTime(candidateIndex) - eventTimes(eventIndex);
            break;
        end
    end
end

batteryPower = reshape(results.signals.P_bat.Data, [], 1);
pvPower = reshape(results.signals.P_pv.Data, [], 1);
soc = reshape(results.signals.SOC_internal.Data, [], 1);
jMpvc = reshape(results.signals.Jmin_mpvc.Data, [], 1);
jMppc = reshape(results.signals.Jmin_MPPC.Data, [], 1);
offMppt = double(reshape(results.signals.offMPPT_final.Data, [], 1));
offMpptTime = results.signals.offMPPT_final.Time(:);
transitionIndices = find(abs(diff(offMppt)) > 0);
transitionTimes = offMpptTime(transitionIndices + 1);

detailedMetrics = struct();
detailedMetrics.vdcMinimum_V = vdcMinimum;
detailedMetrics.vdcMinimumTime_s = vdcTime(minimumIndex);
detailedMetrics.vdcMaximum_V = vdcMaximum;
detailedMetrics.vdcMaximumTime_s = vdcTime(maximumIndex);
detailedMetrics.recoveryDefinition = ...
    'First 20 ms continuously inside 0.98-1.02 pu after each event';
detailedMetrics.eventNames = eventNames;
detailedMetrics.recoverySeconds = recoverySeconds;
detailedMetrics.batteryPowerMinimum_W = min(batteryPower);
detailedMetrics.batteryPowerMaximum_W = max(batteryPower);
detailedMetrics.pvPowerMaximum_W = max(pvPower);
detailedMetrics.socMinimum_percent = min(soc);
detailedMetrics.socMaximum_percent = max(soc);
detailedMetrics.jMpvcMaximum = max(jMpvc);
detailedMetrics.jMppcMaximum = max(jMppc);
detailedMetrics.offMpptTransitionTimes_s = transitionTimes;
detailedMetrics.allLoggedSignalsFinite = all([ ...
    isfinite(vdcData); isfinite(batteryPower); isfinite(pvPower); ...
    isfinite(soc); isfinite(jMpvc); isfinite(jMppc)]);

outputHandle = fopen(fullfile(experimentDirectory, ...
    'stage15_detailed_metrics.json'), 'w');
assert(outputHandle ~= -1, 'Unable to create detailed metrics JSON.');
fprintf(outputHandle, '%s', jsonencode(detailedMetrics, 'PrettyPrint', true));
fclose(outputHandle);

fprintf('VDC_MIN=%.3f @ %.6f s\n', ...
    detailedMetrics.vdcMinimum_V, detailedMetrics.vdcMinimumTime_s);
fprintf('VDC_MAX=%.3f @ %.6f s\n', ...
    detailedMetrics.vdcMaximum_V, detailedMetrics.vdcMaximumTime_s);
for eventIndex = 1:numel(eventNames)
    fprintf('RECOVERY %s=%.6f s\n', eventNames{eventIndex}, ...
        recoverySeconds(eventIndex));
end
fprintf('PBAT=[%.1f %.1f] W PPV_MAX=%.1f W SOC=[%.5f %.5f]%%\n', ...
    min(batteryPower), max(batteryPower), max(pvPower), min(soc), max(soc));
fprintf('JMAX MPVC=%.6f MPPC=%.6f ALL_FINITE=%d\n', ...
    max(jMpvc), max(jMppc), detailedMetrics.allLoggedSignalsFinite);
