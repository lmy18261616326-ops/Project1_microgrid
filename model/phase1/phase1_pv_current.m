function pvCurrent = phase1_pv_current( ...
        pvVoltage, temperatureC, irradiance, p)
%PHASE1_PV_CURRENT Empirical array I-V curve for phase-1 reproduction.
%   The curve passes through Isc, Voc, and the published MPP at standard
%   test conditions. It is a replaceable approximation, not a claim that
%   the paper used this PV parameterization.
%
%#codegen

if irradiance <= p.limits.minimumActiveIrradiance
    pvCurrent = 0;
    return;
end

temperatureOffset = temperatureC - p.pv.referenceTemperatureC;
irradianceRatio = irradiance / p.pv.referenceIrradiance;

moduleShortCircuitCurrent = p.pv.moduleShortCircuitCurrent ...
    + p.pv.shortCircuitCurrentTempCoefficient * temperatureOffset;
arrayShortCircuitCurrent = p.pv.parallelStrings ...
    * irradianceRatio * max(moduleShortCircuitCurrent, 0);

logIrradianceRatio = log(max( ...
    irradianceRatio, p.pv.minimumIrradianceRatio));
moduleOpenCircuitVoltage = p.pv.moduleOpenCircuitVoltage ...
    + p.pv.openCircuitVoltageTempCoefficient * temperatureOffset ...
    + p.pv.openCircuitVoltageLogCoefficient * logIrradianceRatio;
moduleOpenCircuitVoltage = max( ...
    moduleOpenCircuitVoltage, p.pv.minimumModuleOpenCircuitVoltage);
arrayOpenCircuitVoltage = p.pv.seriesModules ...
    * moduleOpenCircuitVoltage;

normalizedVoltage = max(pvVoltage, 0) / arrayOpenCircuitVoltage;
if normalizedVoltage >= 1
    pvCurrent = 0;
else
    pvCurrent = arrayShortCircuitCurrent ...
        * (1 - normalizedVoltage^p.pv.curveExponent);
end
end
