function tests = test_Plant_Derivatives
%TEST_PLANT_DERIVATIVES Component tests for the averaged plant equations.

tests = functiontests(localfunctions);
end

function testChargingOperatingPoint(testCase)
p = phase1_parameters();
[x0, u0, w0, info] = phase1_operating_point( ...
    p, 7.2e3, 0.8, 25, 1000);
[dx, aux] = Plant_Derivatives(x0, u0, w0, p);

verifyLessThan(testCase, max(abs(info.electricalDerivative)), 1e-8);
verifyGreaterThan(testCase, dx(5), 0);
verifyLessThan(testCase, aux(8), 0);
verifyEqual(testCase, aux(9), 7.2e3, "AbsTol", 1e-8);
verifyLessThan(testCase, abs(aux(10)), 1e-8);
end

function testDischargingOperatingPoint(testCase)
p = phase1_parameters();
[x0, u0, w0] = phase1_operating_point( ...
    p, 14.4e3, 0.8, 25, 1000);
[dx, aux] = Plant_Derivatives(x0, u0, w0, p);

verifyLessThan(testCase, max(abs(dx([1:4, 6]))), 1e-8);
verifyLessThan(testCase, dx(5), 0);
verifyGreaterThan(testCase, aux(8), 0);
verifyEqual(testCase, aux(9), 14.4e3, "AbsTol", 1e-8);
end

function testPublishedMppAtStandardConditions(testCase)
p = phase1_parameters();
arrayMppVoltage = p.pv.seriesModules ...
    * p.pv.moduleMaximumPowerVoltage;
arrayCurrent = phase1_pv_current(arrayMppVoltage, 25, 1000, p);
arrayPower = arrayMppVoltage * arrayCurrent;
expectedPower = p.pv.seriesModules * p.pv.parallelStrings ...
    * p.pv.moduleMaximumPower;

verifyEqual(testCase, arrayPower, expectedPower, "RelTol", 1e-12);
end

function testZeroIrradianceProducesZeroPvCurrent(testCase)
p = phase1_parameters();
arrayCurrent = phase1_pv_current(250, 25, 0, p);

verifyEqual(testCase, arrayCurrent, 0);
end

function testDutyCommandsAreLimited(testCase)
p = phase1_parameters();
[x0, ~, w0] = phase1_operating_point( ...
    p, 7.2e3, 0.8, 25, 1000);
[~, aux] = Plant_Derivatives(x0, [-0.2; 1.3], w0, p);

verifyEqual(testCase, aux(4), 0);
verifyEqual(testCase, aux(5), 1);
end

function testPasteReadyBlockMatchesParameterizedFunction(testCase)
p = phase1_parameters();
[x0, u0, w0] = phase1_operating_point( ...
    p, 7.2e3, 0.8, 25, 1000);
[expectedDx, expectedAux] = Plant_Derivatives(x0, u0, w0, p);
[actualDx, actualAux] = Plant_Derivatives_Block(x0, u0, w0);

verifyEqual(testCase, actualDx, expectedDx, "AbsTol", 1e-12);
verifyEqual(testCase, actualAux, expectedAux, "AbsTol", 1e-12);
end
