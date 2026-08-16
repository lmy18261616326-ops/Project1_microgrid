function [scenarioW, initialSoc, stopTime] = ...
        pvbatt_build_scenario(name, p)
%PVBATT_BUILD_SCENARIO Published and reconstruction disturbance profiles.
% Columns of scenarioW are [time, temperature, irradiance, load resistance].

vdc = p.dcBus.referenceVoltage;

switch lower(string(name))
    case "comparison"
        % Paper Section 5.5: T 35->25 C and G 500->1000 W/m2 at 1 s;
        % load 7.2->14.4 kW at 2 s.
        time = [0; 1; 1; 2; 2; 3];
        temperature = [35; 35; 25; 25; 25; 25];
        irradiance = [500; 500; 1000; 1000; 1000; 1000];
        loadPower = [7200; 7200; 7200; 7200; 14400; 14400];
        initialSoc = 0.80;
        stopTime = 3;

    case "mode1"
        time = [0; 4; 4; 16];
        temperature = [30; 30; 25; 25];
        irradiance = [800; 800; 1000; 1000];
        loadPower = [8500; 8500; 6500; 6500];
        initialSoc = 0.80;
        stopTime = 16;

    case "mode2"
        time = [0; 4; 4; 16];
        temperature = [25; 25; 30; 30];
        irradiance = [1000; 1000; 700; 700];
        loadPower = [7000; 7000; 10000; 10000];
        initialSoc = 0.80;
        stopTime = 16;

    case "mode3"
        time = [0; 12.2; 12.2; 16];
        temperature = [25; 25; 25; 25];
        irradiance = [1000; 1000; 1000; 1000];
        loadPower = [6000; 6000; 12000; 12000];
        initialSoc = 0.8994;
        stopTime = 16;

    case "mode3_quick"
        % Accelerated verification of the MPPT-to-curtailment transition.
        time = [0; 1];
        temperature = [25; 25];
        irradiance = [1000; 1000];
        loadPower = [6000; 6000];
        initialSoc = 0.89995;
        stopTime = 1;

    case "mode4"
        time = [0; 4.2; 4.2; 16];
        temperature = [25; 25; 25; 25];
        irradiance = [1000; 1000; 0; 0];
        loadPower = [6000; 6000; 4500; 4500];
        initialSoc = 0.895;
        stopTime = 16;

    case "realworld"
        stopTime = 300;
        time = (0:1:stopTime)';
        daylight = max(sin(pi * (time - 25) / 120), 0);
        irradiance = 950 * daylight .* (0.94 ...
            + 0.06 * sin(0.18 * time));
        temperature = 18 + 14 * max( ...
            sin(pi * (time - 10) / 190), 0);
        loadPower = 4200 + 1300 * sin(2 * pi * time / 95).^2 ...
            + 2400 * exp(-((time - 215) / 40).^2);
        initialSoc = 0.80;

    otherwise
        error("PVBATT:UnknownScenario", ...
            "Unknown scenario '%s'.", name);
end

loadResistance = vdc^2 ./ loadPower;
scenarioW = [time, temperature, irradiance, loadResistance];
end
