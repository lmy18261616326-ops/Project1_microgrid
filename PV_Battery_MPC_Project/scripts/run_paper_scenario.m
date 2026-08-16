function result = run_paper_scenario(scenarioName, controllerName)
%RUN_PAPER_SCENARIO Run one paper scenario through SimulationInput.

arguments
    scenarioName (1, 1) string = "comparison"
    controllerName (1, 1) string ...
        {mustBeMember(controllerName, ["MPC", "PI"])} = "MPC"
end

assignin('base', 'ScenarioName', scenarioName);
assignin('base', 'ControllerSelect', double(controllerName == "MPC"));
pvbatt_initialize();

in = Simulink.SimulationInput('PV_Battery_MPC');
in = in.setModelParameter( ...
    'StopTime', 'ScenarioStopTime', ...
    'ReturnWorkspaceOutputs', 'on');
out = sim(in);

x = out.sim_x.signals.values;
t = out.sim_x.time;
aux = out.sim_aux.signals.values;
u = out.sim_u.signals.values;
mode = out.sim_mode.signals.values;

vdc = x(:, 6);
result.scenario = scenarioName;
result.controller = controllerName;
result.output = out;
result.time = t;
result.state = x;
result.aux = aux;
result.duty = u;
result.mode = mode;
result.meanVRI = mean(abs(vdc - 600) / 600) * 100;
result.maximumVRI = max(abs(vdc - 600) / 600) * 100;
end
