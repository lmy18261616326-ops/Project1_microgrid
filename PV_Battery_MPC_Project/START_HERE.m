function START_HERE()
%START_HERE Open, initialize, and display the reference project model.

projectRoot = fileparts(mfilename('fullpath'));

try
    proj = currentProject;
    if ~strcmpi(proj.RootFolder, projectRoot)
        close(proj);
        proj = openProject(projectRoot);
    end
catch
    proj = openProject(projectRoot);
end

assignin('base', 'ScenarioName', "comparison");
assignin('base', 'ControllerSelect', 1);
pvbatt_initialize();
open_system(fullfile(proj.RootFolder, 'models', 'PV_Battery_MPC.slx'));

fprintf('PV-battery MPC project initialized.\n');
fprintf('ScenarioName = %s, ControllerSelect = %g\n', ...
    string(evalin('base', 'ScenarioName')), ...
    evalin('base', 'ControllerSelect'));
fprintf('Run: result = run_paper_scenario("comparison","MPC");\n');
end
