function nlobj = pvbatt_create_nlmpc(x0, u0, w0, p)
%PVBATT_CREATE_NLMPC Configure the standard Nonlinear MPC Controller block.

nlobj = nlmpc(6, 3, 'MV', [1, 2], 'MD', [3, 4, 5]);
nlobj.Ts = p.controller.interval;
nlobj.PredictionHorizon = p.controller.predictionHorizon;
nlobj.ControlHorizon = 1;
nlobj.Model.IsContinuousTime = false;
nlobj.Model.StateFcn = "pvbatt_state_transition";
nlobj.Model.OutputFcn = "pvbatt_mpc_output";

for k = 1:2
    nlobj.MV(k).Min = p.controller.minimumDuty;
    nlobj.MV(k).Max = p.controller.maximumDuty;
    nlobj.MV(k).RateMin = -0.05;
    nlobj.MV(k).RateMax = 0.05;
    nlobj.MV(k).ScaleFactor = 1;
end

% The paper's weights act directly on V and A rather than normalized
% engineering ranges, so unit scale factors are used here.
nlobj.OV(1).ScaleFactor = 1;
nlobj.OV(2).ScaleFactor = 1;
nlobj.OV(3).ScaleFactor = 1;

nlobj.Weights.OutputVariables = [ ...
    p.controller.weights.dcBusVoltage, ...
    p.controller.weights.pvVoltage, ...
    0];
nlobj.Weights.ManipulatedVariablesRate = ...
    p.controller.weights.dutyRate;

% Numerical guard band permits the controller to arrest residual current at
% the exact mode-transition instant; PMS thresholds remain exactly 20/90%.
nlobj.States(5).Min = p.battery.minimumStateOfCharge - 5e-4;
nlobj.States(5).Max = p.battery.maximumStateOfCharge + 5e-4;

nlobj.Optimization.SolverOptions.Algorithm = "sqp";
nlobj.Optimization.SolverOptions.Display = "none";
nlobj.Optimization.SolverOptions.MaxIterations = 20;
nlobj.Optimization.UseSuboptimalSolution = true;

validateFcns(nlobj, x0, u0(:)', w0(:)');
end
