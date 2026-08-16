function y = pvbatt_mpc_output(x, ~)
%PVBATT_MPC_OUTPUT Outputs used by the paper's mode-dependent objective.
% y = [DC bus voltage; PV voltage; battery inductor current].

y = [x(6); x(1); x(4)];
end
