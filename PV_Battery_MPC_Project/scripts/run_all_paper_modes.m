function results = run_all_paper_modes()
%RUN_ALL_PAPER_MODES Run the four operating modes described in the paper.
% The full set is computationally intensive because the MPC is solved every
% 10 ms while the plant is integrated at the published 10 us step.

names = ["mode1", "mode2", "mode3", "mode4"];
results = struct();

for k = 1:numel(names)
    results.(names(k)) = run_paper_scenario(names(k), "MPC");
end
end
