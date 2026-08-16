%% PV-Battery hybrid AC/DC microgrid paper reproduction
% 1) Build or rebuild the switching template.
projectRoot = fileparts(mfilename('fullpath'));
addpath(fullfile(projectRoot,'scripts'));
modelFile = build_pvess_template();

%% 2) Open the power-circuit template.
open_system(modelFile);

%% 3) Run the 10 ms smoke test first.
validation = run_short_validation(); %#ok<NASGU>

%% 4) After staged commissioning, run the paper event sequence.
% out = run_paper_scenario(0.2);
% out = run_paper_scenario(1.0);
% out = run_paper_scenario(4.0);
