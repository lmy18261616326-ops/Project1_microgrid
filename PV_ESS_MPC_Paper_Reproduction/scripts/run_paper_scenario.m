function out = run_paper_scenario(stopTime)
%RUN_PAPER_SCENARIO Run the Table 2 event sequence with SimulationInput.
% The 4 s switch-level run can be computationally expensive. Start with
% 0.2 s, then 1.0 s, and finally 4.0 s after commissioning each subsystem.
if nargin < 1, stopTime = 4.0; end
thisDir = fileparts(mfilename('fullpath'));
projectRoot = fileparts(thisDir);
addpath(thisDir);
addpath(fullfile(projectRoot,'models'));
model = 'PV_ESS_MPC_Switching_Template';
modelFile = fullfile(projectRoot,'models',[model '.slx']);
if ~exist(modelFile,'file'), build_pvess_template(); end
cfg = pvess_init(false);
in = Simulink.SimulationInput(model);
in = in.setModelParameter('StopTime',num2str(stopTime),'SimulationMode','normal');
names = fieldnames(cfg);
for k = 1:numel(names)
    in = in.setVariable(names{k},cfg.(names{k}),'Workspace',model);
end
out = sim(in);
resultDir = fullfile(projectRoot,'results');
if ~exist(resultDir,'dir'), mkdir(resultDir); end
save(fullfile(resultDir,sprintf('paper_scenario_%gs.mat',stopTime)),'out','cfg');
end
