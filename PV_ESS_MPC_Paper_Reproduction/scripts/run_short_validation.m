function result = run_short_validation()
%RUN_SHORT_VALIDATION Compile and run a 10 ms switching-circuit smoke test.
thisDir = fileparts(mfilename('fullpath'));
projectRoot = fileparts(thisDir);
addpath(thisDir);
addpath(fullfile(projectRoot,'models'));
model = 'PV_ESS_MPC_Switching_Template';
modelFile = fullfile(projectRoot,'models',[model '.slx']);
if ~exist(modelFile,'file')
    build_pvess_template();
end

cfg = pvess_init(false);
in = Simulink.SimulationInput(model);
in = in.setModelParameter('StopTime','0.01','SimulationMode','normal');
% Keep all paper events outside the short validation window.
far = 1e6;
for name = {'T_ac_step','T_sync','T_grid_connect'}
    in = in.setVariable(name{1},far,'Workspace',model);
end
in = in.setVariable('G_profile',timeseries([1;1],[0;0.01]),'Workspace',model);
in = in.setVariable('Rdc1_profile',timeseries([1e6;1e6],[0;0.01]),'Workspace',model);
in = in.setVariable('Rdc2_profile',timeseries([1e6;1e6],[0;0.01]),'Workspace',model);
in = in.setVariable('GridVpu_profile',timeseries([1;1],[0;0.01]),'Workspace',model);
% Isolate the source converters in the smoke test. This validates the
% initialized 1200 V dc link, interlinking bridge, LC filter and load without
% mixing source-controller commissioning into the structural compile test.
in = in.setVariable('PV_blocking',1,'Workspace',model);
in = in.setVariable('Battery_blocking',1,'Workspace',model);
in = in.setVariable('Inverter_enable',0,'Workspace',model);
in = in.setVariable('PV_connect',0,'Workspace',model);
in = in.setVariable('Battery_connect',0,'Workspace',model);
in = in.setVariable('Inverter_connect',0,'Workspace',model);
in = in.setBlockParameter([model '/PV_Blocking'],'Value','1');
in = in.setBlockParameter([model '/Battery_Blocking'],'Value','1');
in = in.setBlockParameter([model '/Inverter_Enable'],'Value','0');
in = in.setBlockParameter([model '/PV_Connect'],'Value','0');
in = in.setBlockParameter([model '/Battery_Connect'],'Value','0');
in = in.setBlockParameter([model '/Inverter_Connect'],'Value','0');
in = in.setBlockParameter([model '/Lpv'],'InitialCurrent','0');
out = sim(in);

vdc = out.Vdc_log;
vpv = out.Vpv_log;
ibat = out.Ibat_log;
finiteSignals = all(isfinite(vdc.Data(:))) && all(isfinite(vpv.Data(:))) && ...
    all(isfinite(ibat.Data(:)));
result = struct('model',modelFile,'stopTime',0.01, ...
    'finiteSignals',finiteSignals,'VdcMin',min(vdc.Data(:)), ...
    'VdcMax',max(vdc.Data(:)),'VpvMin',min(vpv.Data(:)), ...
    'VpvMax',max(vpv.Data(:)),'IbatMin',min(ibat.Data(:)), ...
    'IbatMax',max(ibat.Data(:)));

resultDir = fullfile(projectRoot,'results');
if ~exist(resultDir,'dir'), mkdir(resultDir); end
save(fullfile(resultDir,'short_validation.mat'),'result','out');
fig = figure('Visible','off','Color','w','Position',[100 100 900 650]);
tiledlayout(3,1,'TileSpacing','compact');
nexttile; plot(vdc.Time,vdc.Data,'LineWidth',1); grid on; ylabel('V_{dc} (V)');
nexttile; plot(vpv.Time,vpv.Data,'LineWidth',1); grid on; ylabel('V_{pv} (V)');
nexttile; plot(ibat.Time,ibat.Data,'LineWidth',1); grid on; ylabel('I_{bat} (A)'); xlabel('Time (s)');
exportgraphics(fig,fullfile(resultDir,'short_validation.png'),'Resolution',180);
close(fig);
fprintf('Short validation: finite=%d, Vdc=[%.2f, %.2f] V\n', ...
    finiteSignals,result.VdcMin,result.VdcMax);
end
