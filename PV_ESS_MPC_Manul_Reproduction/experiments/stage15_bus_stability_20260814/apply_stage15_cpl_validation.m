%% Convert the Stage 15 experiment copy to AC/DC constant-power-load validation.
% AC: use the built-in SPS Three-Phase Parallel RLC Load constant-PQ mode.
% DC: keep the built-in SPS Variable Resistors and command R=Vdc^2/P using
% standard Simulink blocks. No MATLAB Function block or waveform clamp is used.

expDir=fileparts(mfilename('fullpath'));
mdl='manul_model_stage15_complete';
cd(expDir);
run(fullfile(expDir,'pvess_stage15_busfix_init.m'));
bdclose('all');
load_system(fullfile(expDir,[mdl '.slx']));

sub=[mdl '/DC_CPL_Resistance_Command'];
if getSimulinkBlockHandle(sub) ~= -1
    error('DC_CPL_Resistance_Command already exists. Refusing to apply twice.');
end

%% Find and repurpose the two existing DC-resistance profile sources.
fw=find_system(mdl,'SearchDepth',1,'BlockType','FromWorkspace');
src1=''; src2='';
for k=1:numel(fw)
    variable=get_param(fw{k},'VariableName');
    if strcmp(variable,'Rdc1_profile'), src1=fw{k}; end
    if strcmp(variable,'Rdc2_profile'), src2=fw{k}; end
end
assert(~isempty(src1) && ~isempty(src2),'Existing DC load profile sources not found.');
set_param(src1,'Name','Pdc1_CPL_Profile','VariableName','Pdc1_cpl_profile', ...
    'Position',[600 105 715 135]);
set_param(src2,'Name','Pdc2_CPL_Profile','VariableName','Pdc2_cpl_profile', ...
    'Position',[600 160 715 190]);
src1=[mdl '/Pdc1_CPL_Profile'];
src2=[mdl '/Pdc2_CPL_Profile'];

for b={src1,src2}
    ph=get_param(b{1},'PortHandles');
    line=get_param(ph.Outport(1),'Line');
    if line~=-1, delete_line(line); end
end

%% Standard-block sampled CPL resistance calculator.
add_block('built-in/SubSystem',sub,'Position',[770 85 1040 220]);
set_param(sub,'BackgroundColor','magenta', ...
    'AttributesFormatString',['DC constant-power load\n' ...
    'R=Vdc(z-1)^2/P, Ts=50 us\n0.70 pu voltage floor; physical R limits']);

add_block('simulink/Sources/In1',[sub '/P1_command_W'], ...
    'Port','1','Position',[30 50 60 64]);
add_block('simulink/Sources/In1',[sub '/P2_command_W'], ...
    'Port','2','Position',[30 170 60 184]);
add_block('simulink/Sources/In1',[sub '/Vdc_measured_V'], ...
    'Port','3','Position',[30 290 60 304]);

add_block('simulink/Discrete/Unit Delay',[sub '/Vdc_z1'], ...
    'SampleTime','Ts_cpl','InitialCondition','Vdc_ref', ...
    'Position',[105 275 160 320]);
add_block('simulink/Math Operations/Abs',[sub '/Abs_Vdc'], ...
    'Position',[200 278 240 317]);
add_block('simulink/Sources/Constant',[sub '/Vdc_Protection_Floor'], ...
    'Value','Vdc_cpl_floor','Position',[190 350 280 380]);
add_block('simulink/Math Operations/MinMax',[sub '/Vdc_Safe_Max'], ...
    'Function','max','Inputs','2','Position',[320 285 370 340]);
add_block('simulink/Math Operations/Product',[sub '/Vdc_Squared'], ...
    'Inputs','**','Position',[415 290 465 335]);

for branch=1:2
    y=35+120*(branch-1);
    prefix=sprintf('Load%d_',branch);
    pIn=sprintf('P%d_command_W',branch);
    add_block('simulink/Sources/Constant',[sub '/' prefix 'P_Floor'], ...
        'Value','Pdc_cpl_eps','Position',[105 y+35 180 y+65]);
    add_block('simulink/Math Operations/MinMax',[sub '/' prefix 'P_Safe_Max'], ...
        'Function','max','Inputs','2','Position',[215 y 265 y+55]);
    add_block('simulink/Math Operations/Divide',[sub '/' prefix 'R_V2_over_P'], ...
        'Position',[515 y 565 y+55]);
    add_block('simulink/Discontinuities/Saturation',[sub '/' prefix 'R_Physical_Limits'], ...
        'LowerLimit','Rdc_cpl_min','UpperLimit','Rdc_cpl_off', ...
        'Position',[610 y 680 y+55]);
    add_block('simulink/Sources/Constant',[sub '/' prefix 'R_Off'], ...
        'Value','Rdc_cpl_off','Position',[615 y+80 680 y+110]);
    add_block('simulink/Signal Routing/Switch',[sub '/' prefix 'On_Off'], ...
        'Criteria','u2 > Threshold','Threshold','Pdc_cpl_eps', ...
        'Position',[740 y 790 y+85]);
    add_block('simulink/Sinks/Out1',[sub '/' prefix 'R_command_ohm'], ...
        'Port',num2str(branch),'Position',[850 y+27 880 y+41]);

    add_line(sub,[pIn '/1'],[prefix 'P_Safe_Max/1'],'autorouting','on');
    add_line(sub,[prefix 'P_Floor/1'],[prefix 'P_Safe_Max/2'],'autorouting','on');
    add_line(sub,'Vdc_Squared/1',[prefix 'R_V2_over_P/1'],'autorouting','on');
    add_line(sub,[prefix 'P_Safe_Max/1'],[prefix 'R_V2_over_P/2'],'autorouting','on');
    add_line(sub,[prefix 'R_V2_over_P/1'],[prefix 'R_Physical_Limits/1'],'autorouting','on');
    add_line(sub,[prefix 'R_Physical_Limits/1'],[prefix 'On_Off/1'],'autorouting','on');
    add_line(sub,[pIn '/1'],[prefix 'On_Off/2'],'autorouting','on');
    add_line(sub,[prefix 'R_Off/1'],[prefix 'On_Off/3'],'autorouting','on');
    add_line(sub,[prefix 'On_Off/1'],[prefix 'R_command_ohm/1'],'autorouting','on');
end
add_line(sub,'Vdc_measured_V/1','Vdc_z1/1','autorouting','on');
add_line(sub,'Vdc_z1/1','Abs_Vdc/1','autorouting','on');
add_line(sub,'Abs_Vdc/1','Vdc_Safe_Max/1','autorouting','on');
add_line(sub,'Vdc_Protection_Floor/1','Vdc_Safe_Max/2','autorouting','on');
add_line(sub,'Vdc_Safe_Max/1','Vdc_Squared/1','autorouting','on');
add_line(sub,'Vdc_Safe_Max/1','Vdc_Squared/2','autorouting','on');

%% Top-level wiring: command profiles + measured DC voltage -> existing loads.
add_block('simulink/Signal Routing/From',[mdl '/Vdc_CPL_From'], ...
    'GotoTag','V_dc','Position',[610 215 705 245]);
add_line(mdl,'Pdc1_CPL_Profile/1','DC_CPL_Resistance_Command/1','autorouting','on');
add_line(mdl,'Pdc2_CPL_Profile/1','DC_CPL_Resistance_Command/2','autorouting','on');
add_line(mdl,'Vdc_CPL_From/1','DC_CPL_Resistance_Command/3','autorouting','on');

for k=1:2
    loadBlock=[mdl sprintf('/DC_load%d',k)];
    ph=get_param(loadBlock,'PortHandles');
    oldLine=get_param(ph.Inport(1),'Line');
    if oldLine~=-1, delete_line(oldLine); end
    add_line(mdl,sprintf('DC_CPL_Resistance_Command/%d',k), ...
        sprintf('DC_load%d/1',k),'autorouting','on');
end

logSpecs={ ...
    'Log_Pdc1_CPL_Command','Pdc1_cpl_command','Pdc1_CPL_Profile/1',[1080 90 1235 120]; ...
    'Log_Pdc2_CPL_Command','Pdc2_cpl_command','Pdc2_CPL_Profile/1',[1080 130 1235 160]; ...
    'Log_Rdc1_CPL_Command','Rdc1_cpl_command','DC_CPL_Resistance_Command/1',[1080 170 1235 200]; ...
    'Log_Rdc2_CPL_Command','Rdc2_cpl_command','DC_CPL_Resistance_Command/2',[1080 210 1235 240]};
for k=1:size(logSpecs,1)
    add_block('simulink/Sinks/To Workspace',[mdl '/' logSpecs{k,1}], ...
        'VariableName',logSpecs{k,2},'SaveFormat','Timeseries', ...
        'Position',logSpecs{k,4});
    add_line(mdl,logSpecs{k,3},[logSpecs{k,1} '/1'],'autorouting','on');
end

%% AC pressure case: change only the two switched active-power loads.
acLoads=find_system(mdl,'SearchDepth',1,'MaskType','Three-Phase Parallel RLC Load');
changed={};
for k=1:numel(acLoads)
    p=get_param(acLoads{k},'ActivePower');
    if strcmp(p,'P_ac_base') || strcmp(p,'P_ac_step')
        set_param(acLoads{k},'LoadType','constant PQ');
        changed{end+1}=acLoads{k}; %#ok<AGROW>
    end
end
assert(numel(changed)==2,'Expected exactly two active-power AC loads.');

ann=Simulink.Annotation(mdl,['CONSTANT-POWER LOAD VALIDATION\n' ...
    'AC loads: built-in constant PQ\n' ...
    'DC loads: sampled R=Vdc^2/P using existing variable resistors']);
ann.Position=[600 15 1250 65];
ann.ForegroundColor='magenta';

set_param(mdl,'SimulationCommand','update');
save_system(mdl);
fprintf('CPL_APPLY_OK AC_loads=%d DC_subsystem_blocks=%d\n', ...
    numel(changed),numel(find_system(sub,'SearchDepth',1,'Type','Block'))-1);
for k=1:numel(changed)
    fprintf('AC_CPL_BLOCK=%s LoadType=%s P=%s\n',changed{k}, ...
        get_param(changed{k},'LoadType'),get_param(changed{k},'ActivePower'));
end
close_system(mdl,0);
