%% Add a transparent P/Q current limiter before the grid-connected MPPC.
% Uses the measured 690 V-side three-phase voltage and the 2.5 MVA
% transformer/converter rating. Q has priority during voltage support.

expDir = fileparts(mfilename('fullpath'));
mdl = 'manul_model_stage15_complete';
cd(expDir);
run(fullfile(expDir,'pvess_stage15_busfix_init.m'));
bdclose('all');
load_system(fullfile(expDir,[mdl '.slx']));

sub = [mdl '/Grid_PQ_Current_Limiter'];
if getSimulinkBlockHandle(sub) ~= -1
    error('Grid_PQ_Current_Limiter already exists. Refusing to apply twice.');
end

pGoto = [mdl '/Goto5'];
qGoto = [mdl '/Goto2'];
pPos = get_param(pGoto,'Position');
qPos = get_param(qGoto,'Position');
left = min(pPos(1),qPos(1))-260;
top = min(pPos(2),qPos(2))-20;
add_block('built-in/SubSystem',sub,'Position',[left top left+210 top+150]);
set_param(sub,'BackgroundColor','lightBlue', ...
    'AttributesFormatString',['2.5 MVA grid P/Q current limit\n' ...
    'Q priority during voltage support']);

% Interface: Pref raw, Qref raw, measured grid Vabc -> limited P/Q, Smax, flag.
add_block('simulink/Sources/In1',[sub '/Pref_raw'],'Port','1','Position',[35 70 65 84]);
add_block('simulink/Sources/In1',[sub '/Qref_raw'],'Port','2','Position',[35 180 65 194]);
add_block('simulink/Sources/In1',[sub '/Vabc_grid_lv'],'Port','3','Position',[35 300 65 314]);
outNames = {'Pref_limited','Qref_limited','S_current_max','grid_current_limit_active'};
for k=1:numel(outNames)
    add_block('simulink/Sinks/Out1',[sub '/' outNames{k}], ...
        'Port',num2str(k),'Position',[1100 70+90*(k-1) 1130 84+90*(k-1)]);
end

% Instantaneous balanced three-phase identity: sqrt(va^2+vb^2+vc^2)=Vll_rms.
add_block('simulink/Math Operations/Dot Product',[sub '/Vabc_Dot'], ...
    'Position',[120 270 170 325]);
add_block('simulink/Sources/Constant',[sub '/Vll_Floor_Squared'], ...
    'Value','Vgrid_limit_floor^2','Position',[120 350 245 380]);
add_block('simulink/Math Operations/MinMax',[sub '/Vll2_Safe_Max'], ...
    'Function','max','Inputs','2','Position',[285 285 335 340]);
add_block('simulink/Math Operations/Math Function',[sub '/Vll_RMS_Sqrt'], ...
    'Operator','sqrt','Position',[375 292 430 333]);
add_block('simulink/Math Operations/Gain',[sub '/Smax_From_Current'], ...
    'Gain','sqrt(3)*Iac_rated_lv','Position',[470 285 600 340]);

% Q-priority clamp: Qlim=max(-Smax,min(Qraw,Smax)).
add_block('simulink/Math Operations/MinMax',[sub '/Q_Clamp_Upper'], ...
    'Function','min','Inputs','2','Position',[655 130 705 185]);
add_block('simulink/Math Operations/Gain',[sub '/Negative_Smax'], ...
    'Gain','-1','Position',[655 215 705 255]);
add_block('simulink/Math Operations/MinMax',[sub '/Q_Clamp_Lower'], ...
    'Function','max','Inputs','2','Position',[750 145 800 205]);

% P allowance from the remaining apparent-power circle.
add_block('simulink/Math Operations/Product',[sub '/Smax_Squared'], ...
    'Inputs','**','Position',[650 300 700 350]);
add_block('simulink/Math Operations/Product',[sub '/Qcmd_Squared'], ...
    'Inputs','**','Position',[830 160 880 210]);
add_block('simulink/Math Operations/Sum',[sub '/Pallow_Squared'], ...
    'Inputs','|+-','Position',[750 295 785 345]);
add_block('simulink/Sources/Constant',[sub '/Zero'], ...
    'Value','0','Position',[750 365 795 395]);
add_block('simulink/Math Operations/MinMax',[sub '/Pallow2_Nonnegative'], ...
    'Function','max','Inputs','2','Position',[830 290 880 345]);
add_block('simulink/Math Operations/Math Function',[sub '/Pallow_Sqrt'], ...
    'Operator','sqrt','Position',[915 298 970 337]);
add_block('simulink/Math Operations/MinMax',[sub '/P_Clamp_Upper'], ...
    'Function','min','Inputs','2','Position',[655 55 705 110]);
add_block('simulink/Math Operations/Gain',[sub '/Negative_Pallow'], ...
    'Gain','-1','Position',[915 365 970 405]);
add_block('simulink/Math Operations/MinMax',[sub '/P_Clamp_Lower'], ...
    'Function','max','Inputs','2','Position',[750 60 800 120]);

% Limit-active flag compares raw and executable commands.
add_block('simulink/Math Operations/Sum',[sub '/P_Limit_Error'], ...
    'Inputs','|+-','Position',[835 50 870 95]);
add_block('simulink/Math Operations/Abs',[sub '/Abs_P_Limit_Error'], ...
    'Position',[895 52 935 92]);
add_block('simulink/Logic and Bit Operations/Compare To Constant', ...
    [sub '/P_Limit_Compare'],'relop','>','const','PQLimit_eps', ...
    'Position',[955 45 1025 100]);
add_block('simulink/Math Operations/Sum',[sub '/Q_Limit_Error'], ...
    'Inputs','|+-','Position',[835 225 870 270]);
add_block('simulink/Math Operations/Abs',[sub '/Abs_Q_Limit_Error'], ...
    'Position',[895 227 935 267]);
add_block('simulink/Logic and Bit Operations/Compare To Constant', ...
    [sub '/Q_Limit_Compare'],'relop','>','const','PQLimit_eps', ...
    'Position',[955 220 1025 275]);
add_block('simulink/Logic and Bit Operations/Logical Operator', ...
    [sub '/Any_PQ_Limit'],'Operator','OR','Inputs','2', ...
    'Position',[1040 165 1080 235]);

% Wiring
add_line(sub,'Vabc_grid_lv/1','Vabc_Dot/1','autorouting','on');
add_line(sub,'Vabc_grid_lv/1','Vabc_Dot/2','autorouting','on');
add_line(sub,'Vabc_Dot/1','Vll2_Safe_Max/1','autorouting','on');
add_line(sub,'Vll_Floor_Squared/1','Vll2_Safe_Max/2','autorouting','on');
add_line(sub,'Vll2_Safe_Max/1','Vll_RMS_Sqrt/1','autorouting','on');
add_line(sub,'Vll_RMS_Sqrt/1','Smax_From_Current/1','autorouting','on');
add_line(sub,'Smax_From_Current/1','S_current_max/1','autorouting','on');

add_line(sub,'Qref_raw/1','Q_Clamp_Upper/1','autorouting','on');
add_line(sub,'Smax_From_Current/1','Q_Clamp_Upper/2','autorouting','on');
add_line(sub,'Smax_From_Current/1','Negative_Smax/1','autorouting','on');
add_line(sub,'Q_Clamp_Upper/1','Q_Clamp_Lower/1','autorouting','on');
add_line(sub,'Negative_Smax/1','Q_Clamp_Lower/2','autorouting','on');
add_line(sub,'Q_Clamp_Lower/1','Qref_limited/1','autorouting','on');

add_line(sub,'Smax_From_Current/1','Smax_Squared/1','autorouting','on');
add_line(sub,'Smax_From_Current/1','Smax_Squared/2','autorouting','on');
add_line(sub,'Q_Clamp_Lower/1','Qcmd_Squared/1','autorouting','on');
add_line(sub,'Q_Clamp_Lower/1','Qcmd_Squared/2','autorouting','on');
add_line(sub,'Smax_Squared/1','Pallow_Squared/1','autorouting','on');
add_line(sub,'Qcmd_Squared/1','Pallow_Squared/2','autorouting','on');
add_line(sub,'Pallow_Squared/1','Pallow2_Nonnegative/1','autorouting','on');
add_line(sub,'Zero/1','Pallow2_Nonnegative/2','autorouting','on');
add_line(sub,'Pallow2_Nonnegative/1','Pallow_Sqrt/1','autorouting','on');
add_line(sub,'Pref_raw/1','P_Clamp_Upper/1','autorouting','on');
add_line(sub,'Pallow_Sqrt/1','P_Clamp_Upper/2','autorouting','on');
add_line(sub,'Pallow_Sqrt/1','Negative_Pallow/1','autorouting','on');
add_line(sub,'P_Clamp_Upper/1','P_Clamp_Lower/1','autorouting','on');
add_line(sub,'Negative_Pallow/1','P_Clamp_Lower/2','autorouting','on');
add_line(sub,'P_Clamp_Lower/1','Pref_limited/1','autorouting','on');

add_line(sub,'Pref_raw/1','P_Limit_Error/1','autorouting','on');
add_line(sub,'P_Clamp_Lower/1','P_Limit_Error/2','autorouting','on');
add_line(sub,'P_Limit_Error/1','Abs_P_Limit_Error/1','autorouting','on');
add_line(sub,'Abs_P_Limit_Error/1','P_Limit_Compare/1','autorouting','on');
add_line(sub,'Qref_raw/1','Q_Limit_Error/1','autorouting','on');
add_line(sub,'Q_Clamp_Lower/1','Q_Limit_Error/2','autorouting','on');
add_line(sub,'Q_Limit_Error/1','Abs_Q_Limit_Error/1','autorouting','on');
add_line(sub,'Abs_Q_Limit_Error/1','Q_Limit_Compare/1','autorouting','on');
add_line(sub,'P_Limit_Compare/1','Any_PQ_Limit/1','autorouting','on');
add_line(sub,'Q_Limit_Compare/1','Any_PQ_Limit/2','autorouting','on');
add_line(sub,'Any_PQ_Limit/1','grid_current_limit_active/1','autorouting','on');

%% Top-level integration and transparent logs.
add_line(mdl,'Saturation4/1','Grid_PQ_Current_Limiter/1','autorouting','on');
add_line(mdl,'Gain/1','Grid_PQ_Current_Limiter/2','autorouting','on');
add_line(mdl,'grid_VI/1','Grid_PQ_Current_Limiter/3','autorouting','on');

targets = {[mdl '/Goto5'],[mdl '/Goto2'],[mdl '/Log_Pref_Command'],[mdl '/Log_Qref_Command']};
for k=1:numel(targets)
    ph=get_param(targets{k},'PortHandles');
    ln=get_param(ph.Inport(1),'Line');
    if ln~=-1, delete_line(ln); end
end
add_line(mdl,'Grid_PQ_Current_Limiter/1','Goto5/1','autorouting','on');
add_line(mdl,'Grid_PQ_Current_Limiter/1','Log_Pref_Command/1','autorouting','on');
add_line(mdl,'Grid_PQ_Current_Limiter/2','Goto2/1','autorouting','on');
add_line(mdl,'Grid_PQ_Current_Limiter/2','Log_Qref_Command/1','autorouting','on');

logSpecs = { ...
    'Log_Pref_Raw','P_ref_raw','Saturation4/1'; ...
    'Log_Qref_Raw','Q_ref_raw','Gain/1'; ...
    'Log_Sgrid_Current_Max','S_grid_current_max','Grid_PQ_Current_Limiter/3'; ...
    'Log_Grid_Current_Limit','grid_current_limit_active','Grid_PQ_Current_Limiter/4'};
for k=1:size(logSpecs,1)
    add_block('simulink/Sinks/To Workspace',[mdl '/' logSpecs{k,1}], ...
        'VariableName',logSpecs{k,2},'SaveFormat','Timeseries', ...
        'Position',[left+260 top+25+45*(k-1) left+420 top+50+45*(k-1)]);
    add_line(mdl,logSpecs{k,3},[logSpecs{k,1} '/1'],'autorouting','on');
end

ann=Simulink.Annotation(mdl,['GRID CURRENT LIMIT\n' ...
    'Smax=sqrt(3)*Vll(LV)*2091.85 A\nQ priority; raw and limited commands logged']);
ann.Position=[left-10 top-100 left+410 top-35];
ann.ForegroundColor='green';

save_system(mdl);
fprintf('Added Grid_PQ_Current_Limiter to %s\n',mdl);
bdclose(mdl);
