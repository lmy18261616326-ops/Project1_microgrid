%% Apply Stage 15 engineering hardening directly to the existing experiment copy.
% This script is idempotent: it refuses to run if the new subsystem exists.
% The original model outside this experiment directory is never opened or saved.

expDir = fileparts(mfilename('fullpath'));
mdl = 'manul_model_stage15_complete';
modelFile = fullfile(expDir, [mdl '.slx']);
cd(expDir);
run(fullfile(expDir, 'pvess_stage15_busfix_init.m'));
bdclose('all');
load_system(modelFile);

newSub = [mdl '/Battery_Reference_Protection'];
if getSimulinkBlockHandle(newSub) ~= -1
    error('Engineering subsystem already exists. Refusing to apply twice.');
end

oldSoc = [mdl '/MATLAB Function'];
assert(getSimulinkBlockHandle(oldSoc) ~= -1, 'Expected SOC MATLAB Function was not found.');
oldPos = get_param(oldSoc, 'Position');

%% 1. Replace the simple SOC MATLAB Function with a standard-block subsystem.
delete_block(oldSoc);
for candidate = {[mdl '/Constant21'], [mdl '/Constant24']}
    if getSimulinkBlockHandle(candidate{1}) ~= -1
        ph = get_param(candidate{1}, 'PortHandles');
        if ph.Outport ~= -1 && get_param(ph.Outport, 'Line') == -1
            delete_block(candidate{1});
        end
    end
end

subPos = [oldPos(1), oldPos(2), oldPos(1)+210, oldPos(2)+150];
add_block('built-in/SubSystem', newSub, 'Position', subPos);
set_param(newSub, 'BackgroundColor', 'lightBlue', ...
    'AttributesFormatString', ['Dynamic 500 kW / current / SOC limits\n' ...
    'Positive Ibat = discharge']);

% Ports
add_block('simulink/Sources/In1', [newSub '/I_cmd_selected'], ...
    'Port','1','Position',[35 78 65 92]);
add_block('simulink/Sources/In1', [newSub '/V_bat'], ...
    'Port','2','Position',[35 178 65 192]);
add_block('simulink/Sources/In1', [newSub '/SOC'], ...
    'Port','3','Position',[35 328 65 342]);

outNames = {'I_ref_protected','I_dis_limit_dyn','I_chg_limit_dyn', ...
    'bat_limit_active','bat_power_limit_active','Vbat_safe'};
for k = 1:numel(outNames)
    add_block('simulink/Sinks/Out1', [newSub '/' outNames{k}], ...
        'Port',num2str(k),'Position',[1105 35+70*(k-1) 1135 49+70*(k-1)]);
end

% Vbat filter and safe denominator
add_block('simulink/Discrete/Discrete Transfer Fcn', [newSub '/Vbat_LPF'], ...
    'Numerator','[1-a_vbat_limit]', 'Denominator','[1 -a_vbat_limit]', ...
    'InitialStates','V_bat_nom', 'SampleTime','Ts_bat_ctrl', ...
    'Position',[105 155 245 215]);
add_block('simulink/Math Operations/Abs', [newSub '/Abs_Vbat'], ...
    'Position',[285 165 325 205]);
add_block('simulink/Sources/Constant', [newSub '/Vbat_Floor'], ...
    'Value','Vbat_div_floor','Position',[285 225 365 255]);
add_block('simulink/Math Operations/MinMax', [newSub '/Vbat_Safe_Max'], ...
    'Function','max','Inputs','2','Position',[405 170 455 225]);

% Discharge limit path
add_block('simulink/Sources/Constant', [newSub '/Pdis_Max'], ...
    'Value','Pbat_dis_max','Position',[500 70 590 100]);
add_block('simulink/Math Operations/Product', [newSub '/Pdis_Over_Vbat'], ...
    'Inputs','*/','Position',[625 60 675 110]);
add_block('simulink/Sources/Constant', [newSub '/Ihw_Dis'], ...
    'Value','Ibat_dis_hw','Position',[625 125 715 155]);
add_block('simulink/Math Operations/MinMax', [newSub '/Min_Dis_Power_Current'], ...
    'Function','min','Inputs','2','Position',[750 70 800 125]);
add_block('simulink/Lookup Tables/1-D Lookup Table', [newSub '/SOC_Dis_Derate'], ...
    'BreakpointsForDimension1','SOC_dis_bp','Table','SOC_dis_k', ...
    'InterpMethod','Linear point-slope','ExtrapMethod','Clip', ...
    'Position',[500 290 650 345]);
add_block('simulink/Math Operations/Product', [newSub '/I_Dis_Limit'], ...
    'Inputs','**','Position',[850 80 900 130]);

% Charge limit path
add_block('simulink/Sources/Constant', [newSub '/Pchg_Max'], ...
    'Value','Pbat_chg_max','Position',[500 190 590 220]);
add_block('simulink/Math Operations/Product', [newSub '/Pchg_Over_Vbat'], ...
    'Inputs','*/','Position',[625 180 675 230]);
add_block('simulink/Sources/Constant', [newSub '/Ihw_Chg'], ...
    'Value','Ibat_chg_hw','Position',[625 245 715 275]);
add_block('simulink/Math Operations/MinMax', [newSub '/Min_Chg_Power_Current'], ...
    'Function','min','Inputs','2','Position',[750 190 800 245]);
add_block('simulink/Lookup Tables/1-D Lookup Table', [newSub '/SOC_Chg_Derate'], ...
    'BreakpointsForDimension1','SOC_chg_bp','Table','SOC_chg_k', ...
    'InterpMethod','Linear point-slope','ExtrapMethod','Clip', ...
    'Position',[500 365 650 420]);
add_block('simulink/Math Operations/Product', [newSub '/I_Chg_Magnitude'], ...
    'Inputs','**','Position',[850 200 900 250]);
add_block('simulink/Math Operations/Gain', [newSub '/Negative_Charge_Limit'], ...
    'Gain','-1','Position',[925 200 975 250]);

% Explicit dynamic clamp: min(command, upper), then max(result, lower).
add_block('simulink/Math Operations/MinMax', [newSub '/Clamp_Upper'], ...
    'Function','min','Inputs','2','Position',[750 455 800 510]);
add_block('simulink/Math Operations/MinMax', [newSub '/Clamp_Lower'], ...
    'Function','max','Inputs','2','Position',[850 455 900 510]);

% Limit active flag
add_block('simulink/Math Operations/Sum', [newSub '/Command_Minus_Limited'], ...
    'Inputs','|+-','Position',[930 455 960 505]);
add_block('simulink/Math Operations/Abs', [newSub '/Abs_Limit_Error'], ...
    'Position',[980 460 1020 500]);
add_block('simulink/Logic and Bit Operations/Compare To Constant', ...
    [newSub '/Limit_Active_Compare'], 'relop','>', 'const','Ilim_eps', ...
    'Position',[1035 455 1090 505]);

% Power/SOC derating active flag. Both charge and discharge reductions are visible.
add_block('simulink/Math Operations/Sum', [newSub '/Dis_Reduction'], ...
    'Inputs','|+-','Position',[755 545 790 585]);
add_block('simulink/Math Operations/Abs', [newSub '/Abs_Chg_Limit'], ...
    'Position',[750 610 790 650]);
add_block('simulink/Math Operations/Sum', [newSub '/Chg_Reduction'], ...
    'Inputs','|+-','Position',[820 605 855 650]);
add_block('simulink/Logic and Bit Operations/Compare To Constant', ...
    [newSub '/Dis_Reduction_Compare'], 'relop','>', 'const','Ilim_eps', ...
    'Position',[820 535 890 590]);
add_block('simulink/Logic and Bit Operations/Compare To Constant', ...
    [newSub '/Chg_Reduction_Compare'], 'relop','>', 'const','Ilim_eps', ...
    'Position',[885 600 955 655]);
add_block('simulink/Logic and Bit Operations/Logical Operator', ...
    [newSub '/Any_Derating'], 'Operator','OR','Inputs','2', ...
    'Position',[985 555 1035 620]);

% Internal wiring
add_line(newSub,'V_bat/1','Vbat_LPF/1','autorouting','on');
add_line(newSub,'Vbat_LPF/1','Abs_Vbat/1','autorouting','on');
add_line(newSub,'Abs_Vbat/1','Vbat_Safe_Max/1','autorouting','on');
add_line(newSub,'Vbat_Floor/1','Vbat_Safe_Max/2','autorouting','on');

add_line(newSub,'Pdis_Max/1','Pdis_Over_Vbat/1','autorouting','on');
add_line(newSub,'Vbat_Safe_Max/1','Pdis_Over_Vbat/2','autorouting','on');
add_line(newSub,'Pdis_Over_Vbat/1','Min_Dis_Power_Current/1','autorouting','on');
add_line(newSub,'Ihw_Dis/1','Min_Dis_Power_Current/2','autorouting','on');
add_line(newSub,'SOC/1','SOC_Dis_Derate/1','autorouting','on');
add_line(newSub,'Min_Dis_Power_Current/1','I_Dis_Limit/1','autorouting','on');
add_line(newSub,'SOC_Dis_Derate/1','I_Dis_Limit/2','autorouting','on');

add_line(newSub,'Pchg_Max/1','Pchg_Over_Vbat/1','autorouting','on');
add_line(newSub,'Vbat_Safe_Max/1','Pchg_Over_Vbat/2','autorouting','on');
add_line(newSub,'Pchg_Over_Vbat/1','Min_Chg_Power_Current/1','autorouting','on');
add_line(newSub,'Ihw_Chg/1','Min_Chg_Power_Current/2','autorouting','on');
add_line(newSub,'SOC/1','SOC_Chg_Derate/1','autorouting','on');
add_line(newSub,'Min_Chg_Power_Current/1','I_Chg_Magnitude/1','autorouting','on');
add_line(newSub,'SOC_Chg_Derate/1','I_Chg_Magnitude/2','autorouting','on');
add_line(newSub,'I_Chg_Magnitude/1','Negative_Charge_Limit/1','autorouting','on');

add_line(newSub,'I_cmd_selected/1','Clamp_Upper/1','autorouting','on');
add_line(newSub,'I_Dis_Limit/1','Clamp_Upper/2','autorouting','on');
add_line(newSub,'Clamp_Upper/1','Clamp_Lower/1','autorouting','on');
add_line(newSub,'Negative_Charge_Limit/1','Clamp_Lower/2','autorouting','on');
add_line(newSub,'Clamp_Lower/1','I_ref_protected/1','autorouting','on');

add_line(newSub,'I_Dis_Limit/1','I_dis_limit_dyn/1','autorouting','on');
add_line(newSub,'Negative_Charge_Limit/1','I_chg_limit_dyn/1','autorouting','on');
add_line(newSub,'I_cmd_selected/1','Command_Minus_Limited/1','autorouting','on');
add_line(newSub,'Clamp_Lower/1','Command_Minus_Limited/2','autorouting','on');
add_line(newSub,'Command_Minus_Limited/1','Abs_Limit_Error/1','autorouting','on');
add_line(newSub,'Abs_Limit_Error/1','Limit_Active_Compare/1','autorouting','on');
add_line(newSub,'Limit_Active_Compare/1','bat_limit_active/1','autorouting','on');

add_line(newSub,'Ihw_Dis/1','Dis_Reduction/1','autorouting','on');
add_line(newSub,'I_Dis_Limit/1','Dis_Reduction/2','autorouting','on');
add_line(newSub,'Dis_Reduction/1','Dis_Reduction_Compare/1','autorouting','on');
add_line(newSub,'Negative_Charge_Limit/1','Abs_Chg_Limit/1','autorouting','on');
add_line(newSub,'Ihw_Chg/1','Chg_Reduction/1','autorouting','on');
add_line(newSub,'Abs_Chg_Limit/1','Chg_Reduction/2','autorouting','on');
add_line(newSub,'Chg_Reduction/1','Chg_Reduction_Compare/1','autorouting','on');
add_line(newSub,'Dis_Reduction_Compare/1','Any_Derating/1','autorouting','on');
add_line(newSub,'Chg_Reduction_Compare/1','Any_Derating/2','autorouting','on');
add_line(newSub,'Any_Derating/1','bat_power_limit_active/1','autorouting','on');
add_line(newSub,'Vbat_Safe_Max/1','Vbat_safe/1','autorouting','on');

%% 2. Top-level connections and logs.
add_block('simulink/Signal Routing/From', [mdl '/Vbat_From_Limit'], ...
    'GotoTag','V_bat','Position',[subPos(1)-125 subPos(2)+50 subPos(1)-25 subPos(2)+80]);
add_line(mdl,'I_ref_From_monitor/1','Battery_Reference_Protection/1','autorouting','on');
add_line(mdl,'Vbat_From_Limit/1','Battery_Reference_Protection/2','autorouting','on');
add_line(mdl,'SOC_realtime_From_aux1/1','Battery_Reference_Protection/3','autorouting','on');
for target = {[mdl '/I_ref_protected'], [mdl '/I_ref_protected_Goto'], ...
        [mdl '/PV_Curtail_Protection']}
    targetPh = get_param(target{1}, 'PortHandles');
    targetPort = targetPh.Inport(1);
    if strcmp(target{1}, [mdl '/PV_Curtail_Protection'])
        targetPort = targetPh.Inport(3);
    end
    targetLine = get_param(targetPort, 'Line');
    if targetLine ~= -1, delete_line(targetLine); end
end
add_line(mdl,'Battery_Reference_Protection/1','I_ref_protected/1','autorouting','on');
add_line(mdl,'Battery_Reference_Protection/1','I_ref_protected_Goto/1','autorouting','on');
add_line(mdl,'Battery_Reference_Protection/1','PV_Curtail_Protection/3','autorouting','on');

logNames = {'I_dis_limit_dyn','I_chg_limit_dyn','bat_limit_active', ...
    'bat_power_limit_active','Vbat_safe_limit'};
for k = 1:numel(logNames)
    p = [subPos(1)+280 subPos(2)+25+45*(k-1) subPos(1)+425 subPos(2)+50+45*(k-1)];
    add_block('simulink/Sinks/To Workspace', [mdl '/Log_' logNames{k}], ...
        'VariableName',logNames{k}, 'SaveFormat','Timeseries', 'Position',p);
    add_line(mdl,sprintf('Battery_Reference_Protection/%d',k+1), ...
        ['Log_' logNames{k} '/1'],'autorouting','on');
end

% Feedforward now uses measured, filtered Vbat instead of nominal 300 V.
dividePh = get_param([mdl '/Divide'],'PortHandles');
oldLine = get_param(dividePh.Inport(2),'Line');
if oldLine ~= -1, delete_line(oldLine); end
add_line(mdl,'Battery_Reference_Protection/6','Divide/2','autorouting','on');

%% 3. Cross-layer tracking for PI_voltage.
set_param([mdl '/PI_voltage'], 'TrackingMode','on', 'Kt','Kt_bat');
piPos = get_param([mdl '/PI_voltage'],'Position');
add_block('simulink/Math Operations/Sum', [mdl '/PI_Track_Equivalent'], ...
    'Inputs','|++','Position',[piPos(1)+15 piPos(2)+125 piPos(1)+45 piPos(2)+175]);
add_block('simulink/Discrete/Unit Delay', [mdl '/PI_Track_z1'], ...
    'SampleTime','Ts_bat_ctrl','InitialCondition','0', ...
    'Position',[piPos(1)+85 piPos(2)+130 piPos(1)+140 piPos(2)+170]);
add_line(mdl,'Battery_Reference_Protection/1','PI_Track_Equivalent/1','autorouting','on');
add_line(mdl,'From69/1','PI_Track_Equivalent/2','autorouting','on');
add_line(mdl,'PI_Track_Equivalent/1','PI_Track_z1/1','autorouting','on');
piPh = get_param([mdl '/PI_voltage'],'PortHandles');
assert(numel(piPh.Inport) >= 2, 'Tracking mode did not expose the expected PI tracking port.');
add_line(mdl,'PI_Track_z1/1','PI_voltage/2','autorouting','on');

add_block('simulink/Sinks/To Workspace', [mdl '/Log_PI_Track'], ...
    'VariableName','I_PI_track','SaveFormat','Timeseries', ...
    'Position',[piPos(1)+165 piPos(2)+130 piPos(1)+290 piPos(2)+155]);
add_line(mdl,'PI_Track_z1/1','Log_PI_Track/1','autorouting','on');

%% 4. Align inner-PI saturation with the reachable final duty command.
set_param([mdl '/Discrete PID Controller'], ...
    'UpperSaturationLimit','Dcorr_max', ...
    'LowerSaturationLimit','Dcorr_min');

%% 5. Add user-visible annotations without hiding any electrical path.
ann = Simulink.Annotation(mdl, ['ENGINEERING HARDENING 2026-08-16\n' ...
    'Dynamic battery P/I/SOC limits after mode selection\n' ...
    'Outer PI tracks the final executable current command\n' ...
    'No ideal DC-bus clamp and no hidden load shedding']);
ann.Position = [subPos(1)-10 subPos(2)-115 subPos(1)+430 subPos(2)-35];
ann.ForegroundColor = 'blue';

%% 6. Save only the experiment copy. Compilation is a separate verification gate.
save_system(mdl, modelFile);

fprintf('Applied Stage 15 battery engineering hardening to: %s\n', modelFile);
fprintf('New subsystem: %s\n', newSub);
fprintf('PI_voltage tracking input count: %d\n', numel(piPh.Inport));
fprintf('Inner PI limits: [%s, %s]\n', ...
    get_param([mdl '/Discrete PID Controller'],'LowerSaturationLimit'), ...
    get_param([mdl '/Discrete PID Controller'],'UpperSaturationLimit'));
bdclose(mdl);
