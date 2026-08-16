%% Verify Stage 15 engineering structure without running the long scenarios.
expDir = fileparts(mfilename('fullpath'));
mdl = 'manul_model_stage15_complete';
cd(expDir);
run(fullfile(expDir,'pvess_stage15_busfix_init.m'));
bdclose('all');
load_system(fullfile(expDir,[mdl '.slx']));

sub = [mdl '/Battery_Reference_Protection'];
assert(getSimulinkBlockHandle(sub) ~= -1, 'Battery_Reference_Protection missing.');
gridSub = [mdl '/Grid_PQ_Current_Limiter'];
assert(getSimulinkBlockHandle(gridSub) ~= -1, 'Grid_PQ_Current_Limiter missing.');

fprintf('\n=== ENGINEERING STRUCTURE ===\n');
fprintf('Subsystem blocks: %d\n', numel(find_system(sub,'SearchDepth',1,'Type','Block'))-1);
fprintf('Grid limiter blocks: %d\n', numel(find_system(gridSub,'SearchDepth',1,'Type','Block'))-1);
fprintf('PI_voltage TrackingMode=%s Kt=%s AntiWindup=%s Kb=%s\n', ...
    get_param([mdl '/PI_voltage'],'TrackingMode'), ...
    get_param([mdl '/PI_voltage'],'Kt'), ...
    get_param([mdl '/PI_voltage'],'AntiWindupMode'), ...
    get_param([mdl '/PI_voltage'],'Kb'));
fprintf('Inner PI output limits=[%s,%s], antiwindup=%s\n', ...
    get_param([mdl '/Discrete PID Controller'],'LowerSaturationLimit'), ...
    get_param([mdl '/Discrete PID Controller'],'UpperSaturationLimit'), ...
    get_param([mdl '/Discrete PID Controller'],'AntiWindupMode'));

targetBlocks = {sub,[mdl '/Divide'],[mdl '/PI_voltage'],[mdl '/PI_Track_Equivalent'], ...
    [mdl '/PI_Track_z1'],[mdl '/I_ref_protected'],[mdl '/I_ref_protected_Goto'], ...
    [mdl '/PV_Curtail_Protection'],gridSub,[mdl '/Goto5'],[mdl '/Goto2'], ...
    [mdl '/Log_Pref_Command'],[mdl '/Log_Qref_Command']};
for k = 1:numel(targetBlocks)
    fprintf('\n%s\n',targetBlocks{k});
    printConnections(targetBlocks{k});
end

fprintf('\n=== UNCONNECTED INPUT AUDIT ===\n');
scopes = {mdl,sub,gridSub};
unconnectedCount = 0;
for s = 1:numel(scopes)
    blocks = find_system(scopes{s},'SearchDepth',1,'Type','Block');
    for k = 1:numel(blocks)
        if strcmp(blocks{k},scopes{s}), continue; end
        ph = get_param(blocks{k},'PortHandles');
        if isfield(ph,'Inport')
            for p = 1:numel(ph.Inport)
                if get_param(ph.Inport(p),'Line') == -1
                    bt = get_param(blocks{k},'BlockType');
                    % Source, Terminator, Enable and Trigger semantics do not use ordinary input ports.
                    if ~any(strcmp(bt,{'Constant','From','Inport','Ground','Terminator'}))
                        fprintf('UNCONNECTED INPUT: %s port %d\n',blocks{k},p);
                        unconnectedCount = unconnectedCount + 1;
                    end
                end
            end
        end
    end
end
fprintf('Unconnected ordinary input count=%d\n',unconnectedCount);
assert(unconnectedCount==0,'New or top-level ordinary inputs are unconnected.');

fprintf('\n=== UPDATE DIAGRAM ===\n');
set_param(mdl,'SimulationCommand','update');
fprintf('UPDATE_DIAGRAM_OK\n');

save_system(mdl);
bdclose(mdl);

function printConnections(blockPath)
ph = get_param(blockPath,'PortHandles');
if isfield(ph,'Inport')
    for p = 1:numel(ph.Inport)
        ln = get_param(ph.Inport(p),'Line');
        if ln == -1
            fprintf('  IN%d <- UNCONNECTED\n',p);
        else
            sh = get_param(ln,'SrcBlockHandle');
            if sh == -1, fprintf('  IN%d <- INVALID\n',p); else, fprintf('  IN%d <- %s\n',p,getfullname(sh)); end
        end
    end
end
if isfield(ph,'Outport')
    for p = 1:numel(ph.Outport)
        ln = get_param(ph.Outport(p),'Line');
        if ln == -1
            fprintf('  OUT%d -> UNCONNECTED\n',p);
        else
            dh = get_param(ln,'DstBlockHandle');
            for q = 1:numel(dh)
                if dh(q) ~= -1, fprintf('  OUT%d -> %s\n',p,getfullname(dh(q))); end
            end
        end
    end
end
end
