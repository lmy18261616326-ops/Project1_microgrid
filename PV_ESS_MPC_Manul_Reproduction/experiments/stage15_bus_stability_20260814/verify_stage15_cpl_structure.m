%% Structural/electrical-interface audit for the Stage 15 CPL implementation.
expDir=fileparts(mfilename('fullpath'));
mdl='manul_model_stage15_complete';
cd(expDir);
run(fullfile(expDir,'pvess_stage15_busfix_init.m'));
bdclose('all');
load_system(fullfile(expDir,[mdl '.slx']));
sub=[mdl '/DC_CPL_Resistance_Command'];
assert(getSimulinkBlockHandle(sub)~=-1,'DC CPL subsystem is missing.');

% Every ordinary Simulink input in the new subsystem must be wired.
blocks=find_system(sub,'SearchDepth',1,'Type','Block');
unconnected={};
for k=2:numel(blocks)
    ph=get_param(blocks{k},'PortHandles');
    if isfield(ph,'Inport')
        for p=1:numel(ph.Inport)
            if get_param(ph.Inport(p),'Line')==-1
                unconnected{end+1}=sprintf('%s/In%d',blocks{k},p); %#ok<AGROW>
            end
        end
    end
end
assert(isempty(unconnected),'Unconnected CPL inputs: %s',strjoin(unconnected,', '));

% Confirm no custom-code block was introduced in the CPL subsystem.
matlabFcn=find_system(sub,'LookUnderMasks','all','FollowLinks','on','BlockType','MATLABFcn');
assert(isempty(matlabFcn),'MATLAB Fcn blocks are not allowed in the CPL subsystem.');

% Confirm top-level source and electrical-load command wiring.
for k=1:2
    src=[mdl sprintf('/Pdc%d_CPL_Profile',k)];
    assert(strcmp(get_param(src,'VariableName'),sprintf('Pdc%d_cpl_profile',k)));
    loadBlock=[mdl sprintf('/DC_load%d',k)];
    ph=get_param(loadBlock,'PortHandles');
    line=get_param(ph.Inport(1),'Line');
    assert(line~=-1,'DC_load%d control input is unconnected.',k);
    sourcePort=get_param(line,'SrcPortHandle');
    sourceBlock=get_param(sourcePort,'Parent');
    assert(strcmp(sourceBlock,sub),'DC_load%d is not driven by the CPL subsystem.',k);
end

acLoads=find_system(mdl,'SearchDepth',1,'MaskType','Three-Phase Parallel RLC Load');
activeCount=0;
for k=1:numel(acLoads)
    p=get_param(acLoads{k},'ActivePower');
    if strcmp(p,'P_ac_base') || strcmp(p,'P_ac_step')
        activeCount=activeCount+1;
        assert(strcmp(get_param(acLoads{k},'LoadType'),'constant PQ'), ...
            '%s is not in constant-PQ mode.',acLoads{k});
    end
end
assert(activeCount==2,'Did not find both active-power AC loads.');

set_param(mdl,'SimulationCommand','update');
fprintf('CPL_STRUCTURE_OK blocks=%d unconnected=%d AC_constantPQ=%d\n', ...
    numel(blocks)-1,numel(unconnected),activeCount);
for k=1:2
    loadBlock=[mdl sprintf('/DC_load%d',k)];
    ph=get_param(loadBlock,'PortHandles');
    line=get_param(ph.Inport(1),'Line');
    fprintf('DC_LOAD%d_CONTROL_SOURCE=%s\n',k, ...
        get_param(get_param(line,'SrcPortHandle'),'Parent'));
end
close_system(mdl,0);
