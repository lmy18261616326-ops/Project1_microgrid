%% Read-only inventory of AC/DC load-related blocks and their connections.
expDir=fileparts(mfilename('fullpath'));
mdl='manul_model_stage15_complete';
cd(expDir);
run(fullfile(expDir,'pvess_stage15_busfix_init.m'));
bdclose('all');
load_system(fullfile(expDir,[mdl '.slx']));

blocks=find_system(mdl,'LookUnderMasks','all','FollowLinks','on','Type','Block');
fprintf('LOAD_RELATED_BLOCKS_BEGIN\n');
for k=1:numel(blocks)
    b=blocks{k};
    name=get_param(b,'Name');
    bt=get_param(b,'BlockType');
    mt=get_param(b,'MaskType');
    rt=get_param(b,'ReferenceBlock');
    hay=lower(strjoin({name,bt,mt,rt},' '));
    if contains(hay,'load') || contains(hay,'rlc') || contains(hay,'resistor') || ...
            contains(hay,'current source') || contains(hay,'breaker')
        fprintf('\nBLOCK=%s\n  BlockType=%s\n  MaskType=%s\n  ReferenceBlock=%s\n', ...
            b,bt,mt,rt);
        dp=get_param(b,'DialogParameters');
        if ~isempty(dp)
            fn=fieldnames(dp);
            keep={'Resistance','Inductance','Capacitance','ActivePower','ReactivePower', ...
                'NominalVoltage','NominalFrequency','Pn','Qn','Vn','fn','Measurements', ...
                'SwitchTimes','InitialStatus','External','LoadType','Power'};
            for j=1:numel(fn)
                if any(strcmpi(fn{j},keep)) || contains(lower(fn{j}),'power') || ...
                        contains(lower(fn{j}),'resist') || contains(lower(fn{j}),'switch')
                    try
                        fprintf('  %s=%s\n',fn{j},get_param(b,fn{j}));
                    catch
                    end
                end
            end
        end
        pc=get_param(b,'PortConnectivity');
        for p=1:numel(pc)
            dst=pc(p).DstBlock;
            src=pc(p).SrcBlock;
            fprintf('  Port%d Type=%s Src=%s Dst=%s\n',p,pc(p).Type, ...
                local_block_names(src),local_block_names(dst));
        end
    end
end
fprintf('\nLOAD_RELATED_BLOCKS_END\n');

fprintf('\nTOP_LEVEL_SUBSYSTEMS_BEGIN\n');
tops=find_system(mdl,'SearchDepth',1,'Type','Block');
for k=2:numel(tops)
    b=tops{k};
    fprintf('%s | %s | %s\n',b,get_param(b,'BlockType'),get_param(b,'MaskType'));
end
fprintf('TOP_LEVEL_SUBSYSTEMS_END\n');
close_system(mdl,0);

function txt=local_block_names(handles)
if isempty(handles)
    txt='[]';
    return
end
names=cell(1,numel(handles));
for ii=1:numel(handles)
    try
        names{ii}=getfullname(handles(ii));
    catch
        names{ii}=num2str(handles(ii));
    end
end
txt=strjoin(names,';');
end
