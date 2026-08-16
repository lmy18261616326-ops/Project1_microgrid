%% Targeted read-only inspection of load blocks and switching devices.
expDir=fileparts(mfilename('fullpath'));
mdl='manul_model_stage15_complete';
cd(expDir);
run(fullfile(expDir,'pvess_stage15_busfix_init.m'));
bdclose('all');
load_system(fullfile(expDir,[mdl '.slx']));

tops=find_system(mdl,'SearchDepth',1,'Type','Block');
tokens={'dc_load','parallel rlc load','breaker2','ac_load','idc_load','vabc_load','iabc_load'};
for k=2:numel(tops)
    b=tops{k};
    name=get_param(b,'Name');
    mt=get_param(b,'MaskType');
    hay=lower([name ' ' mt]);
    if ~any(cellfun(@(s)contains(hay,s),tokens)), continue; end
    fprintf('\n=== %s ===\n',b);
    fprintf('BlockType=%s MaskType=%s ReferenceBlock=%s Position=%s\n', ...
        get_param(b,'BlockType'),mt,get_param(b,'ReferenceBlock'),mat2str(get_param(b,'Position')));
    dp=get_param(b,'DialogParameters');
    if ~isempty(dp)
        fn=fieldnames(dp);
        for j=1:numel(fn)
            try
                val=get_param(b,fn{j});
                if ischar(val) && numel(val)<300
                    fprintf('PARAM %s=%s\n',fn{j},strrep(val,newline,'<NL>'));
                end
            catch
            end
        end
    end
    pc=get_param(b,'PortConnectivity');
    for p=1:numel(pc)
        fprintf('PORT %d Type=%s Src=%s Dst=%s\n',p,pc(p).Type, ...
            local_names(pc(p).SrcBlock),local_names(pc(p).DstBlock));
    end
end
close_system(mdl,0);

function txt=local_names(handles)
if isempty(handles), txt='[]'; return; end
names=cell(1,numel(handles));
for ii=1:numel(handles)
    try, names{ii}=getfullname(handles(ii)); catch, names{ii}=num2str(handles(ii)); end
end
txt=strjoin(names,';');
end
