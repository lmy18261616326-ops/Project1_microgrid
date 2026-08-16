modelFile = fullfile(fileparts(mfilename('fullpath')), ...
    'manul_model_stage15_busfix.slx');
[~, modelName] = fileparts(modelFile);
load_system(modelFile);

sourcePath = [modelName '/Three-Phase Source'];
fprintf('CURRENT_SOURCE=%s\n', sourcePath);
fprintf('CURRENT_MASK=%s\n', get_param(sourcePath, 'MaskType'));
fprintf('CURRENT_MASK_NAMES=%s\n', strjoin(get_param(sourcePath, 'MaskNames'), ','));
fprintf('CURRENT_MASK_VALUES=%s\n', strjoin(get_param(sourcePath, 'MaskValues'), '|'));

portHandles = get_param(sourcePath, 'PortHandles');
connectorFields = {'LConn', 'RConn'};
for fieldIndex = 1:numel(connectorFields)
    fieldName = connectorFields{fieldIndex};
    handles = portHandles.(fieldName);
    for portIndex = 1:numel(handles)
        lineHandle = get_param(handles(portIndex), 'Line');
        fprintf('%s%d line=%g', fieldName, portIndex, lineHandle);
        if lineHandle ~= -1
            sourceHandle = get_param(lineHandle, 'SrcBlockHandle');
            destinationHandles = get_param(lineHandle, 'DstBlockHandle');
            if sourceHandle ~= -1
                fprintf(' src=%s', getfullname(sourceHandle));
            end
            for destinationIndex = 1:numel(destinationHandles)
                fprintf(' dst=%s', getfullname(destinationHandles(destinationIndex)));
            end
        end
        fprintf('\n');
    end
end

libraryName = 'spsThreePhaseProgrammableVoltageSourceLib';
load_system(libraryName);
libraryBlock = [libraryName '/Three-Phase' newline 'Programmable' newline 'Voltage Source'];
fprintf('PROGRAMMABLE_BLOCK=%s\n', libraryBlock);
maskNames = get_param(libraryBlock, 'MaskNames');
maskValues = get_param(libraryBlock, 'MaskValues');
for maskIndex = 1:numel(maskNames)
    fprintf('PARAM %s=%s\n', maskNames{maskIndex}, maskValues{maskIndex});
end

programmablePorts = get_param(libraryBlock, 'PortHandles');
fprintf('PROGRAMMABLE_LCONN=%d RCONN=%d\n', ...
    numel(programmablePorts.LConn), numel(programmablePorts.RConn));

branchLibrary = 'spsThreePhaseSeriesRLCBranchLib';
load_system(branchLibrary);
branchBlocks = find_system(branchLibrary, 'SearchDepth', 1, 'Type', 'Block');
for branchIndex = 1:numel(branchBlocks)
    fprintf('BRANCH %s | MaskType=%s\n', branchBlocks{branchIndex}, ...
        get_param(branchBlocks{branchIndex}, 'MaskType'));
    branchMaskNames = get_param(branchBlocks{branchIndex}, 'MaskNames');
    branchMaskValues = get_param(branchBlocks{branchIndex}, 'MaskValues');
    for maskIndex = 1:numel(branchMaskNames)
        fprintf('  PARAM %s=%s\n', branchMaskNames{maskIndex}, branchMaskValues{maskIndex});
    end
    branchPorts = get_param(branchBlocks{branchIndex}, 'PortHandles');
    fprintf('  LCONN=%d RCONN=%d\n', numel(branchPorts.LConn), numel(branchPorts.RConn));
end

close_system(modelName, 0);
