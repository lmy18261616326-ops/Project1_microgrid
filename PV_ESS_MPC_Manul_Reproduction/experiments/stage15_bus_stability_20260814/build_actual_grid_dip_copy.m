experimentDirectory = fileparts(mfilename('fullpath'));
sourceModelFile = fullfile(experimentDirectory, ...
    'manul_model_stage15_busfix.slx');
targetModelFile = fullfile(experimentDirectory, ...
    'manul_model_stage15_complete.slx');

assert(~isfile(targetModelFile), ...
    'Target model already exists; refusing to overwrite: %s', targetModelFile);
copyfile(sourceModelFile, targetModelFile);

[~, modelName] = fileparts(targetModelFile);
load_system(targetModelFile);

oldSource = [modelName '/Three-Phase Source'];
oldPosition = get_param(oldSource, 'Position');
oldPorts = get_param(oldSource, 'PortHandles');
assert(numel(oldPorts.RConn) == 3, ...
    'Expected three phase connectors on the original grid source.');

transformerPorts = zeros(1, 3);
for phaseIndex = 1:3
    sourcePort = oldPorts.RConn(phaseIndex);
    lineHandle = get_param(sourcePort, 'Line');
    assert(lineHandle ~= -1, 'Original grid phase %d is unconnected.', phaseIndex);
    lineSourcePort = get_param(lineHandle, 'SrcPortHandle');
    lineDestinationPort = get_param(lineHandle, 'DstPortHandle');
    if lineSourcePort == sourcePort
        transformerPorts(phaseIndex) = lineDestinationPort;
    else
        transformerPorts(phaseIndex) = lineSourcePort;
    end
    delete_line(lineHandle);
end
delete_block(oldSource);

programmableLibrary = 'spsThreePhaseProgrammableVoltageSourceLib';
branchLibrary = 'spsThreePhaseSeriesRLCBranchLib';
groundLibrary = 'spsGroundLib';
load_system(programmableLibrary);
load_system(branchLibrary);
load_system(groundLibrary);

programmableCandidates = find_system(programmableLibrary, ...
    'SearchDepth', 1, 'Type', 'Block', ...
    'MaskType', 'Three-Phase Programmable Voltage Source');
branchCandidates = find_system(branchLibrary, ...
    'SearchDepth', 1, 'Type', 'Block', ...
    'MaskType', 'Three-Phase Series RLC Branch');
groundCandidates = find_system(groundLibrary, ...
    'SearchDepth', 1, 'Type', 'Block');
assert(numel(programmableCandidates) == 1, ...
    'Could not uniquely locate the standard programmable source block.');
assert(numel(branchCandidates) == 1, ...
    'Could not uniquely locate the standard three-phase series branch.');
assert(~isempty(groundCandidates), ...
    'Could not locate the standard Specialized Power Systems ground block.');

sourceWidth = oldPosition(3) - oldPosition(1);
sourceHeight = oldPosition(4) - oldPosition(2);
newSourcePosition = oldPosition + [-250 0 -250 0];
branchPosition = [oldPosition(1) - 80, oldPosition(2), ...
    oldPosition(1) + sourceWidth + 40, oldPosition(2) + sourceHeight];
groundPosition = [newSourcePosition(1) - 15, newSourcePosition(4) + 45, ...
    newSourcePosition(1) + 25, newSourcePosition(4) + 85];

newSource = [modelName '/Grid_Programmable_Voltage_Source'];
gridImpedance = [modelName '/Grid_Equivalent_RL'];
gridGround = [modelName '/Grid_Source_Ground'];

add_block(programmableCandidates{1}, newSource, ...
    'Position', newSourcePosition);
set_param(newSource, ...
    'PositiveSequence', '[V_tr_hv 0 f_grid]', ...
    'VariationEntity', 'Amplitude', ...
    'VariationType', 'Table of time-amplitude pairs', ...
    'Amplitudes', '[1 1 0.9 0.9]', ...
    'TimeValues', '[0 T_grid_dip T_grid_dip+Ts_power 4]', ...
    'VariationPhaseA', 'off', ...
    'HarmonicGeneration', 'off', ...
    'BusType', 'swing', ...
    'Pref', '10e3', ...
    'Qref', '0');

add_block(branchCandidates{1}, gridImpedance, ...
    'Position', branchPosition);
set_param(gridImpedance, ...
    'BranchType', 'RL', ...
    'Resistance', '0.8929', ...
    'Inductance', '16.58e-3', ...
    'Measurements', 'None');

add_block(groundCandidates{1}, gridGround, ...
    'Position', groundPosition);

sourcePorts = get_param(newSource, 'PortHandles');
branchPorts = get_param(gridImpedance, 'PortHandles');
groundPorts = get_param(gridGround, 'PortHandles');
assert(numel(sourcePorts.RConn) == 3, ...
    'Programmable source does not expose the expected three phase ports.');
assert(numel(branchPorts.LConn) == 3 && numel(branchPorts.RConn) == 3, ...
    'Grid impedance block does not expose six phase connectors.');

if ~isempty(groundPorts.LConn)
    groundPort = groundPorts.LConn(1);
else
    groundPort = groundPorts.RConn(1);
end
add_line(modelName, sourcePorts.LConn(1), groundPort, 'autorouting', 'smart');

for phaseIndex = 1:3
    add_line(modelName, sourcePorts.RConn(phaseIndex), ...
        branchPorts.LConn(phaseIndex), 'autorouting', 'smart');
    add_line(modelName, branchPorts.RConn(phaseIndex), ...
        transformerPorts(phaseIndex), 'autorouting', 'smart');
end

set_param(modelName, 'SimulationCommand', 'update');
save_system(modelName, targetModelFile);

fprintf('GRID_DIP_COPY_OK=%s\n', targetModelFile);
fprintf('PROGRAMMABLE_SOURCE=%s\n', newSource);
fprintf('GRID_IMPEDANCE=%s\n', gridImpedance);
fprintf('MODEL_DIRTY=%s\n', get_param(modelName, 'Dirty'));
close_system(modelName, 0);
