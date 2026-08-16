%% Correct the grid limiter to the actual 690 V grid_VI measurement side.
expDir=fileparts(mfilename('fullpath'));
mdl='manul_model_stage15_complete';
cd(expDir);
run(fullfile(expDir,'pvess_stage15_busfix_init.m'));
bdclose('all');
load_system(fullfile(expDir,[mdl '.slx']));
sub=[mdl '/Grid_PQ_Current_Limiter'];
assert(getSimulinkBlockHandle(sub)~=-1,'Grid limiter is missing.');

gainBlock=[sub '/Smax_From_Current'];
set_param(gainBlock,'Gain','sqrt(3)*Iac_rated_lv');
oldIn=[sub '/Vabc_grid_hv'];
if getSimulinkBlockHandle(oldIn)~=-1
    set_param(oldIn,'Name','Vabc_grid_lv');
end
set_param(sub,'AttributesFormatString', ...
    ['2.5 MVA grid P/Q current limit\n' ...
    'Measured side: 690 V LV; Q priority']);

set_param(mdl,'SimulationCommand','update');
save_system(mdl);
fprintf('GRID_LIMIT_SIDE_FIX_OK Gain=%s Vfloor=%g V I_rated=%g A\n', ...
    get_param(gainBlock,'Gain'),Vgrid_limit_floor,Iac_rated_lv);
bdclose(mdl);
