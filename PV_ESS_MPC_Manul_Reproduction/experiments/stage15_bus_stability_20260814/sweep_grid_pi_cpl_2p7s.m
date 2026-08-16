%% Physics-based grid DC-voltage PI bandwidth sweep under the 1 MW DC CPL step.
% Kp=2*zeta*wn*Cdc*Vdc and Ki=wn^2*Cdc*Vdc preserve zeta while changing
% closed-loop bandwidth. No load profile, converter rating, or capacitance
% is changed. The 20 Hz row comes from the already-saved full CPL run.

expDir=fileparts(mfilename('fullpath'));
mdl='manul_model_stage15_complete';
cd(expDir);
run(fullfile(expDir,'pvess_stage15_busfix_init.m'));
bdclose('all');
load_system(fullfile(expDir,[mdl '.slx']));

frequencies=[20 30 40 50];
zeta=1/sqrt(2);
KpValues=2*zeta*(2*pi*frequencies)*Cdc*Vdc_ref;
KiValues=(2*pi*frequencies).^2*Cdc*Vdc_ref;
n=numel(frequencies);
minimumVdc=nan(n,1); maximumVdc=nan(n,1); recovery_s=nan(n,1);
steadyMean=nan(n,1); steadyStd=nan(n,1); preEventStd=nan(n,1);
p2Mean=nan(n,1); p2ErrorPct=nan(n,1); gridLimitDuration=nan(n,1);
prefMin=nan(n,1); prefMax=nan(n,1); allFinite=false(n,1);
runs=struct([]);

% Baseline 20 Hz metrics from the completed 4 s CPL result.
saved=load(fullfile(expDir,'stage15_complete_results.mat'),'results');
base=saved.results.signals;
[minimumVdc(1),maximumVdc(1),recovery_s(1),steadyMean(1),steadyStd(1), ...
    preEventStd(1),p2Mean(1),p2ErrorPct(1),gridLimitDuration(1), ...
    prefMin(1),prefMax(1),allFinite(1)]=local_metrics(base);
runs(1).frequency_Hz=frequencies(1);
runs(1).Kp=KpValues(1); runs(1).Ki=KiValues(1);
runs(1).V_dc=base.V_dc;

tw=find_system(mdl,'LookUnderMasks','all','FollowLinks','on','BlockType','ToWorkspace');
oldDec=cell(size(tw));
for k=1:numel(tw)
    oldDec{k}=get_param(tw{k},'Decimation');
    set_param(tw{k},'Decimation','320');
end

for k=2:n
    fprintf('GRID_PI_SWEEP_START f=%.0fHz Kp=%.6g Ki=%.6g\n', ...
        frequencies(k),KpValues(k),KiValues(k));
    in=Simulink.SimulationInput(mdl);
    in=in.setVariable('Kp_dc',KpValues(k));
    in=in.setVariable('Ki_dc',KiValues(k));
    in=in.setModelParameter('StartTime','0','StopTime','2.70', ...
        'SignalLogging','off','SaveOutput','off','SaveState','off');
    timer=tic;
    out=sim(in);
    elapsed=toc(timer);
    sig=struct();
    needed={'V_dc','P_DC_load2','Pdc2_cpl_command','grid_current_limit_active','P_ref_cmd'};
    for j=1:numel(needed), sig.(needed{j})=out.get(needed{j}); end
    [minimumVdc(k),maximumVdc(k),recovery_s(k),steadyMean(k),steadyStd(k), ...
        preEventStd(k),p2Mean(k),p2ErrorPct(k),gridLimitDuration(k), ...
        prefMin(k),prefMax(k),allFinite(k)]=local_metrics(sig);
    runs(k).frequency_Hz=frequencies(k); %#ok<SAGROW>
    runs(k).Kp=KpValues(k); runs(k).Ki=KiValues(k);
    runs(k).elapsed_s=elapsed; runs(k).V_dc=sig.V_dc;
    fprintf(['GRID_PI_SWEEP_DONE f=%.0fHz elapsed=%.1fs Vmin=%.3f V ' ...
        'recovery=%.6fs steadyStd=%.3fV P2err=%.5f%%\n'], ...
        frequencies(k),elapsed,minimumVdc(k),recovery_s(k),steadyStd(k),p2ErrorPct(k));
end

for k=1:numel(tw), set_param(tw{k},'Decimation',oldDec{k}); end

sweepTable=table(frequencies(:),KpValues(:),KiValues(:),minimumVdc,maximumVdc, ...
    recovery_s,steadyMean,steadyStd,preEventStd,p2Mean,p2ErrorPct, ...
    gridLimitDuration,prefMin,prefMax,allFinite, ...
    'VariableNames',{'Bandwidth_Hz','Kp_W_per_V','Ki_W_per_Vs','VdcMin_V', ...
    'VdcMax_V','Recovery_s','SteadyMean_V','SteadyStd_V','PreEventStd_V', ...
    'Pdc2Mean_W','Pdc2Error_percent','GridLimitDuration_s', ...
    'PrefMin_W','PrefMax_W','AllFinite'});
disp(sweepTable);
writetable(sweepTable,fullfile(expDir,'stage15_grid_pi_cpl_sweep.csv'));
save(fullfile(expDir,'stage15_grid_pi_cpl_sweep.mat'),'sweepTable','runs','-v7.3');
fid=fopen(fullfile(expDir,'stage15_grid_pi_cpl_sweep.json'),'w');
fprintf(fid,'%s',jsonencode(table2struct(sweepTable),'PrettyPrint',true)); fclose(fid);

set(groot,'defaultFigureVisible','off');
f=figure('Color','w','Position',[50 50 1250 720]);
tiledlayout(2,1,'TileSpacing','compact','Padding','compact');
nexttile;
for k=1:n
    plot(runs(k).V_dc.Time,runs(k).V_dc.Data,'LineWidth',1.0); hold on;
end
xlim([2.48 2.62]); yline(1200,'--k'); yline(0.7*1200,':r'); grid on;
ylabel('V_{dc} (V)'); title('Grid outer-PI bandwidth sweep: 1 MW DC CPL');
legend(compose('%g Hz',frequencies),'Location','best');
nexttile;
plot(frequencies,minimumVdc,'o-','LineWidth',1.2); grid on;
xlabel('Designed outer-loop bandwidth (Hz)'); ylabel('Minimum V_{dc} (V)');
exportgraphics(f,fullfile(expDir,'stage15_grid_pi_cpl_sweep.png'),'Resolution',180);
close(f);
close_system(mdl,0);

function [vmin,vmax,recovery,meanSteady,stdSteady,stdPre,p2mean,p2err, ...
    limitDuration,prefmin,prefmax,finiteFlag]=local_metrics(sig)
v=sig.V_dc; t=v.Time(:); x=reshape(v.Data,[],1);
event=t>=2.5 & t<2.58;
steady=t>=2.58 & t<=2.70;
pre=t>=2.20 & t<2.40;
vmin=min(x(event)); vmax=max(x(event));
meanSteady=mean(x(steady)); stdSteady=std(x(steady)); stdPre=std(x(pre));
recovery=local_recovery(t,x,2.5,1200,0.02,0.020);
p=sig.P_DC_load2; pmask=p.Time>=2.58 & p.Time<=2.70;
pc=sig.Pdc2_cpl_command; pcmask=pc.Time>=2.58 & pc.Time<=2.70;
p2mean=mean(reshape(p.Data(pmask),[],1));
pcmean=mean(reshape(pc.Data(pcmask),[],1));
p2err=100*(p2mean-pcmean)/pcmean;
g=sig.grid_current_limit_active;
limitDuration=trapz(g.Time,double(g.Data(:)>0.5));
pr=sig.P_ref_cmd; prefmin=min(pr.Data(:)); prefmax=max(pr.Data(:));
finiteFlag=all(isfinite(x)) && all(isfinite(p.Data(:))) && all(isfinite(pr.Data(:)));
end

function recovery=local_recovery(t,x,eventTime,reference,tolerancePu,dwell_s)
inside=abs(x-reference)<=tolerancePu*reference & t>=eventTime;
recovery=NaN;
indices=find(t>=eventTime);
for ii=1:numel(indices)
    first=indices(ii);
    last=find(t>=t(first)+dwell_s,1,'first');
    if isempty(last), break; end
    if all(inside(first:last))
        recovery=t(first)-eventTime;
        return
    end
end
end
