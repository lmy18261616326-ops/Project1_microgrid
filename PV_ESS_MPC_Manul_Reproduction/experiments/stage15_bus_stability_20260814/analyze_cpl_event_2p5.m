%% Diagnose the physically demanding 1 MW DC-CPL connection at 2.5 s.
expDir=fileparts(mfilename('fullpath'));
data=load(fullfile(expDir,'stage15_complete_results.mat'),'results');
r=data.results;
queryTime=[2.4990;2.5000;2.5005;2.5020;2.5050;2.51072;2.5200;2.5500;2.6000];
names={'V_dc','P_DC','P_pv','P_bat','Pf_threephase','P_grid_threephase', ...
    'P_ref_raw','P_ref_cmd','S_grid_current_max','grid_current_limit_active', ...
    'Rdc1_cpl_command','Rdc2_cpl_command'};
values=zeros(numel(queryTime),numel(names));
for k=1:numel(names)
    ts=r.signals.(names{k});
    values(:,k)=interp1(ts.Time(:),double(reshape(ts.Data,[],1)),queryTime,'previous','extrap');
end
eventTable=array2table([queryTime values],'VariableNames',[{'Time_s'} names]);
disp(eventTable);

r1=r.signals.Rdc1_cpl_command;
r2=r.signals.Rdc2_cpl_command;
mask1=r1.Time>=1.0 & r1.Time<3.0;
mask2=r2.Time>=2.5 & r2.Time<3.0;
fprintf('CPL_RESISTANCE_MIN R1=%.6f ohm R2=%.6f ohm FLOOR_EQ=%.6f ohm\n', ...
    min(r1.Data(mask1)),min(r2.Data(mask2)),(0.70*1200)^2/1e6);

set(groot,'defaultFigureVisible','off');
f=figure('Color','w','Position',[50 50 1300 850]);
tiledlayout(3,1,'TileSpacing','compact','Padding','compact');
nexttile;
plot(r.signals.V_dc.Time,r.signals.V_dc.Data,'LineWidth',1.1); hold on;
yline(1200,'--k'); yline(0.7*1200,':r','CPL protection floor');
xlim([2.48 2.62]); grid on; ylabel('V_{dc} (V)'); title('1 MW DC-CPL connection');
nexttile;
sig={'P_DC','P_pv','P_bat','Pf_threephase','P_grid_threephase'};
for k=1:numel(sig)
    ts=r.signals.(sig{k}); plot(ts.Time,reshape(ts.Data,[],1)/1e6,'LineWidth',1.0); hold on;
end
xlim([2.48 2.62]); grid on; ylabel('Power (MW)');
legend({'DC load','PV','Battery','Interlink AC','Grid'},'Location','best');
nexttile;
for sig={'P_ref_raw','P_ref_cmd'}
    ts=r.signals.(sig{1}); plot(ts.Time,reshape(ts.Data,[],1)/1e6,'LineWidth',1.0); hold on;
end
xlim([2.48 2.62]); grid on; ylabel('P reference (MW)'); xlabel('Time (s)');
legend({'Raw','After current limit'},'Location','best');
exportgraphics(f,fullfile(expDir,'stage15_cpl_1MW_event_zoom.png'),'Resolution',180);
close(f);
