function pvess_battery_sfun(block)
%PVESS_BATTERY_SFUN Cascaded dc-link/current and SOC-oriented controller.
setup(block);
end

function setup(block)
% Ts,VdcRef,SOCmin,SOCmax,Imax,Kpv,Kiv,Kpi,Kii,Vbat,Dmin,Dmax
block.NumDialogPrms = 12;
block.NumInputPorts = 1;
block.InputPort(1).Dimensions = 5; % grid, Vdc, Ibat, SOC, battCmd
block.InputPort(1).DirectFeedthrough = true;
block.NumOutputPorts = 1;
block.OutputPort(1).Dimensions = 2; % duty, Iref
block.SampleTimes = [block.DialogPrm(1).Data 0];
block.RegBlockMethod('PostPropagationSetup',@setupDwork);
block.RegBlockMethod('InitializeConditions',@initialize);
block.RegBlockMethod('Outputs',@outputs);
block.RegBlockMethod('Update',@update);
end

function setupDwork(block)
block.NumDworks = 2;
for k = 1:2
    block.Dwork(k).Dimensions = 1;
    block.Dwork(k).DatatypeID = 0;
    block.Dwork(k).Complexity = 'Real';
    block.Dwork(k).UsedAsDiscState = true;
end
block.Dwork(1).Name = 'IntV';
block.Dwork(2).Name = 'IntI';
end

function initialize(block)
block.Dwork(1).Data = 0;
block.Dwork(2).Data = 0;
end

function outputs(block)
u = block.InputPort(1).Data;
gridTied = u(1) > 0.5;
vdc = u(2); ibat = u(3); soc = u(4); cmd = u(5);
vref = block.DialogPrm(2).Data;
socMin = block.DialogPrm(3).Data;
socMax = block.DialogPrm(4).Data;
imax = block.DialogPrm(5).Data;
kpv = block.DialogPrm(6).Data;
kpi = block.DialogPrm(8).Data;
vbat = block.DialogPrm(10).Data;
dmin = block.DialogPrm(11).Data;
dmax = block.DialogPrm(12).Data;
if gridTied
    frac = min(max((soc - socMin)/(socMax - socMin),0),1);
    if cmd > 0.5
        iref = imax*frac;
    elseif cmd < -0.5
        iref = -imax*(1-frac);
    else
        iref = 0;
    end
else
    iref = kpv*(vref-vdc) + block.Dwork(1).Data;
    iref = min(max(iref,-imax),imax);
end
d0 = 1 - vbat/max(vdc,1);
duty = d0 + kpi*(iref-ibat) + block.Dwork(2).Data;
duty = min(max(duty,dmin),dmax);
block.OutputPort(1).Data = [duty;iref];
end

function update(block)
u = block.InputPort(1).Data;
ts = block.DialogPrm(1).Data;
vref = block.DialogPrm(2).Data;
imax = block.DialogPrm(5).Data;
kiv = block.DialogPrm(7).Data;
kii = block.DialogPrm(9).Data;
gridTied = u(1) > 0.5;
if ~gridTied
    block.Dwork(1).Data = block.Dwork(1).Data + kiv*ts*(vref-u(2));
    block.Dwork(1).Data = min(max(block.Dwork(1).Data,-imax),imax);
end
y = block.OutputPort(1).Data;
iref = y(2);
block.Dwork(2).Data = block.Dwork(2).Data + kii*ts*(iref-u(3));
block.Dwork(2).Data = min(max(block.Dwork(2).Data,-0.25),0.25);
end
