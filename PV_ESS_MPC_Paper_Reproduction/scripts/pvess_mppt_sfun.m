function pvess_mppt_sfun(block)
%PVESS_MPPT_SFUN Incremental-conductance MPPT for the PV boost converter.
setup(block);
end

function setup(block)
block.NumDialogPrms = 6; % Ts, D0, Dmin, Dmax, step, Doff
block.NumInputPorts = 1;
block.InputPort(1).Dimensions = 3; % [Vpv Ipv offMPPT]
block.InputPort(1).DirectFeedthrough = true;
block.NumOutputPorts = 1;
block.OutputPort(1).Dimensions = 1;
block.OutputPort(1).SamplingMode = 'Sample';
block.SampleTimes = [block.DialogPrm(1).Data 0];
block.RegBlockMethod('PostPropagationSetup',@setupDwork);
block.RegBlockMethod('InitializeConditions',@initialize);
block.RegBlockMethod('Outputs',@outputs);
block.RegBlockMethod('Update',@update);
end

function setupDwork(block)
block.NumDworks = 3;
for k = 1:3
    block.Dwork(k).Dimensions = 1;
    block.Dwork(k).DatatypeID = 0;
    block.Dwork(k).Complexity = 'Real';
    block.Dwork(k).UsedAsDiscState = true;
end
block.Dwork(1).Name = 'Vprev';
block.Dwork(2).Name = 'Iprev';
block.Dwork(3).Name = 'Duty';
end

function initialize(block)
block.Dwork(1).Data = 0;
block.Dwork(2).Data = 0;
block.Dwork(3).Data = block.DialogPrm(2).Data;
end

function outputs(block)
block.OutputPort(1).Data = block.Dwork(3).Data;
end

function update(block)
u = block.InputPort(1).Data;
v = max(u(1),1);
i = u(2);
offMppt = u(3) > 0.5;
d = block.Dwork(3).Data;
if offMppt
    d = block.DialogPrm(6).Data;
else
    dv = v - block.Dwork(1).Data;
    di = i - block.Dwork(2).Data;
    step = block.DialogPrm(5).Data;
    if abs(dv) < 1e-6
        if di > 1e-3
            d = d - step;
        elseif di < -1e-3
            d = d + step;
        end
    else
        inc = di/dv + i/v;
        if inc > 1e-4
            d = d - step; % left of MPP: raise Vpv by reducing boost duty
        elseif inc < -1e-4
            d = d + step;
        end
    end
end
d = min(max(d,block.DialogPrm(3).Data),block.DialogPrm(4).Data);
block.Dwork(1).Data = v;
block.Dwork(2).Data = i;
block.Dwork(3).Data = d;
end
