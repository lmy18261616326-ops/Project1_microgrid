function pvess_fcs_sfun(block)
%PVESS_FCS_SFUN One-step FCS MPVC/MPPC for a two-level converter.
setup(block);
end

function setup(block)
% Ts,Lf,Cf,Rf,Pbase,Qbase,lambda
block.NumDialogPrms = 7;
block.NumInputPorts = 1;
block.InputPort(1).Dimensions = 19;
block.InputPort(1).DirectFeedthrough = true;
block.NumOutputPorts = 1;
block.OutputPort(1).Dimensions = 8; % six gates, best cost, state index
block.SampleTimes = [block.DialogPrm(1).Data 0];
block.RegBlockMethod('PostPropagationSetup',@setupDwork);
block.RegBlockMethod('InitializeConditions',@initialize);
block.RegBlockMethod('Outputs',@outputs);
block.RegBlockMethod('Update',@update);
end

function setupDwork(block)
block.NumDworks = 1;
block.Dwork(1).Dimensions = 3;
block.Dwork(1).DatatypeID = 0;
block.Dwork(1).Complexity = 'Real';
block.Dwork(1).UsedAsDiscState = true;
block.Dwork(1).Name = 'PreviousState';
end

function initialize(block)
block.Dwork(1).Data = [0;0;0];
end

function outputs(block)
u = block.InputPort(1).Data;
mode = u(1) > 0.5;
vdc = max(u(2),1);
vc = clarke(u(3:5));
ifil = clarke(u(6:8));
iload = clarke(u(9:11));
vg = clarke(u(12:14));
vref = clarke(u(15:17));
pref = u(18); qref = u(19);
ts = block.DialogPrm(1).Data;
lf = block.DialogPrm(2).Data;
cf = block.DialogPrm(3).Data;
rf = block.DialogPrm(4).Data;
pbase = block.DialogPrm(5).Data;
qbase = block.DialogPrm(6).Data;
lambda = block.DialogPrm(7).Data;

states = [0 0 0; 1 0 0; 1 1 0; 0 1 0; ...
          0 1 1; 0 0 1; 1 0 1; 1 1 1];
bestCost = inf; best = 1;
prev = block.Dwork(1).Data(:).';

% Exact discrete LC model used by Eq. (11).
A = [0 1/cf;-1/lf -rf/lf];
B = [0 -1/cf;1/lf 0];
Ad = expm(A*ts);
Bd = A\((Ad-eye(2))*B);

for k = 1:8
    s = states(k,:);
    vi = (2/3)*vdc*[s(1)-0.5*s(2)-0.5*s(3); ...
        (sqrt(3)/2)*(s(2)-s(3))];
    nsw = sum(abs(s-prev));
    if ~mode
        xa = Ad*[vc(1);ifil(1)] + Bd*[vi(1);iload(1)];
        xb = Ad*[vc(2);ifil(2)] + Bd*[vi(2);iload(2)];
        cost = sum((vref-[xa(1);xb(1)]).^2)/(max(sum(vref.^2),1));
    else
        ifp = ifil + (ts/lf)*(vi-vg-rf*ifil);
        pp = 1.5*(vg(1)*ifp(1)+vg(2)*ifp(2));
        qp = 1.5*(vg(2)*ifp(1)-vg(1)*ifp(2));
        cost = ((pref-pp)/pbase)^2 + ((qref-qp)/qbase)^2;
    end
    cost = cost + lambda*nsw;
    if cost < bestCost
        bestCost = cost;
        best = k;
    end
end
s = states(best,:);
% SPS PWM Generator confirms complementary device pairs (1,2), (3,4),
% (5,6), corresponding to phases A, B and C.
g = [s(1);1-s(1);s(2);1-s(2);s(3);1-s(3)];
block.OutputPort(1).Data = [g;bestCost;best];
end

function update(block)
y = block.OutputPort(1).Data;
block.Dwork(1).Data = [y(1);y(3);y(5)];
end

function ab = clarke(abc)
ab = (2/3)*[1 -0.5 -0.5;0 sqrt(3)/2 -sqrt(3)/2]*abc(:);
end
