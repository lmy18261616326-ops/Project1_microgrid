function [gates,stateIndex,Jmin] = pvess_mppc_block( ...
    Vdc,Vg_ab,P,Q,Pref,Qref,Ts,Lf,Rf,omega,Pbase,Qbase, ...
    lambda_sw,previousState)
%PVESS_MPPC_BLOCK Code for the grid-connected MATLAB Function block.
% Copy this whole function into a MATLAB Function block, or call it from
% that block. The equations follow the paper's P/Q predictive model.
%
% IMPORTANT SIGN CONVENTION:
% P,Q and the measured current must use one consistent direction. This
% implementation follows the paper equations with positive current/power
% from grid toward converter. If the Current Measurement arrow points from
% converter toward grid, multiply that current by -1 before P/Q calculation.

states = [0 0 0; ...
          1 0 0; ...
          1 1 0; ...
          0 1 0; ...
          0 1 1; ...
          0 0 1; ...
          1 0 1; ...
          1 1 1];

Jmin = inf;
stateIndex = 1;
previousState = previousState(:).';
Vg2 = Vg_ab(1)^2 + Vg_ab(2)^2;

for k = 1:8
    s = states(k,:);
    Vi_alpha = (2/3)*Vdc*(s(1)-0.5*s(2)-0.5*s(3));
    Vi_beta  = (sqrt(3)/3)*Vdc*(s(2)-s(3));

    % real/imaginary parts of Vg*conj(Vi).
    realTerm = Vg_ab(1)*Vi_alpha + Vg_ab(2)*Vi_beta;
    imagTerm = Vg_ab(2)*Vi_alpha - Vg_ab(1)*Vi_beta;

    % Paper active/reactive power one-step prediction.
    Ppred = P + Ts*(-Rf/Lf*P - omega*Q ...
        + 3/(2*Lf)*(Vg2-realTerm));
    Qpred = Q + Ts*( omega*P - Rf/Lf*Q ...
        - 3/(2*Lf)*imagTerm);

    Jpq = ((Pref-Ppred)/Pbase)^2 + ((Qref-Qpred)/Qbase)^2;
    Jsw = lambda_sw*sum(abs(s-previousState));
    J = Jpq + Jsw;

    if J < Jmin
        Jmin = J;
        stateIndex = k;
    end
end

s = states(stateIndex,:);
gates = [s(1);1-s(1);s(2);1-s(2);s(3);1-s(3)];
end

