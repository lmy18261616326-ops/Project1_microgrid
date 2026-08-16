function [gates,stateIndex,Jmin] = pvess_mpvc_block( ...
    Vdc,Vc_ab,If_ab,IL_ab,Vref_ab,Ad,Bd,lambda_sw,previousState)
%PVESS_MPVC_BLOCK Code for the island-mode MATLAB Function block.
% Copy this whole function into a MATLAB Function block, or call it from
% that block. All vector inputs are real 2-by-1 alpha-beta vectors.
%
% Gate order for the SPS Universal Bridge is:
% [A_upper; A_lower; B_upper; B_lower; C_upper; C_lower].

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

for k = 1:8
    s = states(k,:);

    % Two-level bridge voltage vector in stationary alpha-beta axes.
    Vi_alpha = (2/3)*Vdc*(s(1)-0.5*s(2)-0.5*s(3));
    Vi_beta  = (sqrt(3)/3)*Vdc*(s(2)-s(3));

    % Paper Eq. (11), independently applied to alpha and beta axes.
    xa = Ad*[Vc_ab(1);If_ab(1)] + Bd*[Vi_alpha;IL_ab(1)];
    xb = Ad*[Vc_ab(2);If_ab(2)] + Bd*[Vi_beta; IL_ab(2)];

    voltageScale = max(Vref_ab(1)^2 + Vref_ab(2)^2,1);
    Jv = ((Vref_ab(1)-xa(1))^2 + (Vref_ab(2)-xb(1))^2) ...
        / voltageScale;
    Jsw = lambda_sw*sum(abs(s-previousState));
    J = Jv + Jsw;

    if J < Jmin
        Jmin = J;
        stateIndex = k;
    end
end

s = states(stateIndex,:);
gates = [s(1);1-s(1);s(2);1-s(2);s(3);1-s(3)];
end

