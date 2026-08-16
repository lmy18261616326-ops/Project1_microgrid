function p = phase2_parameters()
%PHASE2_PARAMETERS Discrete P&O MPPT parameters for phase 2.

p.mppt.sampleTime = 2e-3;
p.mppt.initialDuty = 0.5656125;
p.mppt.dutyStep = 5e-4;
p.mppt.minimumDuty = 0.05;
p.mppt.maximumDuty = 0.90;
p.mppt.powerDropThreshold = 0.1;
p.mppt.initialPower = 9591.75;
end
