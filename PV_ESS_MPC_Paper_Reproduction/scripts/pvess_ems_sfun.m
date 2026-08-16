function pvess_ems_sfun(block)
%PVESS_EMS_SFUN Executable form of the paper's Fig. 9 EMS flowchart.
setup(block);
end

function setup(block)
block.NumDialogPrms = 2; % SOCmin, SOCmax
block.NumInputPorts = 1;
block.InputPort(1).Dimensions = 5; % grid, Ppv, Pac, Pdc, SOC
block.InputPort(1).DirectFeedthrough = true;
block.NumOutputPorts = 1;
block.OutputPort(1).Dimensions = 3; % battCmd, offMPPT, loadShed
block.SampleTimes = [1e-3 0];
block.RegBlockMethod('Outputs',@outputs);
end

function outputs(block)
u = block.InputPort(1).Data;
gridTied = u(1) > 0.5;
pnet = u(2) - u(3) - u(4);
soc = u(5);
socMin = block.DialogPrm(1).Data;
socMax = block.DialogPrm(2).Data;
battCmd = 0; offMppt = 0; loadShed = 0;
if gridTied
    % Grid-tied operation permits SOC-oriented charge/discharge.
    if soc < 0.5*(socMin + socMax)
        battCmd = -1;
    elseif pnet < 0 && soc > socMin
        battCmd = 1;
    end
else
    if pnet < 0
        if soc > socMin
            battCmd = 1;
        else
            loadShed = 1;
        end
    else
        if soc < socMax
            battCmd = -1;
        else
            offMppt = 1;
        end
    end
end
block.OutputPort(1).Data = [battCmd;offMppt;loadShed];
end
