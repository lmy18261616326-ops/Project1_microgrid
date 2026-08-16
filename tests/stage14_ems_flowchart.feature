# --- front-matter:toml ---
model = "manul_model.slx"
component = "manul_model/EMS_flowchart"
[inputs]
Ppv = "Ppv"
Pac = "Pac"
Pdc = "Pdc"
SOC = "SOC"
GridConnected = "GridConnected"
[outputs]
Ibat_ref = "Ibat_ref"
offMPPT = "offMPPT"
loadShed = "loadShed"
battCmd = "battCmd"
# --- end front-matter ---

Feature: Stage 14 EMS flowchart truth table
  Verify grid-connected, islanded, and boundary decisions.

Scenario: Grid low SOC charges battery
  Given inputs
    * Ppv = const(1200000)
    * Pac = const(500000)
    * Pdc = const(600000)
    * SOC = const(40)
    * GridConnected = const(1)
  When simulate for 0.3s in Normal mode
  Then outputs
    * GridLowSocCurrent: Ibat_ref == [-1041.68 .. -1041.65] when t > 0.2s
    * GridLowSocMppt: offMPPT == 0 when t > 0.2s
    * GridLowSocShed: loadShed == 0 when t > 0.2s
    * GridLowSocCmd: battCmd == -1 when t > 0.2s

Scenario: Grid sufficient SOC and power deficit discharges battery
  Given inputs
    * Ppv = const(800000)
    * Pac = const(500000)
    * Pdc = const(600000)
    * SOC = const(60)
    * GridConnected = const(1)
  When simulate for 0.3s in Normal mode
  Then outputs
    * GridDeficitCurrent: Ibat_ref == [1041.65 .. 1041.68] when t > 0.2s
    * GridDeficitMppt: offMPPT == 0 when t > 0.2s
    * GridDeficitShed: loadShed == 0 when t > 0.2s
    * GridDeficitCmd: battCmd == 1 when t > 0.2s

Scenario: Grid sufficient SOC and power surplus stops battery
  Given inputs
    * Ppv = const(1200000)
    * Pac = const(500000)
    * Pdc = const(600000)
    * SOC = const(60)
    * GridConnected = const(1)
  When simulate for 0.3s in Normal mode
  Then outputs
    * GridSurplusCurrent: Ibat_ref == 0 when t > 0.2s
    * GridSurplusMppt: offMPPT == 0 when t > 0.2s
    * GridSurplusShed: loadShed == 0 when t > 0.2s
    * GridSurplusCmd: battCmd == 0 when t > 0.2s

Scenario: Island deficit with usable SOC discharges battery
  Given inputs
    * Ppv = const(800000)
    * Pac = const(500000)
    * Pdc = const(600000)
    * SOC = const(50)
    * GridConnected = const(0)
  When simulate for 0.3s in Normal mode
  Then outputs
    * IslandDeficitCurrent: Ibat_ref == [833.32 .. 833.35] when t > 0.2s
    * IslandDeficitMppt: offMPPT == 0 when t > 0.2s
    * IslandDeficitShed: loadShed == 0 when t > 0.2s
    * IslandDeficitCmd: battCmd == 1 when t > 0.2s

Scenario: Island deficit at minimum SOC sheds load
  Given inputs
    * Ppv = const(800000)
    * Pac = const(500000)
    * Pdc = const(600000)
    * SOC = const(10)
    * GridConnected = const(0)
  When simulate for 0.3s in Normal mode
  Then outputs
    * IslandLowSocCurrent: Ibat_ref == 0 when t > 0.2s
    * IslandLowSocMppt: offMPPT == 0 when t > 0.2s
    * IslandLowSocShed: loadShed == 1 when t > 0.2s
    * IslandLowSocCmd: battCmd == 0 when t > 0.2s

Scenario: Island surplus below maximum SOC charges battery
  Given inputs
    * Ppv = const(1200000)
    * Pac = const(500000)
    * Pdc = const(600000)
    * SOC = const(50)
    * GridConnected = const(0)
  When simulate for 0.3s in Normal mode
  Then outputs
    * IslandSurplusCurrent: Ibat_ref == [-833.35 .. -833.32] when t > 0.2s
    * IslandSurplusMppt: offMPPT == 0 when t > 0.2s
    * IslandSurplusShed: loadShed == 0 when t > 0.2s
    * IslandSurplusCmd: battCmd == -1 when t > 0.2s

Scenario: Island surplus at maximum SOC curtails PV
  Given inputs
    * Ppv = const(1200000)
    * Pac = const(500000)
    * Pdc = const(600000)
    * SOC = const(90)
    * GridConnected = const(0)
  When simulate for 0.3s in Normal mode
  Then outputs
    * IslandFullCurrent: Ibat_ref == 0 when t > 0.2s
    * IslandFullMppt: offMPPT == 1 when t > 0.2s
    * IslandFullShed: loadShed == 0 when t > 0.2s
    * IslandFullCmd: battCmd == 0 when t > 0.2s

Scenario: Grid midpoint SOC belongs to upper branch
  Given inputs
    * Ppv = const(800000)
    * Pac = const(500000)
    * Pdc = const(600000)
    * SOC = const(50)
    * GridConnected = const(1)
  When simulate for 0.3s in Normal mode
  Then outputs
    * GridMidCurrent: Ibat_ref == [833.32 .. 833.35] when t > 0.2s
    * GridMidMppt: offMPPT == 0 when t > 0.2s
    * GridMidShed: loadShed == 0 when t > 0.2s
    * GridMidCmd: battCmd == 1 when t > 0.2s

Scenario: Island zero power balance belongs to surplus branch
  Given inputs
    * Ppv = const(1100000)
    * Pac = const(500000)
    * Pdc = const(600000)
    * SOC = const(50)
    * GridConnected = const(0)
  When simulate for 0.3s in Normal mode
  Then outputs
    * IslandZeroCurrent: Ibat_ref == [-833.35 .. -833.32] when t > 0.2s
    * IslandZeroMppt: offMPPT == 0 when t > 0.2s
    * IslandZeroShed: loadShed == 0 when t > 0.2s
    * IslandZeroCmd: battCmd == -1 when t > 0.2s

Scenario: Grid signal at threshold belongs to island branch
  Given inputs
    * Ppv = const(800000)
    * Pac = const(500000)
    * Pdc = const(600000)
    * SOC = const(10)
    * GridConnected = const(0.5)
  When simulate for 0.3s in Normal mode
  Then outputs
    * GridThresholdCurrent: Ibat_ref == 0 when t > 0.2s
    * GridThresholdMppt: offMPPT == 0 when t > 0.2s
    * GridThresholdShed: loadShed == 1 when t > 0.2s
    * GridThresholdCmd: battCmd == 0 when t > 0.2s

Scenario: Island minimum SOC with surplus charges at maximum current
  Given inputs
    * Ppv = const(1200000)
    * Pac = const(500000)
    * Pdc = const(600000)
    * SOC = const(10)
    * GridConnected = const(0)
  When simulate for 0.3s in Normal mode
  Then outputs
    * IslandMinChargeCurrent: Ibat_ref == [-1666.68 .. -1666.65] when t > 0.2s
    * IslandMinChargeMppt: offMPPT == 0 when t > 0.2s
    * IslandMinChargeShed: loadShed == 0 when t > 0.2s
    * IslandMinChargeCmd: battCmd == -1 when t > 0.2s

