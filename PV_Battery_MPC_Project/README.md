# PV-Battery MPC 文献复现项目

本项目对应论文 *Predictive Control of PV/Battery System under Load and
Environmental Uncertainty*，实现独立直流微网中的 9.5 kW 光伏、20 Ah
锂电池、两路 Boost 平均值变换器、P&O、ARIMA 一步预测、四模式功率管理、
非线性 MPC 和 PI 对照。

模型的算法与对象尽量使用常规 Simulink 模块。主模型中没有 MATLAB Function
块；非线性优化由 Model Predictive Control Toolbox 的标准
`Nonlinear MPC Controller` 模块执行。

## 快速开始

1. 双击打开 `PV_Battery_MPC_Project.prj`。
2. 在 MATLAB 命令行运行：

   ```matlab
   START_HERE
   ```

3. 运行论文第 5.5 节 MPC 工况：

   ```matlab
   result = run_paper_scenario("comparison", "MPC");
   plot_paper_scenario(result);
   ```

4. 运行 PI 对照：

   ```matlab
   resultPI = run_paper_scenario("comparison", "PI");
   ```

5. 生成验证结果和图片：

   ```matlab
   report = run_validation;
   ```

## 场景名称

| 场景 | 说明 |
|---|---|
| `comparison` | 论文第 5.5 节的 3 s MPC/PI 对比 |
| `mode1` | PV 由不足转为富余，电池由放电转充电 |
| `mode2` | PV 由富余转为不足，电池由充电转放电 |
| `mode3` | 16 s 完整限功率模式 |
| `mode3_quick` | 1 s 的 Mode III 快速验证 |
| `mode4` | 光伏功率降为零，电池承担负载 |
| `realworld` | 300 s 合成日内环境/负载曲线，用于替换作者未公开的数据 |

## 变量与信号

状态向量：

```text
x = [Vpv, iLpv, Vb, iLb, SoC, Vdc]
```

辅助向量：

```text
aux = [Ipv, Ib, Ebat, Ppv, Pbat, Pload]
```

模式向量：

```text
mode_bus = [alpha, mode_id]
```

其中 `alpha=1` 表示 Mode III 限功率目标，`mode_id=1..4` 对应论文四种
运行模式，`mode_id=5` 是论文未展开的负载切除请求状态。

## 目录

- `models/PV_Battery_MPC.slx`：主模型；
- `scripts/pvbatt_initialize.m`：参数、查表、场景和 MPC 初始化；
- `scripts/run_paper_scenario.m`：单场景仿真入口；
- `scripts/run_validation.m`：主要验证与结果图；
- `docs/reproduction_notes.md`：论文参数映射、缺失量与复现边界；
- `output/pdf/PV_Battery_MPC_Model_Audit_and_System_Report.pdf`：逐系统说明、参数依据和多轮故障检查报告；
- `results/`：验证图片和压缩结果数据。

完整复现边界与假设请先阅读
[`docs/reproduction_notes.md`](docs/reproduction_notes.md)。
