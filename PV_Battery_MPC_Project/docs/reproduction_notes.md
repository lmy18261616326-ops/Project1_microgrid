# 论文复现映射与边界

目标文献：

> S. Batiyah, R. Sharma, S. Abdelwahed, W. Alhosaini, and O. Aldosari,
> “Predictive Control of PV/Battery System under Load and Environmental
> Uncertainty,” *Energies*, vol. 15, no. 11, 4100, 2022.

DOI: <https://doi.org/10.3390/en15114100>

## 1. 已按原文落实的内容

| 原文内容 | 项目实现 |
|---|---|
| 六状态平均值模型，式 (12) | `PV_Battery_Averaged_Plant`，完全由常规 Simulink 运算、查表和积分模块构成 |
| PV 与电池两路 Boost 占空比 | `u = [dpv; db]`，统一限制在 `[0.02, 0.95]`；MPC 内部约束仍为 `[0,1]` |
| 9.5 kW PV、300 V/20 Ah 电池、600 V DC 母线 | 参数脚本和模型初值 |
| Cpv/Cb = 300 uF，Lpv/Lb = 10 mH，Cdc = 1500 uF | 参数脚本与各导数支路 |
| 仿真步长 10 us、控制周期 10 ms、预测时域 1 | 模型配置与 Nonlinear MPC Controller |
| 权重 F/P/R/Q = 0.75/0.15/0.10/0.15 | MPC 对象与在线模式权重 |
| P&O 产生 PV 电压参考 | `P_and_O_MPPT`，仅用标准离散模块 |
| Mode I-IV 与 SoC 20%/90% 阈值 | `Power_Management_Logic` |
| Mode III：电池电流归零、PV 限功率 | 在线 MPC 权重 + 标准模块约束执行层 |
| ARIMA 预测接口 | `ARIMA_OneStep_Forecast` |
| 第 5.5 节 500→1000 W/m²、35→25°C、7.2→14.4 kW 工况 | `comparison` 场景 |

## 2. 原文未公开而必须补足的参数

论文没有给出下列复现必需量，项目没有把推定值伪装成原文值：

1. 9.5 kW 阵列的串并联结构。项目采用 9 串 × 5 并，共 45 块
   213.15 W 模块，额定功率 9.59175 kW。
2. 单二极管模型的 `Rs`、`Rsh`、理想因子和温度系数。项目由原文
   `Isc/Voc/Imp/Vmp` 拟合 STC 参数，并把温度系数保留为显式假设。
3. 电池式 (7)-(8) 的 `E0/K/A/B/Rbat`。项目采用可替换 OCV-SoC 查表和
   0.08 ohm 内阻；这是当前结果与论文不能逐点一致的主要原因。
4. ARIMA 的阶次、系数和训练历史。默认使用稳健的
   ARIMA(0,1,0) 一步持久性预测；模型内保留 ARIMA(1,1,0) 系数支路，
   可在 `pvbatt_parameters.m` 中标定。
5. Mode I-IV 的完整负载/温度/辐照序列。项目按论文文字描述构造阶跃，
   `comparison` 场景则严格采用第 5.5 节公开的时间和值。

## 3. 方程处理说明

原文式 (12) 的 SoC 行写成 `(Ebat - x4)`，但按式 (6)、式 (9) 以及量纲，
应为电池端电压状态 `x3`。项目采用：

```text
dSoC/dt = -(Ebat - vb) / (3600 Q Rbat)
```

主对象保留全部六个连续状态并按 10 us 积分。MPC 的 10 ms 内部预测器把
电池端电容的快状态按准稳态解析处理，避免显式预测因多时间尺度而数值发散；
这只影响控制器内部预测，不改变被控对象。

## 4. 模型中自定义代码的范围

模型不含 MATLAB Function 块。正常 Simulink 模块负责：

- 六状态对象；
- PV、电池查表；
- P&O；
- ARIMA 一步预测；
- 功率管理模式逻辑；
- Mode III 锁存与约束执行；
- PI 对照和全部日志。

外部 MATLAB 函数仅用于：

- 离线生成单二极管 PV 查表；
- 标准 Nonlinear MPC Controller 调用非线性一步预测模型；
- 场景、参数、仿真与作图。

## 5. 已验证结果

- 第 5.5 节 3 s 对比工况：
  - MPC 平均 VRI 约 0.122%，最大 VRI 约 3.95%；
  - PI 平均 VRI 约 0.110%，最大 VRI 约 5.99%。
- 加速 Mode III：
  - 多轮检查发现并修复了 Mode III 进入瞬间一个控制周期的模式编码空档；
  - 修复后约 0.306 s 进入并锁存限功率模式，切换窗口仅出现 Mode I 和 Mode III；
  - 1 s 时 PV 约 5.92 kW、电池功率约 0.156 kW，继续向零收敛；
  - 切换暂态呈阻尼收敛。

论文报告的精确 VRI 不能在缺少电池参数、ARIMA 系数和作者仿真文件时保证
逐点一致。本项目的目标是给出结构、方程、创新逻辑和工况均可检查、可替换、
可继续标定的科研参考模型。
