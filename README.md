### Hi, I'm `XiAzir`

关注方向：**AI Agent 工程化与运行时编排** / **Hybrid RAG 检索链路** / **系统小工具与 FPGA 硬件交付**

---

### Selected Projects

| Project | Stack | What it does |
| --- | --- | --- |
| **[AzAgent](https://github.com/XiAzir/AzAgent)** | TypeScript / Python / DSH | 跨平台自主编程助手：基于 DeepSeek Harness 官方 SuperAgent 思路，打通 Web、GitHub、飞书与 QQ 任务路由，实现外部 Watchdog、独立结果校验、可靠记忆与 000–008 架构规范 |
| **[campus-policy-agent](https://github.com/XiAzir/campus-policy-agent)** | Python / FastAPI / LangGraph | 校园政策 Hybrid RAG 问答平台：BM25（SQLite FTS5）+ 1024 维向量 RRF 混合检索、原文行号溯源、版本管理，配套 81 项后端 pytest 与前后端 E2E 测试，部署于 1C1G 服务器 |
| **[dsh-payload-guard](https://github.com/XiAzir/dsh-payload-guard)** | JavaScript / Node.js / DSH | 多模态长上下文 413 自愈插件：基于 9.79 MB / 10.04 MB 实测边界实现被动渐进图片卸载（每轮 50%），权衡 Prompt Cache 失效代价并以回合预算防死循环 |
| **[FPGA-multimedia-player](https://github.com/XiAzir/FPGA-multimedia-player)** | Verilog / Python / PowerShell | 基于安路 EG4S20BG256 的离线多媒体展示终端最终位流 `c14116_seed2.bit`（Setup `+0.804 ns` / Hold `+0.027 ns`，TNS=0），配套双仓架构 [FPGA-ANLU](https://github.com/XiAzir/FPGA-ANLU) 与 [FPGA-ANLU-docs](https://github.com/XiAzir/FPGA-ANLU-docs) |

> 其他原生系统实验：**[Kumokiri](https://github.com/XiAzir/Kumokiri)** — 基于 Rust + DXGI / Direct3D 11 的 Windows 轻量截图工具（支持 HDR→SDR Reinhard 色调映射与全局热键）。

---

### Engineering Habits

- **规范驱动上下文管理**：习惯用薄 `AGENTS.md`、按需路由表和 `SKILL.md` 约束长周期项目，保持代码与文档边界清晰。
- **先测后交，用证据闭环**：重视自动化回归测试、实测边界数据与阶段状态快照。
- **关注真实工程权衡**：围绕可用性、缓存成本、资源上限与失败降级做清晰取舍。
