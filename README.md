# OpenClaw-Sifu

<p align="center">
  <img src="frontend/src/assets/images/logo.png" alt="OpenClaw-Sifu Logo" width="300" />
</p>

[English](#english) | [中文](#中文)

<a name="english"></a>
## English

OpenClaw-Sifu is the standalone graphical installer and uninstaller for [OpenClaw](https://github.com/simonfqy/openclaw). Built with Wails, it provides a seamless, local installation experience by automating prerequisite checks, dependency acquisition, and environment initialization. For the release version, please click here to [download](https://github.com/blueSLota/openclaw-sifu/releases).

### Features
- **One-Click Setup:** Fully automated GUI for installing and removing OpenClaw.
- **Cross-Platform:** Smart detection of Windows, macOS, and Linux to execute native installation scripts.
- **Visual Feedback:** Real-time tracking of setup phases (environment checks, npm installation, component initialization).
- **System Integration:** Manages administrative permissions for background services and scheduled tasks.

### Tech Stack
- **Framework:** Wails v2
- **Backend:** Go (System calls & execution mapping)
- **Frontend:** React + Vite + TypeScript + TailwindCSS

### Development
Prerequisites: Go and Node.js must be installed.

```bash
# Start Wails dev server (hot reload for both frontend and Go)
wails dev
```

### Build
```bash
# Build a standalone executable for your current platform
wails build
```

---

<a name="中文"></a>
## 中文

OpenClaw-Sifu 是 [OpenClaw](https://github.com/simonfqy/openclaw) 的独立图形化安装与卸载工具。基于 Wails 构建，它通过自动化环境检查、依赖获取和环境初始化，提供无缝的本地安装体验。发布版请点击这里进行[下载](https://github.com/blueSLota/openclaw-sifu/releases)。如果下载速度缓慢，也可以前往[官网](https://openclaw-sifu.com/)通过国内CDN下载

### 核心特性
- **一键安装/卸载：** 全自动化的图形界面，一键完成 OpenClaw 的安装与清理。
- **跨平台兼容：** 智能识别 Windows、macOS 和 Linux，并相应执行原生的各平台安装脚本。
- **可视化反馈：** 实时追踪安装步骤进度（如环境检查、npm 安装、组件初始化等）。
- **系统集成：** 自动处理并申请后台服务与计划任务所需的管理员权限。

### 技术架构
- **桌面框架：** Wails v2
- **后端：** Go（负责系统调用和跨平台代码执行路由）
- **前端：** React + Vite + TypeScript + TailwindCSS

### 本地开发
环境要求：必须预先安装 Go 和 Node.js。

```bash
# 启动 Wails 开发服务器（支持前端和 Go 的热重载）
wails dev
```

### 生产构建
```bash
# 构建当前平台的可执行文件
wails build
```
