import { useEffect, useState } from 'react';
import './App.css';
import {
  GetBootstrapPayload,
  LaunchOpenClaw,
  RunNativeInstaller,
  RunNativeUninstaller,
  RunPostInstallActions,
} from '../wailsjs/go/main/App';
import {
  EventsOn,
  Quit,
  WindowIsMaximised,
  WindowMinimise,
  WindowToggleMaximise,
} from '../wailsjs/runtime/runtime';
import { Minus, Square, Copy, X } from 'lucide-react';

type Environment = {
  hostname: string;
  platform: string;
  architecture: string;
};

type BootstrapPayload = {
  environment: Environment;
};

type InstallerStepUpdate = {
  step: string;
  status: string;
  message: string;
};

type PostInstallActionResult = {
  success: boolean;
  message: string;
  error?: string;
  cancelled: boolean;
};

export default function App() {
  const [payload, setPayload] = useState<BootstrapPayload | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');
  const [isMaximised, setIsMaximised] = useState(false);
  const [appPhase, setAppPhase] = useState<'home' | 'installing' | 'uninstalling' | 'failed' | 'success'>('home');
  const [currentAction, setCurrentAction] = useState<'install' | 'uninstall'>('install');
  const [installerSteps, setInstallerSteps] = useState<InstallerStepUpdate[]>([]);
  const [enableSystemIntegration, setEnableSystemIntegration] = useState(true);
  const [postInstallBusy, setPostInstallBusy] = useState(false);
  const [postInstallMessage, setPostInstallMessage] = useState<{
    tone: 'info' | 'error';
    text: string;
  } | null>(null);
  const supportsPostInstallElevation = payload?.environment.platform === 'windows';

  async function syncWindowState() {
    try {
      setIsMaximised(await WindowIsMaximised());
    } catch {
      setIsMaximised(false);
    }
  }

  async function loadPayload() {
    setLoading(true);
    setError('');

    try {
      const localPayload = await GetBootstrapPayload();
      setPayload(localPayload as unknown as BootstrapPayload);
    } catch (err) {
      setError(err instanceof Error ? err.message : '无法获取本地运行时上下文。');
    } finally {
      setLoading(false);
    }
  }

  useEffect(() => {
    void loadPayload();
    void syncWindowState();
    const unsubscribeInstaller = EventsOn('installer:step', (update: InstallerStepUpdate) => {
      setInstallerSteps((prev) => [...prev, update]);
    });
    const handleResize = () => void syncWindowState();
    window.addEventListener('resize', handleResize);
    return () => {
      window.removeEventListener('resize', handleResize);
      unsubscribeInstaller();
    };
  }, []);

  async function toggleMaximise() {
    WindowToggleMaximise();
    window.setTimeout(() => void syncWindowState(), 120);
  }

  async function startLocalInstall() {
    if (!payload) {
      setError('无法获取当前环境信息。');
      return;
    }

    setCurrentAction('install');
    setAppPhase('installing');
    setError('');
    setInstallerSteps([]);
    setEnableSystemIntegration(true);
    setPostInstallMessage(null);

    try {
      const result = await RunNativeInstaller({
        tag: 'latest',
        installMethod: 'npm',
        noOnboard: true,
        noGitUpdate: false,
        dryRun: false,
        useCnMirrors: false,
        npmRegistry: '',
        installBaseUrl: '',
        repoUrl: '',
        gitDir: '',
      });

      if (result.success) {
        setAppPhase('success');
      } else {
        setError(result.error || '安装失败');
        setAppPhase('failed');
      }
    } catch (err) {
      setError(err instanceof Error ? err.message : '启动安装程序失败。');
      setAppPhase('failed');
    }
  }

  async function startLocalUninstall() {
    if (!payload) {
      setError('无法获取当前环境信息。');
      return;
    }

    setCurrentAction('uninstall');
    setAppPhase('uninstalling');
    setError('');
    setInstallerSteps([]);
    setPostInstallMessage(null);

    try {
      const result = await RunNativeUninstaller();
      if (result.success) {
        setAppPhase('success');
      } else {
        setError(result.error || '卸载失败');
        setAppPhase('failed');
      }
    } catch (err) {
      setError(err instanceof Error ? err.message : '启动卸载程序失败。');
      setAppPhase('failed');
    }
  }

  async function openOpenClawAndQuit() {
    await LaunchOpenClaw();
    window.setTimeout(() => Quit(), 2000);
  }

  async function finishInstall() {
    if (postInstallBusy) {
      return;
    }

    setPostInstallBusy(true);
    setPostInstallMessage(null);

    try {
      if (!supportsPostInstallElevation || !enableSystemIntegration) {
        await openOpenClawAndQuit();
        return;
      }

      const result = await RunPostInstallActions() as unknown as PostInstallActionResult;
      if (!result.success) {
        setPostInstallMessage({
          tone: result.cancelled ? 'info' : 'error',
          text: result.message || result.error || '管理员授权未完成，系统集成未启用。',
        });
        return;
      }

      await openOpenClawAndQuit();
    } catch (err) {
      setPostInstallMessage({
        tone: 'error',
        text: err instanceof Error ? err.message : '执行安装后的授权步骤失败。',
      });
    } finally {
      setPostInstallBusy(false);
    }
  }

  if (loading) {
    return (
      <main className="flex flex-col h-screen w-screen bg-slate-50 text-slate-800 overflow-hidden">
        <WindowChrome isMaximised={isMaximised} onToggleMaximise={toggleMaximise} />
        <div className="flex-1 flex items-center justify-center p-8">
          <div className="flex flex-col items-center gap-4 text-slate-500">
            <div className="w-10 h-10 border-4 border-slate-300 border-t-orange-500 rounded-full animate-spin"></div>
            <p className="font-medium tracking-wider">正在准备安装环境...</p>
          </div>
        </div>
      </main>
    );
  }

  if (error || !payload) {
    return (
      <main className="flex flex-col h-screen w-screen bg-slate-50 text-slate-800 overflow-hidden">
        <WindowChrome isMaximised={isMaximised} onToggleMaximise={toggleMaximise} />
        <div className="flex-1 flex items-center justify-center p-8">
          <div className="bg-white border border-rose-200 shadow-xl shadow-rose-100 rounded-3xl p-8 max-w-md w-full text-center">
            <div className="w-16 h-16 bg-rose-100 text-rose-500 rounded-2xl flex items-center justify-center mx-auto mb-4">
              <X size={32} />
            </div>
            <h1 className="text-2xl font-bold text-slate-800 mb-2">启动失败</h1>
            <p className="text-slate-500 mb-8">{error || '未能从本地获取数据。'}</p>
            <button
              className="w-full bg-slate-800 hover:bg-slate-700 text-white font-medium py-3 px-6 rounded-xl transition-colors"
              onClick={() => void loadPayload()}
            >
              重试
            </button>
          </div>
        </div>
      </main>
    );
  }

  return (
    <main className="flex flex-col h-screen w-screen bg-gradient-to-br from-slate-50 to-slate-100 text-slate-800 overflow-hidden">
      <WindowChrome isMaximised={isMaximised} onToggleMaximise={toggleMaximise} />

      <div className="flex-1 flex flex-col p-6 lg:p-10 overflow-hidden">
        <div className="flex-1 flex flex-col overflow-hidden rounded-3xl bg-white border border-slate-200/80 shadow-xl shadow-slate-200/30 max-w-4xl w-full mx-auto">
          <div className="installer-viewport flex-1">

            {/* ===== SLIDE 1: Welcome / Home ===== */}
            <div className={`installer-slide ${appPhase === 'home' ? 'slide-active' : 'slide-left'}`}>
              <div className="flex flex-col items-center text-center px-8 max-w-xl">
                {/* Logo */}
                <div className="w-20 h-20 rounded-3xl bg-gradient-to-br from-orange-500 to-rose-500 flex items-center justify-center text-white text-3xl font-black tracking-tighter shadow-lg shadow-orange-500/30 mb-8 animate-fade-in-up">
                  OC
                </div>

                <h1 className="text-4xl font-black text-slate-900 tracking-tight leading-tight mb-4 animate-fade-in-up animate-delay-100">
                  准备安装{' '}
                  <span className="text-transparent bg-clip-text bg-gradient-to-r from-orange-500 to-rose-500">
                    OpenClaw
                  </span>
                </h1>

                <p className="text-slate-500 text-lg leading-relaxed font-medium mb-8 animate-fade-in-up animate-delay-200">
                  您的个人专属生产力智能体。<br />
                  系统已检测到您的环境，点击下方即可开始全自动安装。
                </p>

                {/* Environment badges */}
                <div className="flex flex-wrap justify-center gap-2 mb-10 animate-fade-in-up animate-delay-300">
                  <span className="px-3 py-1.5 bg-slate-100 border border-slate-200 rounded-full text-xs font-medium text-slate-600">
                    {payload.environment.platform} · {payload.environment.architecture}
                  </span>
                  <span className="px-3 py-1.5 bg-slate-100 border border-slate-200 rounded-full text-xs font-medium text-slate-600">
                    {payload.environment.hostname}
                  </span>
                </div>

                {/* CTA */}
                <div className="flex flex-col sm:flex-row gap-3 animate-fade-in-up animate-delay-400">
                  <button
                    className="px-10 py-4 rounded-2xl font-bold text-lg text-white bg-gradient-to-r from-orange-500 to-orange-600 hover:from-orange-600 hover:to-orange-700 shadow-xl shadow-orange-500/25 transition-all hover:-translate-y-1 hover:shadow-2xl hover:shadow-orange-500/30 active:translate-y-0"
                    onClick={() => void startLocalInstall()}
                  >
                    开始一键安装
                  </button>
                  <button
                    className="px-10 py-4 rounded-2xl font-bold text-lg text-rose-700 bg-white border border-rose-200 hover:bg-rose-50 shadow-xl shadow-rose-100/40 transition-all hover:-translate-y-1 active:translate-y-0"
                    onClick={() => void startLocalUninstall()}
                  >
                    一键卸载
                  </button>
                </div>

                <p className="text-xs text-slate-400 mt-4 animate-fade-in-up animate-delay-500">
                  安装或卸载过程完全自动，通常需要 1-3 分钟
                </p>
              </div>
            </div>

            {/* ===== SLIDE 2: Installing (Phase cards) ===== */}
            <div className={`installer-slide ${appPhase === 'installing' || appPhase === 'uninstalling' ? 'slide-active' : appPhase === 'home' ? 'slide-right' : 'slide-left'}`}>
              <div className="flex flex-col items-center w-full max-w-lg px-8">
                <div className="w-16 h-16 rounded-2xl bg-gradient-to-br from-orange-500 to-rose-500 flex items-center justify-center text-white text-xl font-black shadow-lg shadow-orange-500/20 mb-6 shimmer-active">
                  OC
                </div>

                <h2 className="text-2xl font-bold text-slate-800 mb-2">
                  {currentAction === 'install' ? '正在安装 OpenClaw' : '正在卸载 OpenClaw'}
                </h2>
                <p className="text-sm text-slate-500 mb-8">
                  {currentAction === 'install' ? '请稍候，安装程序正在配置您的环境' : '请稍候，卸载程序正在清理您的环境'}
                </p>

                {/* Phase cards */}
                <div className="w-full space-y-3">
                  {(() => {
                    const phases = currentAction === 'install'
                      ? [
                        { id: 'node', label: '检测运行环境', icon: '⬡', desc: '检查 Node.js 版本' },
                        { id: 'npm-install', label: '安装核心组件', icon: '⚙', desc: '通过 npm 安装 OpenClaw' },
                        { id: 'path', label: '配置系统路径', icon: '⛓', desc: '确保命令行可用' },
                        { id: 'gateway', label: '初始化服务', icon: '◈', desc: '配置后台服务' },
                        { id: 'done', label: '完成安装', icon: '✦', desc: '验证安装结果' },
                      ]
                      : [
                        { id: 'detect', label: '检查安装状态', icon: '⬡', desc: '确认 OpenClaw 已安装' },
                        { id: 'uninstall', label: '移除服务与数据', icon: '⌦', desc: '执行官方卸载流程' },
                        { id: 'gateway', label: '移除计划任务', icon: '◈', desc: '清理 OpenClaw Gateway 任务' },
                        { id: 'cli-remove', label: '清理 CLI', icon: '⌫', desc: '移除全局 OpenClaw 命令' },
                        { id: 'done', label: '完成卸载', icon: '✦', desc: '验证卸载结果' },
                      ];

                    const currentStep = installerSteps.length > 0 ? installerSteps[installerSteps.length - 1] : null;

                    // Determine which phase is active
                    const phaseOrder = phases.map(p => p.id);
                    const activePhaseIdx = currentStep
                      ? (() => {
                        const stepId = currentStep.step;
                        // Map sub-steps to their parent phase
                        const mapping: Record<string, string> = currentAction === 'install'
                          ? {
                            'init': 'node', 'node': 'node', 'node-install': 'node',
                            'detect': 'npm-install', 'npm-install': 'npm-install', 'git-install': 'npm-install', 'git-check': 'npm-install',
                            'path': 'path',
                            'gateway': 'gateway', 'doctor': 'gateway', 'onboard': 'gateway', 'setup': 'gateway',
                            'done': 'done',
                          }
                          : {
                            'init': 'detect',
                            'detect': 'detect',
                            'uninstall': 'uninstall',
                            'gateway': 'gateway',
                            'cli-remove': 'cli-remove',
                            'done': 'done',
                          };
                        const mapped = mapping[stepId] || stepId;
                        const idx = phaseOrder.indexOf(mapped);
                        return idx >= 0 ? idx : 0;
                      })()
                      : 0;

                    return phases.map((phase, idx) => {
                      const isDone = idx < activePhaseIdx || (idx === activePhaseIdx && currentStep?.status === 'ok' && phase.id === 'done');
                      const isActive = idx === activePhaseIdx && !isDone;
                      const isPending = idx > activePhaseIdx;
                      const isError = isActive && currentStep?.status === 'error';

                      return (
                        <div
                          key={phase.id}
                          className={`phase-card flex items-center gap-4 p-4 rounded-2xl border transition-all duration-500 ${isActive
                            ? 'bg-orange-50/80 border-orange-200 shadow-sm'
                            : isDone
                              ? 'bg-emerald-50/50 border-emerald-100'
                              : 'bg-slate-50/50 border-slate-100'
                            } ${isError ? '!bg-rose-50/80 !border-rose-200' : ''}`}
                        >
                          {/* Icon */}
                          <div className={`w-10 h-10 rounded-xl flex items-center justify-center text-lg shrink-0 transition-all duration-500 ${isDone
                            ? 'bg-emerald-500 text-white'
                            : isActive
                              ? isError
                                ? 'bg-rose-500 text-white'
                                : 'bg-orange-500 text-white'
                              : 'bg-slate-200 text-slate-400'
                            }`}>
                            {isDone ? '✓' : isActive ? (
                              isError ? '✗' : <span className="progress-spinner inline-block">⟳</span>
                            ) : phase.icon}
                          </div>

                          {/* Text */}
                          <div className="flex-1 min-w-0">
                            <div className={`text-sm font-bold transition-colors duration-300 ${isDone ? 'text-emerald-700' : isActive ? (isError ? 'text-rose-700' : 'text-slate-800') : 'text-slate-400'
                              }`}>
                              {phase.label}
                            </div>
                            <div className={`text-xs mt-0.5 transition-colors duration-300 truncate ${isDone ? 'text-emerald-600/70' : isActive ? (isError ? 'text-rose-500' : 'text-slate-500') : 'text-slate-300'
                              }`}>
                              {isActive && currentStep ? currentStep.message : isDone ? '已完成' : phase.desc}
                            </div>
                          </div>

                          {/* Status dot */}
                          {isActive && !isError && (
                            <div className="w-2 h-2 rounded-full bg-orange-500 animate-pulse shrink-0" />
                          )}
                        </div>
                      );
                    });
                  })()}
                </div>
              </div>
            </div>

            {/* ===== SLIDE 3: Success ===== */}
            <div className={`installer-slide ${appPhase === 'success' ? 'slide-active' : 'slide-right'}`}>
              <div className="flex flex-col items-center text-center px-8 max-w-md">
                <div className="w-20 h-20 rounded-3xl bg-gradient-to-br from-emerald-400 to-emerald-600 flex items-center justify-center text-white text-4xl shadow-lg shadow-emerald-500/30 mb-8 animate-sparkle">
                  ✓
                </div>

                <h2 className="text-3xl font-black text-slate-900 mb-3 animate-fade-in-up">
                  {currentAction === 'install' ? '安装完成！' : '卸载完成！'}
                </h2>

                {installerSteps.find(s => s.step === 'done' && s.status === 'ok') && (
                  <p className="text-sm text-emerald-600 font-medium bg-emerald-50 border border-emerald-100 rounded-full px-4 py-1.5 mb-4 animate-fade-in-up animate-delay-100">
                    {installerSteps.find(s => s.step === 'done' && s.status === 'ok')?.message}
                  </p>
                )}

                <p className="text-slate-500 text-base leading-relaxed font-medium mb-10 animate-fade-in-up animate-delay-200">
                  {currentAction === 'install' ? (
                    <>
                      OpenClaw 已成功安装到您的电脑中。<br />
                      现在打开 OpenClaw 开始使用吧！
                    </>
                  ) : (
                    <>
                      OpenClaw 已从这台电脑中移除。<br />
                      如需重新使用，可返回首页再次安装。
                    </>
                  )}
                </p>

                {currentAction === 'install' && supportsPostInstallElevation && (
                  <label className="w-full text-left bg-slate-50 border border-slate-200 rounded-2xl p-4 mb-6 animate-fade-in-up animate-delay-300 cursor-pointer">
                    <div className="flex items-start gap-3">
                      <input
                        type="checkbox"
                        className="mt-1 h-4 w-4 rounded border-slate-300 text-orange-500 focus:ring-orange-500"
                        checked={enableSystemIntegration}
                        onChange={(event) => setEnableSystemIntegration(event.target.checked)}
                        disabled={postInstallBusy}
                      />
                      <div>
                        <div className="flex items-center gap-2 mb-1">
                          <span className="text-sm font-bold text-slate-800">启用后台服务和计划任务</span>
                          <span className="text-[10px] font-bold tracking-wider text-orange-600 bg-orange-100 rounded-full px-2 py-0.5">
                            推荐
                          </span>
                        </div>
                        <p className="text-sm text-slate-500 leading-relaxed">
                          需要管理员授权执行 OpenClaw 的系统集成配置与定时任务等。
                        </p>
                      </div>
                    </div>
                  </label>
                )}

                {currentAction === 'install' && postInstallMessage && (
                  <div className={`w-full rounded-2xl p-4 mb-6 text-left animate-fade-in-up ${postInstallMessage.tone === 'error'
                    ? 'bg-rose-50 border border-rose-200 text-rose-700'
                    : 'bg-amber-50 border border-amber-200 text-amber-700'
                    }`}>
                    <div className="text-xs font-bold uppercase tracking-wider mb-1">
                      {postInstallMessage.tone === 'error' ? '授权失败' : '授权已取消'}
                    </div>
                    <div className="text-sm leading-relaxed">{postInstallMessage.text}</div>
                  </div>
                )}

                {currentAction === 'install' ? (
                  <button
                    className="px-10 py-4 rounded-2xl font-bold text-lg text-white bg-gradient-to-r from-emerald-500 to-emerald-600 hover:from-emerald-600 hover:to-emerald-700 shadow-xl shadow-emerald-500/25 transition-all hover:-translate-y-1 hover:shadow-2xl hover:shadow-emerald-500/30 active:translate-y-0 animate-fade-in-up animate-delay-400 disabled:opacity-70 disabled:hover:translate-y-0"
                    onClick={() => void finishInstall()}
                    disabled={postInstallBusy}
                  >
                    {postInstallBusy
                      ? supportsPostInstallElevation && enableSystemIntegration
                        ? '正在请求管理员授权...'
                        : '正在打开 OpenClaw...'
                      : supportsPostInstallElevation && enableSystemIntegration
                        ? '授权并打开 OpenClaw'
                        : '打开 OpenClaw'}
                  </button>
                ) : (
                  <button
                    className="px-10 py-4 rounded-2xl font-bold text-lg text-white bg-gradient-to-r from-slate-700 to-slate-800 hover:from-slate-800 hover:to-slate-900 shadow-xl shadow-slate-400/20 transition-all hover:-translate-y-1 active:translate-y-0 animate-fade-in-up animate-delay-400"
                    onClick={() => {
                      setAppPhase('home');
                      setCurrentAction('install');
                      setError('');
                      setInstallerSteps([]);
                    }}
                  >
                    返回首页
                  </button>
                )}

                <button
                  className="mt-4 text-sm text-slate-400 hover:text-slate-600 transition-colors animate-fade-in-up animate-delay-500"
                  onClick={() => Quit()}
                  disabled={postInstallBusy}
                >
                  关闭安装程序
                </button>
              </div>
            </div>

            {/* ===== SLIDE 3b: Failed ===== */}
            <div className={`installer-slide ${appPhase === 'failed' ? 'slide-active' : 'slide-right'}`}>
              <div className="flex flex-col items-center text-center px-8 max-w-lg">
                <div className="w-20 h-20 rounded-3xl bg-gradient-to-br from-rose-400 to-rose-600 flex items-center justify-center text-white text-4xl shadow-lg shadow-rose-500/30 mb-8 animate-sparkle">
                  ✗
                </div>

                <h2 className="text-3xl font-black text-slate-900 mb-3 animate-fade-in-up">
                  {currentAction === 'install' ? '安装遇到问题' : '卸载遇到问题'}
                </h2>

                <p className="text-slate-500 text-base leading-relaxed font-medium mb-6 animate-fade-in-up animate-delay-100">
                  {currentAction === 'install' ? '安装过程中遇到了一些障碍，您可以重试安装。' : '卸载过程中遇到了一些障碍，您可以返回首页后重试卸载。'}
                </p>

                {/* Error detail */}
                {(error || installerSteps.some(s => s.status === 'error')) && (
                  <div className="w-full bg-rose-50 border border-rose-200 rounded-2xl p-4 mb-8 text-left animate-fade-in-up animate-delay-200">
                    <div className="text-xs font-bold text-rose-400 uppercase tracking-wider mb-2">错误详情</div>
                    <div className="text-sm text-rose-700 font-mono leading-relaxed break-all">
                      {error || installerSteps.filter(s => s.status === 'error').map(s => s.message).join('\n')}
                    </div>
                  </div>
                )}

                <div className="flex gap-3 animate-fade-in-up animate-delay-300">
                  <button
                    className="px-8 py-3.5 rounded-2xl font-bold text-white bg-gradient-to-r from-orange-500 to-orange-600 hover:from-orange-600 hover:to-orange-700 shadow-lg shadow-orange-500/20 transition-all hover:-translate-y-0.5"
                    onClick={() => {
                      setAppPhase('home');
                      setCurrentAction('install');
                      setError('');
                      setInstallerSteps([]);
                    }}
                  >
                    返回首页
                  </button>
                </div>
              </div>
            </div>

          </div>
        </div>
      </div>
    </main>
  );
}

type WindowChromeProps = {
  isMaximised: boolean;
  onToggleMaximise: () => void;
};

function WindowChrome({ isMaximised, onToggleMaximise }: WindowChromeProps) {
  return (
    <header className="flex items-center justify-between transition-all select-none no-drag drag-zone px-4 py-3 bg-white/40 backdrop-blur-md border-b border-slate-200/50 z-50">
      <div
        className="flex items-center gap-2 flex-1 drag-zone h-full min-w-0"
        onDoubleClick={onToggleMaximise}
      >
        <div className="w-7 h-7 text-[10px] rounded-lg bg-gradient-to-br from-orange-500 to-red-500 flex items-center justify-center text-white font-black tracking-tighter shrink-0 shadow-sm shadow-orange-500/20">
          OC
        </div>
        <div className="flex flex-col min-w-0">
          <span className="text-[13px] font-bold text-slate-800 truncate tracking-wide">
            虾师傅
          </span>
          <span className="text-[10px] text-slate-500 truncate -mt-[1px]">
            OpenClaw-Sifu
          </span>
        </div>
      </div>

      <div className="flex items-center gap-1 shrink-0 no-drag ml-2">
        <button
          className="w-8 h-8 flex items-center justify-center rounded-md text-slate-500 hover:bg-slate-200/50 hover:text-slate-800 transition-colors"
          onClick={() => WindowMinimise()}
          title="最小化"
        >
          <Minus size={15} />
        </button>
        <button
          className="w-8 h-8 flex items-center justify-center rounded-md text-slate-500 hover:bg-slate-200/50 hover:text-slate-800 transition-colors"
          onClick={onToggleMaximise}
          title={isMaximised ? "还原" : "最大化"}
        >
          {isMaximised ? <Copy size={13} /> : <Square size={13} />}
        </button>
        <button
          className="w-8 h-8 flex items-center justify-center rounded-md text-slate-500 hover:bg-rose-500 hover:text-white transition-colors"
          onClick={() => Quit()}
          title="关闭"
        >
          <X size={16} />
        </button>
      </div>
    </header>
  );
}
