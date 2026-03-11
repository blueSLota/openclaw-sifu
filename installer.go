package main

import (
	"bufio"
	"context"
	"fmt"
	"os"
	"os/exec"
	"path/filepath"
	"regexp"
	"strconv"
	"strings"
	"sync"
	"time"

	wruntime "github.com/wailsapp/wails/v2/pkg/runtime"
)

// ---------------------------------------------------------------------------
// Configuration & result types
// ---------------------------------------------------------------------------

// InstallerConfig holds all configuration for an installation run.
type InstallerConfig struct {
	Tag            string `json:"tag"`            // e.g. "latest", "beta"
	InstallMethod  string `json:"installMethod"`  // "npm" or "git"
	GitDir         string `json:"gitDir"`         // clone target for git method
	NoOnboard      bool   `json:"noOnboard"`      // skip onboarding
	NoGitUpdate    bool   `json:"noGitUpdate"`    // skip git pull
	DryRun         bool   `json:"dryRun"`         // print plan, do nothing
	UseCnMirrors   bool   `json:"useCnMirrors"`   // use Chinese npm/download mirrors
	NpmRegistry    string `json:"npmRegistry"`    // custom npm registry
	InstallBaseUrl string `json:"installBaseUrl"` // custom download base URL
	RepoUrl        string `json:"repoUrl"`        // git repository URL
}

// InstallerStepUpdate is emitted for every progress change.
type InstallerStepUpdate struct {
	Step    string `json:"step"`
	Status  string `json:"status"` // "running", "ok", "warn", "error", "skip"
	Message string `json:"message"`
}

// InstallerResult is the final outcome.
type InstallerResult struct {
	Success          bool   `json:"success"`
	InstalledVersion string `json:"installedVersion"`
	IsUpgrade        bool   `json:"isUpgrade"`
	Message          string `json:"message"`
	Error            string `json:"error,omitempty"`
}

// ---------------------------------------------------------------------------
// Entry point — exposed to Wails frontend
// ---------------------------------------------------------------------------

// RunNativeInstaller runs the full installation flow in pure Go.
// It replaces the previous RunBundledInstaller which delegated to install.ps1.
func (a *App) RunNativeInstaller(cfg InstallerConfig) InstallerResult {
	cfg = applyInstallerDefaults(cfg)

	emit := func(step, status, msg string) {
		if a.ctx != nil {
			wruntime.EventsEmit(a.ctx, "installer:step", InstallerStepUpdate{
				Step:    step,
				Status:  status,
				Message: msg,
			})
		}
	}

	emit("init", "running", "OpenClaw Installer")

	// Dry run ---------------------------------------------------------------
	if cfg.DryRun {
		emit("init", "ok", fmt.Sprintf("Dry run — method=%s tag=%s", cfg.InstallMethod, cfg.Tag))
		return InstallerResult{Success: true, Message: "Dry run completed"}
	}

	// Step 1: Check / install Node.js ---------------------------------------
	nodeVer, nodeOk := checkNodeJS(emit)
	if !nodeOk {
		installed := installNodeJS(emit)
		if !installed {
			emit("node", "error", "Node.js 22+ is required but could not be installed automatically")
			return InstallerResult{
				Success: false,
				Error:   "Node.js 22+ is required. Please install it manually from https://nodejs.org/",
			}
		}
		refreshSystemPath()
		nodeVer, nodeOk = checkNodeJS(emit)
		if !nodeOk {
			emit("node", "error", "Node.js still not detected after install — restart may be required")
			return InstallerResult{
				Success: false,
				Error:   "Node.js was installed but not detected. Please restart this application and try again.",
			}
		}
	}
	_ = nodeVer

	// Step 2: Detect existing installation ----------------------------------
	isUpgrade := false
	if existingPath := findOpenClawCommand(); existingPath != "" {
		emit("detect", "ok", fmt.Sprintf("Existing installation found: %s", existingPath))
		isUpgrade = true
	}

	// Step 3: Install OpenClaw ----------------------------------------------
	switch cfg.InstallMethod {
	case "git":
		if !checkCommandExists("git") {
			emit("git", "error", "Git is required for git install method")
			return InstallerResult{
				Success: false,
				Error:   "Git is required. Install from https://git-scm.com/download/win",
			}
		}
		if err := installOpenClawGit(cfg, emit); err != nil {
			return InstallerResult{Success: false, Error: err.Error()}
		}
	default:
		if !checkCommandExists("git") {
			emit("git-check", "warn", "Git not found — npm install may fail for packages with git dependencies")
		}
		if err := installOpenClawNpm(cfg, emit); err != nil {
			return InstallerResult{Success: false, Error: err.Error()}
		}
	}

	// Step 4: Ensure openclaw is on PATH ------------------------------------
	ensureOpenClawOnPath(emit)

	// Step 5: Refresh gateway if loaded -------------------------------------
	refreshGatewayServiceIfLoaded(emit)

	// Step 6: Run doctor for migrations (if upgrading or git) ---------------
	if isUpgrade || cfg.InstallMethod == "git" {
		runDoctor(emit)
	}

	// Step 7: Detect installed version --------------------------------------
	installedVersion := detectInstalledVersion()

	// Step 8: Onboard or Setup ----------------------------------------------
	if !isUpgrade {
		if !cfg.NoOnboard {
			runOnboard(emit)
		} else {
			runSetup(emit)
		}
	}

	msg := "OpenClaw installed successfully!"
	if installedVersion != "" {
		msg = fmt.Sprintf("OpenClaw installed successfully (%s)!", installedVersion)
	}
	emit("done", "ok", msg)

	return InstallerResult{
		Success:          true,
		InstalledVersion: installedVersion,
		IsUpgrade:        isUpgrade,
		Message:          msg,
	}
}

func (a *App) RunNativeUninstaller() InstallerResult {
	emit := func(step, status, msg string) {
		if a.ctx != nil {
			wruntime.EventsEmit(a.ctx, "installer:step", InstallerStepUpdate{
				Step:    step,
				Status:  status,
				Message: msg,
			})
		}
	}

	emit("init", "running", "OpenClaw Uninstaller")

	openclawPath := findOpenClawCommand()
	taskExists := hasOpenClawScheduledTask()
	if openclawPath == "" && !taskExists {
		emit("detect", "error", "OpenClaw command not found")
		return InstallerResult{
			Success: false,
			Error:   "未检测到 OpenClaw 安装。",
		}
	}

	if openclawPath != "" {
		emit("detect", "ok", fmt.Sprintf("OpenClaw found: %s", openclawPath))
		emit("uninstall", "running", "Removing OpenClaw service, state, and workspace...")

		if err := streamCommand("openclaw", []string{"uninstall", "--all", "--yes", "--non-interactive"}, nil, emit, "uninstall"); err != nil {
			emit("uninstall", "error", fmt.Sprintf("OpenClaw uninstall failed: %v", err))
			return InstallerResult{
				Success: false,
				Error:   fmt.Sprintf("OpenClaw 卸载失败: %v", err),
			}
		}
		emit("uninstall", "ok", "OpenClaw local data removed")
	} else {
		emit("detect", "warn", "OpenClaw CLI not found, cleaning residual scheduled task only")
		emit("uninstall", "skip", "Skipped official uninstall because OpenClaw CLI is already missing")
	}

	if taskExists {
		emit("gateway", "running", "Removing OpenClaw scheduled task...")
		if err := removeOpenClawScheduledTask(); err != nil {
			emit("gateway", "error", fmt.Sprintf("Scheduled task cleanup failed: %v", err))
			return InstallerResult{
				Success: false,
				Error:   fmt.Sprintf("删除 OpenClaw 计划任务失败: %v", err),
			}
		}
		emit("gateway", "ok", "OpenClaw scheduled task removed")
	}

	if err := removeInstalledOpenClawCLI(openclawPath, emit); err != nil {
		emit("cli-remove", "warn", fmt.Sprintf("CLI cleanup incomplete: %v", err))
	} else {
		emit("cli-remove", "ok", "OpenClaw CLI removed")
	}

	refreshSystemPath()

	msg := "OpenClaw uninstalled successfully!"
	emit("done", "ok", msg)
	return InstallerResult{
		Success: true,
		Message: msg,
	}
}

// ---------------------------------------------------------------------------
// Defaults from environment
// ---------------------------------------------------------------------------

func applyInstallerDefaults(cfg InstallerConfig) InstallerConfig {
	if cfg.Tag == "" {
		cfg.Tag = "latest"
	}
	if cfg.InstallMethod == "" {
		cfg.InstallMethod = envOrDefault("OPENCLAW_INSTALL_METHOD", "npm")
	}
	if cfg.GitDir == "" {
		cfg.GitDir = envOrDefault("OPENCLAW_GIT_DIR", "")
		if cfg.GitDir == "" {
			home, _ := os.UserHomeDir()
			cfg.GitDir = filepath.Join(home, "openclaw")
		}
	}
	if os.Getenv("OPENCLAW_NO_ONBOARD") == "1" {
		cfg.NoOnboard = true
	}
	if os.Getenv("OPENCLAW_GIT_UPDATE") == "0" {
		cfg.NoGitUpdate = true
	}
	if os.Getenv("OPENCLAW_DRY_RUN") == "1" {
		cfg.DryRun = true
	}

	useCn := os.Getenv("OPENCLAW_USE_CN_MIRRORS") == "1" || os.Getenv("OPENCLAW_MIRROR_PRESET") == "cn"
	if cfg.UseCnMirrors || useCn {
		cfg.UseCnMirrors = true
	}

	if cfg.NpmRegistry == "" {
		cfg.NpmRegistry = os.Getenv("OPENCLAW_NPM_REGISTRY")
		if cfg.NpmRegistry == "" && cfg.UseCnMirrors {
			cfg.NpmRegistry = "https://registry.npmmirror.com"
		}
	}

	if cfg.InstallBaseUrl == "" {
		cfg.InstallBaseUrl = os.Getenv("OPENCLAW_INSTALL_BASE_URL")
		if cfg.InstallBaseUrl == "" {
			if cfg.UseCnMirrors {
				cfg.InstallBaseUrl = "https://clawd.org.cn"
			} else {
				cfg.InstallBaseUrl = "https://openclaw.ai"
			}
		}
	}
	cfg.InstallBaseUrl = strings.TrimRight(cfg.InstallBaseUrl, "/")

	if cfg.RepoUrl == "" {
		cfg.RepoUrl = envOrDefault("OPENCLAW_GIT_REPO_URL", "https://github.com/openclaw/openclaw.git")
	}

	return cfg
}

func envOrDefault(key, fallback string) string {
	if v := os.Getenv(key); v != "" {
		return v
	}
	return fallback
}

// ---------------------------------------------------------------------------
// Step: Node.js
// ---------------------------------------------------------------------------

var nodeVersionRe = regexp.MustCompile(`v(\d+)\.`)

func checkNodeJS(emit func(string, string, string)) (string, bool) {
	emit("node", "running", "Checking Node.js...")

	out, err := execOutput("node", "-v")
	if err != nil {
		emit("node", "warn", "Node.js not found")
		return "", false
	}

	ver := strings.TrimSpace(out)
	matches := nodeVersionRe.FindStringSubmatch(ver)
	if len(matches) < 2 {
		emit("node", "warn", fmt.Sprintf("Could not parse Node.js version: %s", ver))
		return ver, false
	}

	major, _ := strconv.Atoi(matches[1])
	if major < 22 {
		emit("node", "warn", fmt.Sprintf("Node.js %s found but v22+ required", ver))
		return ver, false
	}

	emit("node", "ok", fmt.Sprintf("Node.js %s found", ver))
	return ver, true
}

func installNodeJS(emit func(string, string, string)) bool {
	emit("node-install", "running", "Installing Node.js...")

	// Try winget
	if checkCommandExists("winget") {
		emit("node-install", "running", "Installing Node.js via winget...")
		err := streamCommand("winget", []string{
			"install", "OpenJS.NodeJS.LTS",
			"--accept-package-agreements", "--accept-source-agreements",
		}, nil, emit, "node-install")
		if err == nil {
			emit("node-install", "ok", "Node.js installed via winget")
			return true
		}
		emit("node-install", "warn", fmt.Sprintf("winget install failed: %v", err))
	}

	// Try Chocolatey
	if checkCommandExists("choco") {
		emit("node-install", "running", "Installing Node.js via Chocolatey...")
		err := streamCommand("choco", []string{"install", "nodejs-lts", "-y"}, nil, emit, "node-install")
		if err == nil {
			emit("node-install", "ok", "Node.js installed via Chocolatey")
			return true
		}
		emit("node-install", "warn", fmt.Sprintf("Chocolatey install failed: %v", err))
	}

	// Try Scoop
	if checkCommandExists("scoop") {
		emit("node-install", "running", "Installing Node.js via Scoop...")
		err := streamCommand("scoop", []string{"install", "nodejs-lts"}, nil, emit, "node-install")
		if err == nil {
			emit("node-install", "ok", "Node.js installed via Scoop")
			return true
		}
		emit("node-install", "warn", fmt.Sprintf("Scoop install failed: %v", err))
	}

	emit("node-install", "error", "No package manager found (winget, choco, scoop)")
	return false
}

// ---------------------------------------------------------------------------
// Step: Detect existing OpenClaw
// ---------------------------------------------------------------------------

func findOpenClawCommand() string {
	for _, name := range []string{"openclaw.cmd", "openclaw"} {
		if p, err := exec.LookPath(name); err == nil {
			return p
		}
	}
	return ""
}

func detectInstalledVersion() string {
	out, err := execOutput("openclaw", "--version")
	if err == nil {
		v := strings.TrimSpace(out)
		if v != "" {
			return v
		}
	}

	// Fallback: npm list
	out, err = execOutput("npm", "list", "-g", "--depth", "0", "--json")
	if err == nil && strings.Contains(out, "openclaw") {
		// simple parse — look for "version": "x.y.z" after "openclaw"
		idx := strings.Index(out, `"openclaw"`)
		if idx >= 0 {
			sub := out[idx:]
			vIdx := strings.Index(sub, `"version"`)
			if vIdx >= 0 {
				sub = sub[vIdx:]
				start := strings.Index(sub, `"`)
				if start >= 0 {
					sub = sub[start+1:]
					start = strings.Index(sub, `"`)
					if start >= 0 {
						sub = sub[start+1:]
						end := strings.Index(sub, `"`)
						if end >= 0 {
							return sub[:end]
						}
					}
				}
			}
		}
	}

	return ""
}

// ---------------------------------------------------------------------------
// Step: Install OpenClaw via npm
// ---------------------------------------------------------------------------

func installOpenClawNpm(cfg InstallerConfig, emit func(string, string, string)) error {
	packageName := "openclaw"
	emit("npm-install", "running", fmt.Sprintf("Installing %s@%s via npm...", packageName, cfg.Tag))

	env := buildNpmEnv(cfg)
	args := []string{"install", "-g", fmt.Sprintf("%s@%s", packageName, cfg.Tag)}

	err := streamCommand("npm", args, env, emit, "npm-install")
	if err != nil {
		emit("npm-install", "error", fmt.Sprintf("npm install failed: %v", err))
		return fmt.Errorf("npm install failed: %w", err)
	}

	emit("npm-install", "ok", "OpenClaw installed via npm")
	return nil
}

func buildNpmEnv(cfg InstallerConfig) []string {
	env := os.Environ()
	// Suppress noise
	env = setEnv(env, "NPM_CONFIG_LOGLEVEL", "error")
	env = setEnv(env, "NPM_CONFIG_UPDATE_NOTIFIER", "false")
	env = setEnv(env, "NPM_CONFIG_FUND", "false")
	env = setEnv(env, "NPM_CONFIG_AUDIT", "false")
	// Avoid PowerShell lifecycle scripts — use cmd.exe instead
	env = setEnv(env, "NPM_CONFIG_SCRIPT_SHELL", "cmd.exe")

	if cfg.NpmRegistry != "" {
		env = setEnv(env, "NPM_CONFIG_REGISTRY", cfg.NpmRegistry)
		env = setEnv(env, "npm_config_registry", cfg.NpmRegistry)
		env = setEnv(env, "COREPACK_NPM_REGISTRY", cfg.NpmRegistry)
	}

	return env
}

func removeInstalledOpenClawCLI(openclawPath string, emit func(string, string, string)) error {
	removedAny := false

	if checkCommandExists("npm") {
		emit("cli-remove", "running", "Removing global OpenClaw CLI package...")
		if err := streamCommand("npm", []string{"uninstall", "-g", "openclaw"}, buildNpmEnv(InstallerConfig{}), emit, "cli-remove"); err == nil {
			removedAny = true
		} else {
			return fmt.Errorf("npm uninstall failed: %w", err)
		}
	}

	userHome, _ := os.UserHomeDir()
	wrapperDir := filepath.Join(userHome, ".local", "bin")
	if strings.HasPrefix(strings.ToLower(openclawPath), strings.ToLower(wrapperDir)) {
		if err := os.Remove(openclawPath); err == nil || os.IsNotExist(err) {
			removedAny = true
		} else {
			return fmt.Errorf("remove wrapper %s: %w", openclawPath, err)
		}
	}

	if !removedAny && emit != nil {
		emit("cli-remove", "skip", "No removable CLI package detected")
	}

	return nil
}

// ---------------------------------------------------------------------------
// Step: Install OpenClaw from Git
// ---------------------------------------------------------------------------

func installOpenClawGit(cfg InstallerConfig, emit func(string, string, string)) error {
	emit("git-install", "running", fmt.Sprintf("Installing from %s...", cfg.RepoUrl))

	// Ensure pnpm
	if !checkCommandExists("pnpm") {
		emit("git-install", "running", "Installing pnpm...")
		env := buildNpmEnv(cfg)
		err := streamCommand("npm", []string{"install", "-g", "pnpm"}, env, emit, "git-install")
		if err != nil {
			return fmt.Errorf("failed to install pnpm: %w", err)
		}
	}

	// Clone or pull
	if _, err := os.Stat(cfg.GitDir); os.IsNotExist(err) {
		emit("git-install", "running", "Cloning repository...")
		if err := streamCommand("git", []string{"clone", cfg.RepoUrl, cfg.GitDir}, nil, emit, "git-install"); err != nil {
			return fmt.Errorf("git clone failed: %w", err)
		}
	} else if !cfg.NoGitUpdate {
		// Check if repo is dirty
		out, _ := execOutput("git", "-C", cfg.GitDir, "status", "--porcelain")
		if strings.TrimSpace(out) == "" {
			emit("git-install", "running", "Pulling latest changes...")
			_ = streamCommand("git", []string{"-C", cfg.GitDir, "pull", "--rebase"}, nil, emit, "git-install")
		} else {
			emit("git-install", "warn", "Repository has uncommitted changes; skipping git pull")
		}
	}

	// Remove legacy submodule
	legacyDir := filepath.Join(cfg.GitDir, "Peekaboo")
	if info, err := os.Stat(legacyDir); err == nil && info.IsDir() {
		emit("git-install", "running", "Removing legacy submodule...")
		_ = os.RemoveAll(legacyDir)
	}

	env := buildNpmEnv(cfg)

	// pnpm install
	emit("git-install", "running", "Installing dependencies with pnpm...")
	if err := streamCommand("pnpm", []string{"-C", cfg.GitDir, "install"}, env, emit, "git-install"); err != nil {
		return fmt.Errorf("pnpm install failed: %w", err)
	}

	// Build UI
	emit("git-install", "running", "Building UI...")
	_ = streamCommand("pnpm", []string{"-C", cfg.GitDir, "ui:build"}, env, emit, "git-install")

	// Build
	emit("git-install", "running", "Building OpenClaw...")
	if err := streamCommand("pnpm", []string{"-C", cfg.GitDir, "build"}, env, emit, "git-install"); err != nil {
		return fmt.Errorf("build failed: %w", err)
	}

	// Create wrapper script
	binDir := filepath.Join(os.Getenv("USERPROFILE"), ".local", "bin")
	if err := os.MkdirAll(binDir, 0o755); err != nil {
		return fmt.Errorf("create bin dir: %w", err)
	}

	cmdPath := filepath.Join(binDir, "openclaw.cmd")
	cmdContent := fmt.Sprintf("@echo off\r\nnode \"%s\" %%*\r\n", filepath.Join(cfg.GitDir, "dist", "entry.js"))
	if err := os.WriteFile(cmdPath, []byte(cmdContent), 0o644); err != nil {
		return fmt.Errorf("write wrapper script: %w", err)
	}

	addToUserPath(binDir, emit)

	emit("git-install", "ok", fmt.Sprintf("OpenClaw installed from source to %s", cmdPath))
	return nil
}

// ---------------------------------------------------------------------------
// Step: Ensure openclaw on PATH
// ---------------------------------------------------------------------------

func ensureOpenClawOnPath(emit func(string, string, string)) {
	if findOpenClawCommand() != "" {
		return
	}

	emit("path", "running", "Checking PATH for openclaw...")
	refreshSystemPath()

	if findOpenClawCommand() != "" {
		return
	}

	// Try to find it in known npm global dirs and add to PATH
	npmPrefix, err := execOutput("npm", "config", "get", "prefix")
	if err != nil {
		emit("path", "warn", "Could not determine npm prefix")
		return
	}
	npmPrefix = strings.TrimSpace(npmPrefix)

	candidates := []string{}
	if npmPrefix != "" {
		candidates = append(candidates, npmPrefix, filepath.Join(npmPrefix, "bin"))
	}
	appdata := os.Getenv("APPDATA")
	if appdata != "" {
		candidates = append(candidates, filepath.Join(appdata, "npm"))
	}

	for _, dir := range candidates {
		cmdPath := filepath.Join(dir, "openclaw.cmd")
		if _, err := os.Stat(cmdPath); err == nil {
			addToUserPath(dir, emit)
			refreshSystemPath()
			return
		}
	}

	emit("path", "warn", "openclaw not found on PATH — restart terminal may be required")
}

// ---------------------------------------------------------------------------
// Step: Gateway service refresh
// ---------------------------------------------------------------------------

func refreshGatewayServiceIfLoaded(emit func(string, string, string)) {
	if findOpenClawCommand() == "" {
		return
	}

	// Check if gateway service is loaded
	out, err := execOutput("openclaw", "daemon", "status", "--json")
	if err != nil || !strings.Contains(out, `"loaded"`) || !strings.Contains(out, "true") {
		return
	}

	emit("gateway", "running", "Refreshing gateway service...")

	// gateway install --force requires admin (schtasks create).
	// Run it quietly; if it fails (e.g. access denied), just skip.
	installErr := runQuietCommand("openclaw", "gateway", "install", "--force")
	if installErr != nil {
		emit("gateway", "skip", "Gateway service refresh deferred to the completion screen (requires admin privileges)")
		return
	}

	_ = runQuietCommand("openclaw", "gateway", "restart")
	emit("gateway", "ok", "Gateway service refreshed")
}

// ---------------------------------------------------------------------------
// Step: Doctor & Onboard
// ---------------------------------------------------------------------------

func runDoctor(emit func(string, string, string)) {
	emit("doctor", "running", "Running doctor to migrate settings...")
	_ = streamCommand("openclaw", []string{"doctor", "--non-interactive"}, nil, emit, "doctor")
	emit("doctor", "ok", "Migration complete")
}

func runOnboard(emit func(string, string, string)) {
	emit("onboard", "running", "Starting first-time setup...")
	_ = streamCommand("openclaw", []string{"onboard"}, nil, emit, "onboard")
	emit("onboard", "ok", "Setup complete")
}

func runSetup(emit func(string, string, string)) {
	emit("setup", "running", "Initializing local configuration...")
	_ = streamCommand("openclaw", []string{"setup"}, nil, emit, "setup")
	emit("setup", "ok", "Configuration initialized")
}

// ---------------------------------------------------------------------------
// PATH helpers (cross-platform safe, Windows-specific in installer_windows.go)
// ---------------------------------------------------------------------------

// refreshSystemPath reloads the current process PATH from system + user vars.
func refreshSystemPath() {
	machine := os.Getenv("Path")
	// On Windows, read from registry for the most up-to-date values.
	// This is a simple refresh — the Windows-specific file can do registry reads.
	machineEnv, _ := readRegistryPath("SYSTEM\\CurrentControlSet\\Control\\Session Manager\\Environment", "Path")
	userEnv, _ := readRegistryPath("Environment", "Path")

	if machineEnv != "" || userEnv != "" {
		combined := machineEnv + ";" + userEnv
		os.Setenv("Path", combined)
	} else {
		_ = machine // keep current
	}
}

// ---------------------------------------------------------------------------
// Command execution helpers
// ---------------------------------------------------------------------------

func checkCommandExists(name string) bool {
	_, err := exec.LookPath(name)
	return err == nil
}

func hasOpenClawScheduledTask() bool {
	cmd := exec.Command("schtasks", "/Query", "/TN", "OpenClaw Gateway")
	return cmd.Run() == nil
}

// execOutput runs a command and returns its combined stdout as a string.
func execOutput(name string, args ...string) (string, error) {
	cmd := exec.Command(name, args...)
	out, err := cmd.Output()
	return string(out), err
}

// streamCommand runs a command, streaming output line-by-line to the emit callback.
// Returns nil on exit code 0, error otherwise.
func streamCommand(name string, args []string, env []string, emit func(string, string, string), step string) error {
	path, err := exec.LookPath(name)
	if err != nil {
		return fmt.Errorf("command not found: %s", name)
	}

	ctx, cancel := context.WithTimeout(context.Background(), 30*time.Minute)
	defer cancel()

	cmd := exec.CommandContext(ctx, path, args...)
	if env != nil {
		cmd.Env = env
	}

	stdout, err := cmd.StdoutPipe()
	if err != nil {
		return fmt.Errorf("stdout pipe: %w", err)
	}
	stderr, err := cmd.StderrPipe()
	if err != nil {
		return fmt.Errorf("stderr pipe: %w", err)
	}

	if err := cmd.Start(); err != nil {
		return fmt.Errorf("start %s: %w", name, err)
	}

	var wg sync.WaitGroup
	wg.Add(2)

	// Use decodeOutputBytes (defined in executor.go) to handle GBK / GB18030
	// output from Windows system tools (schtasks, winget, etc.).
	go func() {
		defer wg.Done()
		scanner := bufio.NewScanner(stdout)
		for scanner.Scan() {
			line := sanitizeOutput(decodeOutputBytes(scanner.Bytes()))
			if line != "" {
				emit(step, "running", line)
			}
		}
	}()

	go func() {
		defer wg.Done()
		scanner := bufio.NewScanner(stderr)
		for scanner.Scan() {
			line := sanitizeOutput(decodeOutputBytes(scanner.Bytes()))
			if line != "" {
				emit(step, "running", line)
			}
		}
	}()

	wg.Wait()

	if err := cmd.Wait(); err != nil {
		return fmt.Errorf("%s exited with error: %w", name, err)
	}

	return nil
}

// setEnv adds or replaces an environment variable in a []string slice.
func setEnv(env []string, key, value string) []string {
	prefix := strings.ToUpper(key) + "="
	for i, e := range env {
		if strings.HasPrefix(strings.ToUpper(e), prefix) {
			env[i] = key + "=" + value
			return env
		}
	}
	return append(env, key+"="+value)
}

// runQuietCommand runs a command silently and returns any error.
// It does NOT stream output to the UI — used for commands where
// raw output would be confusing (e.g. gateway install needing admin).
func runQuietCommand(name string, args ...string) error {
	cmd := exec.Command(name, args...)
	_, err := cmd.CombinedOutput()
	return err
}
