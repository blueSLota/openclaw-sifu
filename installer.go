package main

import (
	"bufio"
	"context"
	"encoding/json"
	"errors"
	"fmt"
	"io"
	"net/http"
	"os"
	"os/exec"
	"path/filepath"
	"regexp"
	"runtime"
	"sort"
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
		installed := installNodeJS(cfg, emit)
		if !installed {
			emit("node", "error", "Node.js 22+ is required but could not be installed automatically")
			return InstallerResult{
				Success: false,
				Error:   fmt.Sprintf("Node.js %s+ is required. Please install it manually from %s.", formatSemverParts(minimumNodeVersion), nodeMirrorDownloadURL()),
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
		if err := ensureGitAvailable(cfg, emit); err != nil {
			emit("git-check", "error", err.Error())
			return InstallerResult{
				Success: false,
				Error:   err.Error(),
			}
		}
		if err := installOpenClawGit(cfg, emit); err != nil {
			return InstallerResult{Success: false, Error: err.Error()}
		}
	default:
		if err := ensureGitAvailable(cfg, emit); err != nil {
			return InstallerResult{Success: false, Error: err.Error()}
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
	// Keep all installer download paths on domestic mirrors.
	cfg.UseCnMirrors = true

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
	if cfg.UseCnMirrors {
		cfg.RepoUrl = rewriteGitHubURLToMirror(cfg.RepoUrl)
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

var nodeVersionRe = regexp.MustCompile(`v?(\d+)\.(\d+)\.(\d+)`)

type semverParts struct {
	Major int
	Minor int
	Patch int
}

var minimumNodeVersion = semverParts{Major: 22, Minor: 12, Patch: 0}

func checkNodeJS(emit func(string, string, string)) (string, bool) {
	emit("node", "running", "Checking Node.js...")

	out, err := execOutput("node", "-v")
	if err != nil {
		emit("node", "warn", "Node.js not found")
		return "", false
	}

	ver := strings.TrimSpace(out)
	parsed, ok := parseSemverParts(ver)
	if !ok {
		emit("node", "warn", fmt.Sprintf("Could not parse Node.js version: %s", ver))
		return ver, false
	}

	if compareSemverParts(parsed, minimumNodeVersion) < 0 {
		emit("node", "warn", fmt.Sprintf("Node.js %s found but v%s+ required", ver, formatSemverParts(minimumNodeVersion)))
		return ver, false
	}

	emit("node", "ok", fmt.Sprintf("Node.js %s found", ver))
	return ver, true
}

func installNodeJS(cfg InstallerConfig, emit func(string, string, string)) bool {
	emit("node-install", "running", "Installing Node.js...")

	if cfg.UseCnMirrors && runtime.GOOS == "windows" {
		emit("node-install", "running", "Installing Node.js from domestic mirror...")
		if err := installNodeFromMirror(emit); err == nil {
			emit("node-install", "ok", "Node.js installed from domestic mirror")
			return true
		} else {
			emit("node-install", "warn", fmt.Sprintf("mirror install failed: %v", err))
			emit("node-install", "error", fmt.Sprintf("Mirror install failed. Please use the mirror download page instead: %s", nodeMirrorDownloadURL()))
			return false
		}
	}

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

	emit("node-install", "error", fmt.Sprintf("No package manager found (winget, choco, scoop). Use the mirror download page instead: %s", nodeMirrorDownloadURL()))
	return false
}

type nodeMirrorVersion struct {
	Version string   `json:"version"`
	Files   []string `json:"files"`
}

func installNodeFromMirror(emit func(string, string, string)) error {
	version, assetName, err := latestNodeMirrorInstaller()
	if err != nil {
		return err
	}

	assetURL := fmt.Sprintf("%s%s/%s", nodeMirrorBinaryAPI(), version, assetName)
	emit("node-install", "running", fmt.Sprintf("Downloading Node.js installer from mirror: %s", assetName))
	installerPath, err := downloadMirrorAsset(assetURL, assetName)
	if err != nil {
		return fmt.Errorf("download node mirror asset: %w", err)
	}
	defer os.Remove(installerPath)

	emit("node-install", "running", "Running Node.js mirror installer...")
	if strings.HasSuffix(strings.ToLower(assetName), ".msi") {
		return streamCommand("msiexec.exe", []string{"/i", installerPath, "/qn", "/norestart"}, nil, emit, "node-install")
	}

	return streamCommand(installerPath, []string{"/quiet"}, nil, emit, "node-install")
}

func latestNodeMirrorInstaller() (string, string, error) {
	items, err := fetchNodeMirrorVersions()
	if err != nil {
		return "", "", err
	}

	assetMarker, assetName := nodeMirrorAssetSpec()
	candidates := make([]nodeMirrorVersion, 0, len(items))
	for _, item := range items {
		parsed, ok := parseSemverParts(item.Version)
		if !ok || parsed.Major != 22 || compareSemverParts(parsed, minimumNodeVersion) < 0 {
			continue
		}
		if containsString(item.Files, assetMarker) {
			candidates = append(candidates, item)
		}
	}

	if len(candidates) == 0 {
		return "", "", fmt.Errorf("no Node.js mirror installer found for %s and Node %s+", runtime.GOARCH, formatSemverParts(minimumNodeVersion))
	}

	sort.Slice(candidates, func(i, j int) bool {
		left, _ := parseSemverParts(candidates[i].Version)
		right, _ := parseSemverParts(candidates[j].Version)
		return compareSemverParts(left, right) > 0
	})

	return candidates[0].Version, fmt.Sprintf("node-%s-%s", candidates[0].Version, assetName), nil
}

func fetchNodeMirrorVersions() ([]nodeMirrorVersion, error) {
	client := &http.Client{Timeout: 2 * time.Minute}
	resp, err := client.Get(nodeMirrorIndexAPI())
	if err != nil {
		return nil, err
	}
	defer resp.Body.Close()

	if resp.StatusCode < 200 || resp.StatusCode >= 300 {
		return nil, fmt.Errorf("unexpected status %s", resp.Status)
	}

	var items []nodeMirrorVersion
	if err := json.NewDecoder(resp.Body).Decode(&items); err != nil {
		return nil, err
	}
	return items, nil
}

func nodeMirrorAssetSpec() (string, string) {
	switch runtime.GOARCH {
	case "arm64":
		return "win-arm64-msi", "arm64.msi"
	case "386":
		return "win-x86-msi", "x86.msi"
	default:
		return "win-x64-msi", "x64.msi"
	}
}

func ensureGitAvailable(cfg InstallerConfig, emit func(string, string, string)) error {
	if checkCommandExists("git") {
		emit("git-check", "ok", "Git found")
		return nil
	}

	emit("git-check", "running", "Git not found. Installing Git...")
	if !installGit(cfg, emit) {
		return fmt.Errorf("Git is required by the current OpenClaw install flow. Please install it manually from %s and retry.", gitMirrorDownloadURL())
	}

	refreshSystemPath()
	if checkCommandExists("git") {
		emit("git-check", "ok", "Git installed")
		return nil
	}

	return fmt.Errorf("Git was installed but is still not detected. Please restart this application and retry. If needed, install Git manually from %s.", gitMirrorDownloadURL())
}

func installGit(cfg InstallerConfig, emit func(string, string, string)) bool {
	if cfg.UseCnMirrors && runtime.GOOS == "windows" {
		emit("git-check", "running", "Installing Git from domestic mirror...")
		if err := installGitFromMirror(emit); err == nil {
			emit("git-check", "ok", "Git installed from domestic mirror")
			return true
		} else {
			emit("git-check", "warn", fmt.Sprintf("mirror install failed: %v", err))
			emit("git-check", "error", fmt.Sprintf("Mirror install failed. Please use the mirror download page instead: %s", gitMirrorDownloadURL()))
			return false
		}
	}

	if checkCommandExists("winget") {
		emit("git-check", "running", "Installing Git via winget...")
		err := streamCommand("winget", []string{
			"install", "Git.Git",
			"--accept-package-agreements", "--accept-source-agreements",
		}, nil, emit, "git-check")
		if err == nil {
			emit("git-check", "ok", "Git installed via winget")
			return true
		}
		emit("git-check", "warn", fmt.Sprintf("winget install failed: %v", err))
	}

	if checkCommandExists("choco") {
		emit("git-check", "running", "Installing Git via Chocolatey...")
		err := streamCommand("choco", []string{"install", "git", "-y"}, nil, emit, "git-check")
		if err == nil {
			emit("git-check", "ok", "Git installed via Chocolatey")
			return true
		}
		emit("git-check", "warn", fmt.Sprintf("Chocolatey install failed: %v", err))
	}

	if checkCommandExists("scoop") {
		emit("git-check", "running", "Installing Git via Scoop...")
		err := streamCommand("scoop", []string{"install", "git"}, nil, emit, "git-check")
		if err == nil {
			emit("git-check", "ok", "Git installed via Scoop")
			return true
		}
		emit("git-check", "warn", fmt.Sprintf("Scoop install failed: %v", err))
	}

	emit("git-check", "error", fmt.Sprintf("No package manager found to install Git automatically (winget, choco, scoop). Use the mirror download page instead: %s", gitMirrorDownloadURL()))
	return false
}

type mirrorIndexItem struct {
	Name     string `json:"name"`
	Type     string `json:"type"`
	URL      string `json:"url"`
	Modified string `json:"modified"`
}

type gitMirrorRelease struct {
	item    mirrorIndexItem
	version [4]int
}

var gitMirrorReleaseRe = regexp.MustCompile(`^v(\d+)\.(\d+)\.(\d+)\.windows\.(\d+)/$`)

func installGitFromMirror(emit func(string, string, string)) error {
	release, err := latestGitMirrorRelease()
	if err != nil {
		return err
	}

	asset, err := selectGitMirrorInstaller(release.item.URL)
	if err != nil {
		return err
	}

	emit("git-check", "running", fmt.Sprintf("Downloading Git installer from mirror: %s", asset.Name))
	installerPath, err := downloadMirrorAsset(asset.URL, asset.Name)
	if err != nil {
		return fmt.Errorf("download mirror asset: %w", err)
	}
	defer os.Remove(installerPath)

	emit("git-check", "running", "Running Git mirror installer...")
	if err := streamCommand(installerPath, []string{"/VERYSILENT", "/NORESTART", "/NOCANCEL", "/SP-"}, nil, emit, "git-check"); err != nil {
		return fmt.Errorf("run mirror installer: %w", err)
	}

	return nil
}

func latestGitMirrorRelease() (gitMirrorRelease, error) {
	items, err := fetchMirrorIndex(gitMirrorBinaryAPI())
	if err != nil {
		return gitMirrorRelease{}, fmt.Errorf("fetch git mirror index: %w", err)
	}

	releases := make([]gitMirrorRelease, 0, len(items))
	for _, item := range items {
		if item.Type != "dir" {
			continue
		}

		matches := gitMirrorReleaseRe.FindStringSubmatch(item.Name)
		if len(matches) != 5 {
			continue
		}

		version := [4]int{}
		valid := true
		for i := 1; i < len(matches); i++ {
			value, convErr := strconv.Atoi(matches[i])
			if convErr != nil {
				valid = false
				break
			}
			version[i-1] = value
		}
		if !valid {
			continue
		}

		releases = append(releases, gitMirrorRelease{item: item, version: version})
	}

	if len(releases) == 0 {
		return gitMirrorRelease{}, fmt.Errorf("no stable Git for Windows releases found on mirror")
	}

	sort.Slice(releases, func(i, j int) bool {
		for idx := 0; idx < len(releases[i].version); idx++ {
			if releases[i].version[idx] != releases[j].version[idx] {
				return releases[i].version[idx] > releases[j].version[idx]
			}
		}
		return releases[i].item.Name > releases[j].item.Name
	})

	return releases[0], nil
}

func selectGitMirrorInstaller(releaseURL string) (mirrorIndexItem, error) {
	items, err := fetchMirrorIndex(releaseURL)
	if err != nil {
		return mirrorIndexItem{}, fmt.Errorf("fetch git release assets: %w", err)
	}

	suffixes := []string{}
	switch runtime.GOARCH {
	case "amd64":
		suffixes = []string{"-64-bit.exe"}
	case "arm64":
		suffixes = []string{"-arm64.exe"}
	case "386":
		suffixes = []string{"-32-bit.exe"}
	default:
		suffixes = []string{"-64-bit.exe", "-arm64.exe"}
	}

	for _, suffix := range suffixes {
		for _, item := range items {
			if item.Type == "file" && strings.HasPrefix(item.Name, "Git-") && strings.HasSuffix(item.Name, suffix) {
				return item, nil
			}
		}
	}

	return mirrorIndexItem{}, fmt.Errorf("no Git installer found for %s on mirror", runtime.GOARCH)
}

func fetchMirrorIndex(url string) ([]mirrorIndexItem, error) {
	client := &http.Client{Timeout: 2 * time.Minute}
	resp, err := client.Get(url)
	if err != nil {
		return nil, err
	}
	defer resp.Body.Close()

	if resp.StatusCode < 200 || resp.StatusCode >= 300 {
		return nil, fmt.Errorf("unexpected status %s", resp.Status)
	}

	var items []mirrorIndexItem
	if err := json.NewDecoder(resp.Body).Decode(&items); err != nil {
		return nil, err
	}
	return items, nil
}

func downloadMirrorAsset(url, fileName string) (string, error) {
	client := &http.Client{Timeout: 30 * time.Minute}
	resp, err := client.Get(url)
	if err != nil {
		return "", err
	}
	defer resp.Body.Close()

	if resp.StatusCode < 200 || resp.StatusCode >= 300 {
		return "", fmt.Errorf("unexpected status %s", resp.Status)
	}

	ext := filepath.Ext(fileName)
	tempFile, err := os.CreateTemp("", "openclaw-mirror-*"+ext)
	if err != nil {
		return "", err
	}
	defer tempFile.Close()

	if _, err := io.Copy(tempFile, resp.Body); err != nil {
		os.Remove(tempFile.Name())
		return "", err
	}

	return tempFile.Name(), nil
}

func gitMirrorBinaryAPI() string {
	return "https://registry.npmmirror.com/-/binary/git-for-windows/"
}

func nodeMirrorBinaryAPI() string {
	return "https://registry.npmmirror.com/-/binary/node/"
}

func nodeMirrorIndexAPI() string {
	return "https://registry.npmmirror.com/-/binary/node/index.json"
}

func gitMirrorDownloadURL() string {
	return "https://npmmirror.com/mirrors/git-for-windows/"
}

func nodeMirrorDownloadURL() string {
	return "https://npmmirror.com/mirrors/node/"
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
	if runtime.GOOS == "windows" {
		emit("npm-install", "warn", "Enabling Windows compatibility mode for native AI bindings during install")
	}
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
	// Avoid PowerShell lifecycle scripts. On Windows, prefer an absolute cmd.exe
	// path so npm lifecycle hooks do not fail with ENOENT when PATH is incomplete.
	if runtime.GOOS == "windows" {
		if cmdShell := resolveWindowsCmdPath(); cmdShell != "" {
			env = setEnv(env, "ComSpec", cmdShell)
			env = setEnv(env, "NPM_CONFIG_SCRIPT_SHELL", cmdShell)
		}
		if systemRoot := os.Getenv("SystemRoot"); systemRoot != "" {
			env = setEnv(env, "SystemRoot", systemRoot)
		}
	} else {
		env = setEnv(env, "NPM_CONFIG_SCRIPT_SHELL", "sh")
	}
	// node-llama-cpp performs native binary probing in postinstall, which
	// crashes on some Windows machines. Skip that install-time step and let
	// users download/build local runtime support later if they need it.
	if runtime.GOOS == "windows" {
		env = setEnv(env, "NODE_LLAMA_CPP_SKIP_DOWNLOAD", "true")
	}
	env = forceGitHubMirrorForNpm(cfg, env)

	if cfg.NpmRegistry != "" {
		env = setEnv(env, "NPM_CONFIG_REGISTRY", cfg.NpmRegistry)
		env = setEnv(env, "npm_config_registry", cfg.NpmRegistry)
		env = setEnv(env, "COREPACK_NPM_REGISTRY", cfg.NpmRegistry)
	}

	return env
}

func forceGitHubMirrorForNpm(cfg InstallerConfig, env []string) []string {
	emptyConfigPath := ensureEmptyGitConfigPath()
	if emptyConfigPath != "" {
		env = setEnv(env, "GIT_CONFIG_GLOBAL", emptyConfigPath)
	}
	env = setEnv(env, "GIT_CONFIG_NOSYSTEM", "1")

	replacement := "https://github.com/"
	if cfg.UseCnMirrors {
		replacement = gitHubCloneMirrorPrefix()
	}

	rewritePairs := []struct {
		key   string
		value string
	}{
		{key: fmt.Sprintf("url.%s.insteadOf", replacement), value: "ssh://git@github.com/"},
		{key: fmt.Sprintf("url.%s.insteadOf", replacement), value: "git@github.com:"},
		{key: fmt.Sprintf("url.%s.insteadOf", replacement), value: "git+ssh://git@github.com/"},
		{key: fmt.Sprintf("url.%s.insteadOf", replacement), value: "git://github.com/"},
	}
	if cfg.UseCnMirrors {
		rewritePairs = append(rewritePairs, struct {
			key   string
			value string
		}{key: fmt.Sprintf("url.%s.insteadOf", replacement), value: "https://github.com/"})
	}

	env = setEnv(env, "GIT_CONFIG_COUNT", strconv.Itoa(len(rewritePairs)))
	for index, pair := range rewritePairs {
		env = setEnv(env, fmt.Sprintf("GIT_CONFIG_KEY_%d", index), pair.key)
		env = setEnv(env, fmt.Sprintf("GIT_CONFIG_VALUE_%d", index), pair.value)
	}

	return env
}

func ensureEmptyGitConfigPath() string {
	configPath := filepath.Join(os.TempDir(), "openclaw-empty.gitconfig")
	if _, err := os.Stat(configPath); err == nil {
		return configPath
	}
	if err := os.WriteFile(configPath, []byte(""), 0o644); err != nil {
		return ""
	}
	return configPath
}

func rewriteGitHubURLToMirror(raw string) string {
	replacements := []struct {
		source string
		target string
	}{
		{source: "https://github.com/", target: gitHubCloneMirrorPrefix()},
		{source: "ssh://git@github.com/", target: gitHubCloneMirrorPrefix()},
		{source: "git@github.com:", target: gitHubCloneMirrorPrefix()},
		{source: "git+ssh://git@github.com/", target: gitHubCloneMirrorPrefix()},
	}

	for _, replacement := range replacements {
		if strings.HasPrefix(raw, replacement.source) {
			return replacement.target + strings.TrimPrefix(raw, replacement.source)
		}
	}

	return raw
}

func gitHubCloneMirrorPrefix() string {
	return "https://gitclone.com/github.com/"
}

func containsString(items []string, target string) bool {
	for _, item := range items {
		if item == target {
			return true
		}
	}
	return false
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
	path := name
	if strings.ContainsAny(name, `\/`) {
		if _, err := os.Stat(name); err != nil {
			return fmt.Errorf("command not found: %s", name)
		}
	} else {
		var err error
		path, err = exec.LookPath(name)
		if err != nil {
			return fmt.Errorf("command not found: %s", name)
		}
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
	var outputTail commandOutputTail
	wg.Add(2)

	// Use decodeOutputBytes (defined in executor.go) to handle GBK / GB18030
	// output from Windows system tools (schtasks, winget, etc.).
	go func() {
		defer wg.Done()
		scanner := bufio.NewScanner(stdout)
		for scanner.Scan() {
			line := sanitizeOutput(decodeOutputBytes(scanner.Bytes()))
			if line != "" {
				outputTail.Add(line)
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
				outputTail.Add(line)
				emit(step, "running", line)
			}
		}
	}()

	wg.Wait()

	if err := cmd.Wait(); err != nil {
		return newCommandRunError(name, err, outputTail.Lines())
	}

	return nil
}

type commandOutputTail struct {
	mu    sync.Mutex
	lines []string
}

func (t *commandOutputTail) Add(line string) {
	t.mu.Lock()
	defer t.mu.Unlock()

	t.lines = append(t.lines, line)
	const maxLines = 12
	if len(t.lines) > maxLines {
		t.lines = append([]string(nil), t.lines[len(t.lines)-maxLines:]...)
	}
}

func (t *commandOutputTail) Lines() []string {
	t.mu.Lock()
	defer t.mu.Unlock()
	return append([]string(nil), t.lines...)
}

type commandRunError struct {
	name      string
	exitCode  string
	diagnosis string
	lastLines []string
	err       error
}

func newCommandRunError(name string, err error, lastLines []string) error {
	exitCode, diagnosis := describeProcessExit(err)
	return &commandRunError{
		name:      name,
		exitCode:  exitCode,
		diagnosis: diagnosis,
		lastLines: lastLines,
		err:       err,
	}
}

func (e *commandRunError) Error() string {
	message := fmt.Sprintf("%s exited with error", e.name)
	if e.exitCode != "" {
		message += fmt.Sprintf(" (%s)", e.exitCode)
	}
	if e.diagnosis != "" {
		message += ": " + e.diagnosis
	} else if e.err != nil {
		message += fmt.Sprintf(": %v", e.err)
	}
	if len(e.lastLines) > 0 {
		message += "\nLast output:\n" + strings.Join(e.lastLines, "\n")
	}
	return message
}

func (e *commandRunError) Unwrap() error {
	return e.err
}

func describeProcessExit(err error) (string, string) {
	var exitErr *exec.ExitError
	if !errors.As(err, &exitErr) {
		return "", ""
	}

	code := exitErr.ExitCode()
	if code < 0 {
		return fmt.Sprintf("exit code %d", code), ""
	}

	if runtime.GOOS != "windows" {
		return fmt.Sprintf("exit code %d", code), ""
	}

	unsignedCode := uint32(code)
	signedCode := int32(unsignedCode)
	exitCode := fmt.Sprintf("exit code %d (0x%08x)", signedCode, unsignedCode)

	switch signedCode {
	case -4058:
		return exitCode, "ENOENT: a required file or command was not found. On Windows this usually means npm tried to run a missing executable such as cmd.exe or git."
	case -1073741819:
		return exitCode, "native process crashed with an access violation. This usually comes from a native dependency or driver/runtime incompatibility on Windows."
	case -1073741510:
		return exitCode, "the process was interrupted or forcibly terminated"
	case -1073741502:
		return exitCode, "a required DLL failed to initialize"
	default:
		return exitCode, ""
	}
}

func parseSemverParts(raw string) (semverParts, bool) {
	matches := nodeVersionRe.FindStringSubmatch(strings.TrimSpace(raw))
	if len(matches) != 4 {
		return semverParts{}, false
	}

	major, err := strconv.Atoi(matches[1])
	if err != nil {
		return semverParts{}, false
	}
	minor, err := strconv.Atoi(matches[2])
	if err != nil {
		return semverParts{}, false
	}
	patch, err := strconv.Atoi(matches[3])
	if err != nil {
		return semverParts{}, false
	}

	return semverParts{Major: major, Minor: minor, Patch: patch}, true
}

func compareSemverParts(left, right semverParts) int {
	switch {
	case left.Major != right.Major:
		return left.Major - right.Major
	case left.Minor != right.Minor:
		return left.Minor - right.Minor
	default:
		return left.Patch - right.Patch
	}
}

func formatSemverParts(version semverParts) string {
	return fmt.Sprintf("%d.%d.%d", version.Major, version.Minor, version.Patch)
}

func resolveWindowsCmdPath() string {
	if runtime.GOOS != "windows" {
		return ""
	}

	candidates := []string{
		os.Getenv("ComSpec"),
		filepath.Join(os.Getenv("SystemRoot"), "System32", "cmd.exe"),
		filepath.Join(os.Getenv("WINDIR"), "System32", "cmd.exe"),
	}

	for _, candidate := range candidates {
		if candidate == "" {
			continue
		}
		if _, err := os.Stat(candidate); err == nil {
			return candidate
		}
	}

	if path, err := exec.LookPath("cmd.exe"); err == nil {
		return path
	}

	return "cmd.exe"
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
