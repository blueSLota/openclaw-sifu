//go:build windows

package main

import (
	"context"
	"fmt"
	"os"
	"os/exec"
	"strings"
	"time"
)

func (a *App) RunPostInstallActions() PostInstallActionResult {
	openclawPath := findOpenClawCommand()
	if openclawPath == "" {
		return PostInstallActionResult{
			Message: "未找到 OpenClaw 命令，无法执行安装后的系统集成。",
			Error:   "openclaw command not found",
		}
	}

	if err := runElevatedGatewayInstall(openclawPath); err != nil {
		if isElevationCancelled(err) {
			return PostInstallActionResult{
				Message:   "您取消了管理员授权，后台服务和计划任务未启用。",
				Error:     err.Error(),
				Cancelled: true,
			}
		}

		return PostInstallActionResult{
			Message: "管理员授权未完成，后台服务和计划任务未启用。",
			Error:   err.Error(),
		}
	}

	return PostInstallActionResult{
		Success: true,
		Message: "已启用 OpenClaw 后台服务和相关计划任务。",
	}
}

func runElevatedGatewayInstall(openclawPath string) error {
	scriptPath, err := writeBatchScript("openclaw-postinstall-*.cmd", []string{
		"@echo off",
		"setlocal",
		fmt.Sprintf("call \"%s\" config set gateway.mode local >nul 2>&1", openclawPath),
		"if errorlevel 1 exit /b %errorlevel%",
		fmt.Sprintf("call \"%s\" gateway install --force", openclawPath),
		"if errorlevel 1 exit /b %errorlevel%",
		fmt.Sprintf("call \"%s\" gateway start", openclawPath),
		"if errorlevel 1 exit /b %errorlevel%",
		"exit /b 0",
		"",
	})
	if err != nil {
		return fmt.Errorf("create post-install helper: %w", err)
	}
	defer os.Remove(scriptPath)

	return runElevatedBatchScript(scriptPath, 2*time.Minute, "post-install action timed out while waiting for OpenClaw to finish configuring system integration")
}

func removeOpenClawScheduledTask() error {
	scriptPath, err := writeBatchScript("openclaw-task-cleanup-*.cmd", []string{
		"@echo off",
		"setlocal",
		"schtasks /Query /TN \"OpenClaw Gateway\" >nul 2>&1",
		"if errorlevel 1 exit /b 0",
		"schtasks /Delete /TN \"OpenClaw Gateway\" /F",
		"if errorlevel 1 exit /b %errorlevel%",
		"exit /b 0",
		"",
	})
	if err != nil {
		return fmt.Errorf("create task cleanup helper: %w", err)
	}
	defer os.Remove(scriptPath)

	return runElevatedBatchScript(scriptPath, time.Minute, "scheduled-task cleanup timed out")
}

func runElevatedBatchScript(scriptPath string, timeout time.Duration, timeoutMessage string) error {

	shellPath := detectPowerShell()
	if shellPath == "not-found" {
		shellPath = "powershell.exe"
	}

	psScript := buildElevationScript(scriptPath)
	ctx, cancel := context.WithTimeout(context.Background(), 2*time.Minute)
	defer cancel()

	cmd := exec.CommandContext(ctx, shellPath, "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-Command", psScript)
	output, err := cmd.CombinedOutput()
	if ctx.Err() == context.DeadlineExceeded {
		return fmt.Errorf("%s", timeoutMessage)
	}
	if err == nil {
		return nil
	}

	message := sanitizeOutput(decodeOutputBytes(output))
	if message == "" {
		return err
	}

	return fmt.Errorf("%s", message)
}

func writeBatchScript(pattern string, lines []string) (string, error) {
	file, err := os.CreateTemp("", pattern)
	if err != nil {
		return "", err
	}
	defer file.Close()

	content := strings.Join(lines, "\r\n")
	if _, err := file.WriteString(content); err != nil {
		return "", err
	}

	return file.Name(), nil
}

func buildElevationScript(scriptPath string) string {
	quotedPath := strings.ReplaceAll(scriptPath, "'", "''")
	return strings.TrimSpace(fmt.Sprintf(`
$scriptPath = '%s'
try {
    $proc = Start-Process -FilePath 'cmd.exe' -ArgumentList @('/d', '/c', $scriptPath) -Verb RunAs -WindowStyle Hidden -Wait -PassThru -ErrorAction Stop
    exit $proc.ExitCode
} catch {
    Write-Error $_.Exception.Message
    exit 1
}
`, quotedPath))
}

func isElevationCancelled(err error) bool {
	if err == nil {
		return false
	}

	message := strings.ToLower(err.Error())
	return strings.Contains(message, "cancel") ||
		strings.Contains(message, "取消") ||
		strings.Contains(message, "拒绝")
}
