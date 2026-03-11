//go:build !windows

package main

func (a *App) RunPostInstallActions() PostInstallActionResult {
	return PostInstallActionResult{
		Message: "当前平台暂不需要额外的管理员授权步骤。",
	}
}

func removeOpenClawScheduledTask() error {
	return nil
}
