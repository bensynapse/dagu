// Copyright (C) 2026 Yota Hamada
// SPDX-License-Identifier: GPL-3.0-or-later

package engine_test

import (
	"os"
	"os/exec"
	"runtime"
	"testing"
)

func TestMain(m *testing.M) {
	warmUpPowerShell()
	os.Exit(m.Run())
}

// warmUpPowerShell runs a no-op through the shell dagu selects on Windows,
// pwsh when present and powershell otherwise, before any test does. The
// first PowerShell launch on a machine pays .NET start-up and module
// import, which on a loaded CI runner has exceeded a test's 20 second
// budget; paying it here keeps that cost out of every timed run.
func warmUpPowerShell() {
	if runtime.GOOS != "windows" {
		return
	}
	for _, shell := range []string{"pwsh", "powershell"} {
		path, err := exec.LookPath(shell)
		if err != nil {
			continue
		}
		_ = exec.Command(path, "-NoProfile", "-NonInteractive", "-Command", "exit").Run()
		return
	}
}
