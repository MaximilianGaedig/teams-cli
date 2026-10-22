// Package testenv keeps unit tests away from the user's real teams-cli state.
package testenv

import (
	"bytes"
	"fmt"
	"os"
	"path/filepath"
	"testing"
)

// guardedFiles are the files under the real HOME that tests must never change.
var guardedFiles = []string{
	filepath.Join(".config", "teams-cli", "credentials.json"),
	filepath.Join(".config", "teams-cli", "config.json"),
}

// Run runs the tests with HOME pointed at a throwaway directory, so nothing
// can read or overwrite the real login. Afterwards it re-checks the real files
// and fails the run if any of them changed, which catches a test that builds
// its path without going through HOME. Use it from TestMain:
//
//	func TestMain(m *testing.M) { os.Exit(testenv.Run(m)) }
func Run(m *testing.M) int {
	realHome, _ := os.UserHomeDir()
	before := snapshot(realHome)

	home, err := os.MkdirTemp("", "teams-cli-test-home-")
	if err != nil {
		fmt.Fprintln(os.Stderr, "testenv: cannot create test HOME:", err)
		return 1
	}
	defer os.RemoveAll(home) //nolint:errcheck

	_ = os.Setenv("HOME", home)
	_ = os.Unsetenv("XDG_CONFIG_HOME")
	code := m.Run()
	_ = os.Setenv("HOME", realHome)

	if realHome != "" {
		after := snapshot(realHome)
		for _, f := range guardedFiles {
			if !bytes.Equal(before[f], after[f]) {
				fmt.Fprintf(os.Stderr, "testenv: tests modified the real %s\n", filepath.Join(realHome, f))
				code = 1
			}
		}
	}
	return code
}

func snapshot(home string) map[string][]byte {
	s := make(map[string][]byte, len(guardedFiles))
	if home == "" {
		return s
	}
	for _, f := range guardedFiles {
		data, err := os.ReadFile(filepath.Join(home, f))
		if err != nil {
			data = nil
		}
		s[f] = data
	}
	return s
}
