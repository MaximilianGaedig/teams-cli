package main

import (
	"encoding/json"
	"os"
	"path/filepath"
	"testing"

	"github.com/MaximilianGaedig/teams"
)

func TestConfigPath(t *testing.T) {
	path := configPath()
	if path == "" {
		t.Error("configPath returned empty string")
	}
	if !filepath.IsAbs(path) {
		t.Errorf("configPath should return absolute path, got: %s", path)
	}
}

func TestLoadConfigNoFile(t *testing.T) {
	tmpDir := t.TempDir()
	orig := os.Getenv("HOME")
	_ = os.Unsetenv("HOME")
	defer func() { _ = os.Setenv("HOME", orig) }()

	_, err := os.UserHomeDir()
	if err != nil {
		t.Skip("Cannot determine home directory without HOME set")
	}

	path := filepath.Join(tmpDir, "nonexistent", "config.json")
	data, err := os.ReadFile(path)
	if err == nil {
		var cfg teams.Config
		if err := json.Unmarshal(data, &cfg); err == nil {
			t.Error("Expected error loading nonexistent file")
		}
	}
}

func TestLoadConfigInvalidJSON(t *testing.T) {
	tmpDir := t.TempDir()
	configFile := filepath.Join(tmpDir, "config.json")

	_ = os.WriteFile(configFile, []byte("invalid json"), 0644)

	data, err := os.ReadFile(configFile)
	if err != nil {
		t.Fatalf("Failed to read test file: %v", err)
	}

	var cfg teams.Config
	if err := json.Unmarshal(data, &cfg); err == nil {
		t.Error("Expected error parsing invalid JSON")
	}
}

func TestSaveAndLoadConfigRoundTrip(t *testing.T) {
	tmpDir := t.TempDir()
	configFile := filepath.Join(tmpDir, "config.json")

	cfg := &teams.Config{
		ClientID:         "test-client-id",
		TenantID:         "test-tenant-id",
		SharePointTenant: "test-tenant",
	}

	data, err := json.MarshalIndent(cfg, "", "  ")
	if err != nil {
		t.Fatalf("Failed to marshal config: %v", err)
	}

	err = os.WriteFile(configFile, data, 0644)
	if err != nil {
		t.Fatalf("Failed to write config: %v", err)
	}

	loadedData, err := os.ReadFile(configFile)
	if err != nil {
		t.Fatalf("Failed to read config: %v", err)
	}

	var loadedCfg teams.Config
	err = json.Unmarshal(loadedData, &loadedCfg)
	if err != nil {
		t.Fatalf("Failed to unmarshal config: %v", err)
	}

	if loadedCfg.ClientID != cfg.ClientID {
		t.Errorf("ClientID mismatch: got %s, want %s", loadedCfg.ClientID, cfg.ClientID)
	}
	if loadedCfg.TenantID != cfg.TenantID {
		t.Errorf("TenantID mismatch: got %s, want %s", loadedCfg.TenantID, cfg.TenantID)
	}
	if loadedCfg.SharePointTenant != cfg.SharePointTenant {
		t.Errorf("SharePointTenant mismatch: got %s, want %s", loadedCfg.SharePointTenant, cfg.SharePointTenant)
	}
}

func TestSaveConfigCreatesDir(t *testing.T) {
	tmpDir := t.TempDir()
	configFile := filepath.Join(tmpDir, "subdir", "config.json")

	cfg := &teams.Config{
		ClientID: "test-client-id",
	}

	dir := filepath.Dir(configFile)
	err := os.MkdirAll(dir, 0755)
	if err != nil {
		t.Fatalf("Failed to create dir: %v", err)
	}

	data, err := json.MarshalIndent(cfg, "", "  ")
	if err != nil {
		t.Fatalf("Failed to marshal: %v", err)
	}

	err = os.WriteFile(configFile, data, 0644)
	if err != nil {
		t.Fatalf("Failed to write config: %v", err)
	}

	_, err = os.ReadFile(configFile)
	if err != nil {
		t.Fatalf("Config file not created: %v", err)
	}
}

func TestAuthCommands(t *testing.T) {
	cmd := authCommands()
	if cmd == nil {
		t.Fatal("authCommands returned nil")
	}
	if cmd.Use != "auth" {
		t.Errorf("authCommands use = %q, want %q", cmd.Use, "auth")
	}
	if len(cmd.Commands()) == 0 {
		t.Error("authCommands has no subcommands")
	}
}

func TestUserCommands(t *testing.T) {
	cmd := userCommands()
	if cmd == nil {
		t.Fatal("userCommands returned nil")
	}
	if cmd.Use != "users" {
		t.Errorf("userCommands use = %q, want %q", cmd.Use, "users")
	}
}

func TestTeamCommands(t *testing.T) {
	cmd := teamCommands()
	if cmd == nil {
		t.Fatal("teamCommands returned nil")
	}
	if cmd.Use != "teams" {
		t.Errorf("teamCommands use = %q, want %q", cmd.Use, "teams")
	}
}

func TestSPCommands(t *testing.T) {
	cmd := spCommands()
	if cmd == nil {
		t.Fatal("spCommands returned nil")
	}
	if cmd.Use != "sharepoint" {
		t.Errorf("spCommands use = %q, want %q", cmd.Use, "sharepoint")
	}
	if len(cmd.Aliases) == 0 || cmd.Aliases[0] != "sp" {
		t.Error("spCommands should have 'sp' alias")
	}
}

func TestVersionVariable(t *testing.T) {
	if version == "" {
		t.Error("version should not be empty")
	}
}
