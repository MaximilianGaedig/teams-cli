package auth_test

import (
	"context"
	"encoding/json"
	"os"
	"path/filepath"
	"testing"

	"github.com/MaximilianGaedig/teams"
	"teams-cli/lib/auth"
	"teams-cli/lib/m365"
)

func loadConfigForTest(t *testing.T) *m365.Client {
	home, err := os.UserHomeDir()
	if err != nil {
		t.Fatalf("Cannot determine home directory: %v", err)
	}
	path := filepath.Join(home, ".config", "teams-cli", "config.json")
	data, err := os.ReadFile(path)
	if err != nil {
		t.Fatalf("No config file: %v", err)
	}
	var cfg teams.Config
	if err := json.Unmarshal(data, &cfg); err != nil {
		t.Fatalf("Cannot parse config: %v", err)
	}

	client := m365.NewClientFromConfig(&cfg)
	if !client.IsAuthenticated() {
		t.Fatal("Not authenticated")
	}
	return client
}

func TestService_StartDeviceLogin_API(t *testing.T) {
	client := loadConfigForTest(t)
	ctx := context.Background()
	svc := auth.NewService()

	dc, err := svc.StartDeviceLogin(ctx, client)
	if err != nil {
		t.Fatalf("StartDeviceLogin failed: %v", err)
	}

	if dc.UserCode == "" {
		t.Fatal("UserCode should not be empty")
	}
	if dc.DeviceCode == "" {
		t.Fatal("DeviceCode should not be empty")
	}
	if dc.Message == "" {
		t.Fatal("Message should not be empty")
	}

	t.Logf("Device login initiated: user code = %s", dc.UserCode)
}

func TestService_DiscoverSharePointTenant_API(t *testing.T) {
	client := loadConfigForTest(t)
	ctx := context.Background()
	svc := auth.NewService()

	tenant, err := svc.DiscoverSharePointTenant(ctx, client)
	if err != nil {
		t.Fatalf("DiscoverSharePointTenant failed: %v", err)
	}

	if tenant == "" {
		t.Fatal("Tenant should not be empty")
	}

	t.Logf("Discovered SharePoint tenant: %s", tenant)
}

func TestService_GetSharePointTenantFromTeamsAPI(t *testing.T) {
	client := loadConfigForTest(t)
	ctx := context.Background()
	svc := auth.NewService()

	tenant, err := svc.DiscoverSharePointTenant(ctx, client)
	if err != nil {
		t.Fatalf("DiscoverSharePointTenant failed: %v", err)
	}

	t.Logf("Tenant from Teams API: %s", tenant)
}
