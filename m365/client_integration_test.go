//go:build integration

package m365_test

import (
	"context"
	"encoding/json"
	"os"
	"path/filepath"
	"testing"

	"github.com/MaximilianGaedig/teams"
	"github.com/MaximilianGaedig/teams/m365"
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

func TestClient_AcquireAccessToken_API(t *testing.T) {
	client := loadConfigForTest(t)
	ctx := context.Background()

	token, err := client.AcquireAccessToken(ctx, m365.GraphResource)
	if err != nil {
		t.Fatalf("AcquireAccessToken failed: %v", err)
	}

	if token == "" {
		t.Fatal("Token should not be empty")
	}

	t.Logf("Acquired token: %s...", token[:50])
}

func TestClient_ValidateAndRefreshToken_API(t *testing.T) {
	client := loadConfigForTest(t)
	ctx := context.Background()

	err := client.ValidateAndRefreshToken(ctx)
	if err != nil {
		t.Fatalf("ValidateAndRefreshToken failed: %v", err)
	}

	if !client.IsTokenValid() {
		t.Error("Token should be valid after refresh")
	}
}

func TestClient_GraphRequest_API(t *testing.T) {
	client := loadConfigForTest(t)
	ctx := context.Background()

	data, err := client.GraphRequest(ctx, "GET", "me", nil)
	if err != nil {
		t.Fatalf("GraphRequest failed: %v", err)
	}

	if len(data) == 0 {
		t.Error("Response should not be empty")
	}

	t.Logf("Got user data: %s...", string(data[:100]))
}

func TestClient_AcquireSharePointCookie_API(t *testing.T) {
	client := loadConfigForTest(t)
	ctx := context.Background()

	cookie, expiry, err := client.AcquireSharePointCookie(ctx)
	if err != nil {
		t.Fatalf("AcquireSharePointCookie failed: %v", err)
	}

	if cookie == "" {
		t.Fatal("Cookie should not be empty")
	}

	if expiry.IsZero() {
		t.Error("Expiry should not be zero")
	}

	t.Logf("Got SharePoint cookie, expires at %v", expiry)
}

func TestClient_GetGraphToken_API(t *testing.T) {
	client := loadConfigForTest(t)

	token, err := client.GetGraphToken()
	if err != nil {
		t.Fatalf("GetGraphToken failed: %v", err)
	}

	if token == "" {
		t.Fatal("Token should not be empty")
	}

	t.Logf("Got graph token: %s...", token[:50])
}

func TestClient_GetSharePointToken_API(t *testing.T) {
	client := loadConfigForTest(t)

	token, err := client.GetSharePointToken()
	if err != nil {
		t.Fatalf("GetSharePointToken failed: %v", err)
	}

	if token == "" {
		t.Fatal("Token should not be empty")
	}

	t.Logf("Got SharePoint token: %s...", token[:50])
}

func TestClient_SharePointRequest_API(t *testing.T) {
	client := loadConfigForTest(t)
	ctx := context.Background()

	data, err := client.SharePointRequest(ctx, "GET", "_api/web?$select=Title", nil)
	if err != nil {
		t.Fatalf("SharePointRequest failed: %v", err)
	}

	if len(data) == 0 {
		t.Error("Response should not be empty")
	}

	t.Logf("Got SharePoint data: %s...", string(data[:100]))
}

func TestClient_IsAuthenticated_API(t *testing.T) {
	client := loadConfigForTest(t)

	if !client.IsAuthenticated() {
		t.Error("Client should be authenticated")
	}
}

func TestClient_IsTokenValid_API(t *testing.T) {
	client := loadConfigForTest(t)

	if !client.IsTokenValid() {
		t.Error("Token should be valid")
	}
}

func TestClient_HTTPClient_API(t *testing.T) {
	client := loadConfigForTest(t)

	httpClient := client.HTTPClient()
	if httpClient == nil {
		t.Error("HTTPClient should not be nil")
	}
}

func TestClient_GetDeviceCodeURL_API(t *testing.T) {
	client := loadConfigForTest(t)

	url := client.GetDeviceCodeURL()
	if url == "" {
		t.Error("DeviceCodeURL should not be empty")
	}

	t.Logf("Device code URL: %s", url)
}

func TestClient_GetSharePointURL_API(t *testing.T) {
	client := loadConfigForTest(t)

	url := client.GetSharePointURL()
	if url == "" {
		t.Error("SharePointURL should not be empty")
	}

	t.Logf("SharePoint URL: %s", url)
}

func TestClient_GetSharePointMyURL_API(t *testing.T) {
	client := loadConfigForTest(t)

	url := client.GetSharePointMyURL()
	if url == "" {
		t.Error("SharePointMyURL should not be empty")
	}

	t.Logf("SharePoint My URL: %s", url)
}

func TestClient_RawRequest_Graph_API(t *testing.T) {
	client := loadConfigForTest(t)
	ctx := context.Background()

	resp, err := client.RawRequest(ctx, m365.APITypeGraph, "GET", "me", nil)
	if err != nil {
		t.Fatalf("RawRequest failed: %v", err)
	}

	if resp == nil {
		t.Fatal("Response should not be nil")
	}

	body, err := resp.GetBody()
	if err != nil {
		t.Fatalf("GetBody failed: %v", err)
	}

	if len(body) == 0 {
		t.Error("Response body should not be empty")
	}

	t.Logf("Got raw response: %s...", string(body[:100]))
}

func TestClient_RawRequest_SharePoint_API(t *testing.T) {
	client := loadConfigForTest(t)
	ctx := context.Background()

	resp, err := client.RawRequest(ctx, m365.APITypeSharePoint, "GET", "_api/web?$select=Title", nil)
	if err != nil {
		t.Fatalf("RawRequest failed: %v", err)
	}

	if resp == nil {
		t.Fatal("Response should not be nil")
	}

	body, err := resp.GetBody()
	if err != nil {
		t.Fatalf("GetBody failed: %v", err)
	}

	if len(body) == 0 {
		t.Error("Response body should not be empty")
	}

	t.Logf("Got SharePoint raw response: %s...", string(body[:100]))
}
