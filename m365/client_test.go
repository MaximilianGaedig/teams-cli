package m365

import (
	"testing"
	"time"

	"github.com/MaximilianGaedig/teams"
)

func TestTokenInfo(t *testing.T) {
	info := teams.TokenInfo{
		AccessToken: "token",
		ExpiresAt:   time.Now().Add(time.Hour),
	}
	if info.AccessToken != "token" {
		t.Error("AccessToken not set correctly")
	}
	if info.ExpiresAt.IsZero() {
		t.Error("ExpiresAt should not be zero")
	}
}

func TestWithVerbose(t *testing.T) {
	opts := []ClientOption{
		WithVerbose(true),
	}
	client := NewClient("client-id", "tenant-id", "sharepoint-tenant", opts...)
	if !client.verbose {
		t.Error("verbose should be true after WithVerbose(true)")
	}
}

func TestClient_NewClientFromConfig(t *testing.T) {
	client := NewClient("client-id", "tenant-id", "sharepoint-tenant")
	if client == nil {
		t.Error("NewClient returned nil")
	}
}
