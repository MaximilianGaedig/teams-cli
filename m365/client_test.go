package m365

import (
	"os"
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
	t.Setenv("HOME", t.TempDir())
	opts := []ClientOption{
		WithVerbose(true),
	}
	client := NewClient("client-id", "tenant-id", "sharepoint-tenant", opts...)
	if !client.verbose {
		t.Error("verbose should be true after WithVerbose(true)")
	}
}

func TestClient_NewClientFromConfig(t *testing.T) {
	t.Setenv("HOME", t.TempDir())
	client := NewClient("client-id", "tenant-id", "sharepoint-tenant")
	if client == nil {
		t.Error("NewClient returned nil")
	}
}

func TestClient_SaveCredentials_AtomicRoundTrip(t *testing.T) {
	t.Setenv("HOME", t.TempDir())
	dir := t.TempDir()
	client := NewClient("client-id", "tenant-id", "sharepoint-tenant", WithCredentialsDir(dir))

	for _, rt := range []string{"first-token", "second-token"} {
		if err := client.SaveCredentials(rt); err != nil {
			t.Fatalf("SaveCredentials(%q): %v", rt, err)
		}
	}

	entries, err := os.ReadDir(dir)
	if err != nil {
		t.Fatal(err)
	}
	if len(entries) != 1 || entries[0].Name() != "credentials.json" {
		var names []string
		for _, e := range entries {
			names = append(names, e.Name())
		}
		t.Fatalf("credentials dir = %v, want only credentials.json", names)
	}
	info, err := entries[0].Info()
	if err != nil {
		t.Fatal(err)
	}
	if perm := info.Mode().Perm(); perm != 0600 {
		t.Errorf("credentials.json mode = %o, want 600", perm)
	}

	reloaded := NewClient("client-id", "tenant-id", "sharepoint-tenant", WithCredentialsDir(dir))
	if got := reloaded.GetRefreshToken(); got != "second-token" {
		t.Errorf("reloaded refresh token = %q, want %q", got, "second-token")
	}
}
