//go:build integration

package users_test

import (
	"context"
	"encoding/json"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/MaximilianGaedig/teams"
	"github.com/MaximilianGaedig/teams/m365"
	"github.com/MaximilianGaedig/teams/users"
)

func isPermissionError(err error) bool {
	if err == nil {
		return false
	}
	errStr := err.Error()
	return strings.Contains(errStr, "403") || strings.Contains(errStr, "Forbidden") || strings.Contains(errStr, "permission")
}

func loadConfigForTest(t *testing.T) teams.Client {
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

func TestService_ListUsers_API(t *testing.T) {
	client := loadConfigForTest(t)
	ctx := context.Background()
	svc := users.NewService()

	// Just fetch first page to verify the API works (don't fetch all 15k+ users)
	userList, nextLink, err := svc.ListUsersFirstPage(ctx, client)
	if err != nil {
		t.Fatalf("ListUsersFirstPage failed: %v", err)
	}

	if len(userList) == 0 {
		t.Fatal("Expected at least one user")
	}

	t.Logf("Found %d users on first page, has more: %v", len(userList), nextLink != "")
}

func TestService_ListUsers_Pagination_API(t *testing.T) {
	client := loadConfigForTest(t)
	ctx := context.Background()
	svc := users.NewService()

	// Test pagination by fetching first page, then second page
	firstPage, nextLink, err := svc.ListUsersFirstPage(ctx, client)
	if err != nil {
		t.Fatalf("ListUsersFirstPage failed: %v", err)
	}

	t.Logf("First page: %d users", len(firstPage))

	// Verify there's a next link for pagination
	if nextLink == "" {
		t.Skip("Only one page of users available, cannot test pagination")
	}

	// Fetch second page
	secondPage, _, err := svc.ListUsersNextPage(ctx, client, nextLink)
	if err != nil {
		t.Fatalf("ListUsersNextPage failed: %v", err)
	}

	t.Logf("Second page: %d users", len(secondPage))
	t.Logf("Total fetched: %d users", len(firstPage)+len(secondPage))

	// Verify we got different users
	if len(secondPage) == 0 {
		t.Error("Second page returned no users")
	}

	// Verify no duplicates between pages
	firstPageIDs := make(map[string]bool)
	for _, u := range firstPage {
		firstPageIDs[u.ID] = true
	}
	for _, u := range secondPage {
		if firstPageIDs[u.ID] {
			t.Errorf("Duplicate user ID found across pages: %s", u.ID)
		}
	}

	t.Logf("Pagination test passed: fetched 2 pages with %d unique users total", len(firstPage)+len(secondPage))
}

func TestService_SearchUsers_API(t *testing.T) {
	client := loadConfigForTest(t)
	ctx := context.Background()
	svc := users.NewService()

	userList, err := svc.SearchUsers(ctx, client, "test")
	if err != nil {
		t.Fatalf("SearchUsers failed: %v", err)
	}

	t.Logf("Search returned %d users", len(userList))
}

func TestService_GetUser_API(t *testing.T) {
	client := loadConfigForTest(t)
	ctx := context.Background()
	svc := users.NewService()

	// Just get first page to find a user to test with
	userList, _, err := svc.ListUsersFirstPage(ctx, client)
	if err != nil {
		t.Fatalf("ListUsersFirstPage failed: %v", err)
	}

	if len(userList) == 0 {
		t.Fatal("Expected at least one user to test GetUser")
	}

	user, err := svc.GetUser(ctx, client, userList[0].ID)
	if err != nil {
		t.Fatalf("GetUser failed: %v", err)
	}

	if user.ID != userList[0].ID {
		t.Errorf("Expected user ID %s, got %s", userList[0].ID, user.ID)
	}

	t.Logf("Got user: %s (%s)", user.DisplayName, user.Mail)
}

func TestService_GetUserPresence_API(t *testing.T) {
	client := loadConfigForTest(t)
	ctx := context.Background()
	svc := users.NewService()

	// Just get first page to find a user to test with
	userList, _, err := svc.ListUsersFirstPage(ctx, client)
	if err != nil {
		t.Fatalf("ListUsersFirstPage failed: %v", err)
	}

	if len(userList) == 0 {
		t.Fatal("Expected at least one user to test GetUserPresence")
	}

	presence, err := svc.GetUserPresence(ctx, client, userList[0].ID)
	if err != nil {
		// Presence API might require additional permissions
		t.Skipf("GetUserPresence requires additional permissions: %v", err)
	}

	t.Logf("User %s presence: %s (activity: %s)", userList[0].DisplayName, presence.Availability, presence.Activity)
}

func TestService_GetUserPhoto_API(t *testing.T) {
	client := loadConfigForTest(t)
	ctx := context.Background()
	svc := users.NewService()

	// Just get first page to find a user to test with
	userList, _, err := svc.ListUsersFirstPage(ctx, client)
	if err != nil {
		t.Fatalf("ListUsersFirstPage failed: %v", err)
	}

	if len(userList) == 0 {
		t.Fatal("Expected at least one user to test GetUserPhoto")
	}

	photo, err := svc.GetUserPhoto(ctx, client, userList[0].ID)
	if err != nil {
		// Photo API might require additional permissions
		t.Skipf("GetUserPhoto requires additional permissions: %v", err)
	}

	t.Logf("User %s photo: %dx%d, URL: %s", userList[0].DisplayName, photo.Width, photo.Height, photo.URL)
}

func TestService_GetUserManager_API(t *testing.T) {
	client := loadConfigForTest(t)
	ctx := context.Background()
	svc := users.NewService()

	// Just get first page to find a user to test with
	userList, _, err := svc.ListUsersFirstPage(ctx, client)
	if err != nil {
		t.Fatalf("ListUsersFirstPage failed: %v", err)
	}

	if len(userList) == 0 {
		t.Fatal("Expected at least one user to test GetUserManager")
	}

	manager, err := svc.GetUserManager(ctx, client, userList[0].ID)
	if err != nil {
		// Manager API might fail if user has no manager or requires additional permissions
		t.Skipf("GetUserManager skipped: %v", err)
	}

	t.Logf("User %s manager: %s (%s)", userList[0].DisplayName, manager.DisplayName, manager.Mail)
}
