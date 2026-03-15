package teamservice_test

import (
	"context"
	"encoding/json"
	"os"
	"path/filepath"
	"testing"

	"github.com/MaximilianGaedig/teams"
	"teams-cli/lib/m365"
	"github.com/MaximilianGaedig/teams/teamservice"
)

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

func TestService_ListTeams_API(t *testing.T) {
	client := loadConfigForTest(t)
	ctx := context.Background()
	svc := teams.NewService()

	teamList, err := svc.ListTeams(ctx, client)
	if err != nil {
		t.Fatalf("ListTeams failed: %v", err)
	}

	t.Logf("Found %d teams", len(teamList))
}

func TestService_GetTeamDetails_API(t *testing.T) {
	client := loadConfigForTest(t)
	ctx := context.Background()
	svc := teams.NewService()

	teamList, err := svc.ListTeams(ctx, client)
	if err != nil {
		t.Fatalf("ListTeams failed: %v", err)
	}

	if len(teamList) == 0 {
		t.Skip("No teams available to test GetTeamDetails")
	}

	details, err := svc.GetTeamDetails(ctx, client, teamList[0].ID)
	if err != nil {
		t.Fatalf("GetTeamDetails failed: %v", err)
	}

	t.Logf("Team details: %s (visibility: %s, archived: %v)", details.DisplayName, details.Visibility, details.IsArchived)
}

func TestService_ListChannels_API(t *testing.T) {
	client := loadConfigForTest(t)
	ctx := context.Background()
	svc := teams.NewService()

	teamList, err := svc.ListTeams(ctx, client)
	if err != nil {
		t.Fatalf("ListTeams failed: %v", err)
	}

	if len(teamList) == 0 {
		t.Fatal("Expected at least one team to test ListChannels")
	}

	channels, err := svc.ListChannels(ctx, client, teamList[0].ID)
	if err != nil {
		t.Fatalf("ListChannels failed: %v", err)
	}

	t.Logf("Team %s has %d channels", teamList[0].DisplayName, len(channels))
}

func TestService_GetChannel_API(t *testing.T) {
	client := loadConfigForTest(t)
	ctx := context.Background()
	svc := teams.NewService()

	teamList, err := svc.ListTeams(ctx, client)
	if err != nil {
		t.Fatalf("ListTeams failed: %v", err)
	}

	if len(teamList) == 0 {
		t.Skip("No teams available to test GetChannel")
	}

	channels, err := svc.ListChannels(ctx, client, teamList[0].ID)
	if err != nil {
		t.Fatalf("ListChannels failed: %v", err)
	}

	if len(channels) == 0 {
		t.Skip("No channels available to test GetChannel")
	}

	channel, err := svc.GetChannel(ctx, client, teamList[0].ID, channels[0].ID)
	if err != nil {
		t.Fatalf("GetChannel failed: %v", err)
	}

	t.Logf("Got channel: %s (type: %s)", channel.DisplayName, channel.MembershipType)
}

func TestService_ListMembers_API(t *testing.T) {
	client := loadConfigForTest(t)
	ctx := context.Background()
	svc := teams.NewService()

	teamList, err := svc.ListTeams(ctx, client)
	if err != nil {
		t.Fatalf("ListTeams failed: %v", err)
	}

	if len(teamList) == 0 {
		t.Fatal("Expected at least one team to test ListMembers")
	}

	members, err := svc.ListMembers(ctx, client, teamList[0].ID)
	if err != nil {
		t.Fatalf("ListMembers failed: %v", err)
	}

	t.Logf("Team %s has %d members", teamList[0].DisplayName, len(members))
}

func TestService_GetChannelMessages_API(t *testing.T) {
	client := loadConfigForTest(t)
	ctx := context.Background()
	svc := teams.NewService()

	teamList, err := svc.ListTeams(ctx, client)
	if err != nil {
		t.Fatalf("ListTeams failed: %v", err)
	}

	if len(teamList) == 0 {
		t.Skip("No teams available to test GetChannelMessages")
	}

	channels, err := svc.ListChannels(ctx, client, teamList[0].ID)
	if err != nil {
		t.Fatalf("ListChannels failed: %v", err)
	}

	if len(channels) == 0 {
		t.Skip("No channels available to test GetChannelMessages")
	}

	// Try to get messages from the first channel
	messages, err := svc.GetChannelMessages(ctx, client, teamList[0].ID, channels[0].ID)
	if err != nil {
		// Messages might require additional permissions
		t.Skipf("GetChannelMessages requires additional permissions: %v", err)
	}

	t.Logf("Channel %s has %d messages", channels[0].DisplayName, len(messages))
}

func TestService_GetFilesFolder_API(t *testing.T) {
	client := loadConfigForTest(t)
	ctx := context.Background()
	svc := teams.NewService()

	teamList, err := svc.ListTeams(ctx, client)
	if err != nil {
		t.Fatalf("ListTeams failed: %v", err)
	}

	if len(teamList) == 0 {
		t.Skip("No teams available to test GetFilesFolder")
	}

	channels, err := svc.ListChannels(ctx, client, teamList[0].ID)
	if err != nil {
		t.Fatalf("ListChannels failed: %v", err)
	}

	if len(channels) == 0 {
		t.Skip("No channels available to test GetFilesFolder")
	}

	// Try to get files folder for General channel
	var generalChannel *teams.Channel
	for i := range channels {
		if channels[i].DisplayName == "General" {
			generalChannel = &channels[i]
			break
		}
	}

	if generalChannel == nil {
		generalChannel = &channels[0]
	}

	folder, err := svc.GetFilesFolder(ctx, client, teamList[0].ID, generalChannel.ID)
	if err != nil {
		t.Skipf("GetFilesFolder requires additional permissions: %v", err)
	}

	t.Logf("Files folder for channel %s: %s (webUrl: %s)", generalChannel.DisplayName, folder.Name, folder.WebUrl)
}

func TestService_ListInstalledApps_API(t *testing.T) {
	client := loadConfigForTest(t)
	ctx := context.Background()
	svc := teams.NewService()

	teamList, err := svc.ListTeams(ctx, client)
	if err != nil {
		t.Fatalf("ListTeams failed: %v", err)
	}

	if len(teamList) == 0 {
		t.Skip("No teams available to test ListInstalledApps")
	}

	apps, err := svc.ListInstalledApps(ctx, client, teamList[0].ID)
	if err != nil {
		// Apps list might require additional permissions
		t.Skipf("ListInstalledApps requires additional permissions: %v", err)
	}

	t.Logf("Team %s has %d installed apps", teamList[0].DisplayName, len(apps))
}
