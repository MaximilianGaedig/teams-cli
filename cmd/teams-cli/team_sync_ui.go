package main

import (
	"context"
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"
	"strings"

	"github.com/MaximilianGaedig/teams"
	"github.com/MaximilianGaedig/teams/m365"
	"github.com/MaximilianGaedig/teams/sharepoint"
	"github.com/MaximilianGaedig/teams/teamservice"
	"github.com/charmbracelet/lipgloss"
	"github.com/spf13/cobra"
)

var (
	spinnerStyle = lipgloss.NewStyle().Foreground(lipgloss.Color("#7D56F4"))
	helpStyle    = lipgloss.NewStyle().Foreground(lipgloss.Color("#626262")).Render
)

type syncStats struct {
	totalFiles   int
	totalFolders int
	totalErrors  int
}

func addTeamSyncCommand(cmd *cobra.Command) {
	syncCmd := &cobra.Command{
		Use:   "sync [team-name] [directory]",
		Short: "Sync entire team files to local directory",
		Long: `Synchronizes all files from a Microsoft Teams team to a local directory.

This downloads files from all channels' SharePoint folders, maintaining
the complete channel and folder structure locally.

Examples:
  teams-cli teams sync "Marketing" ./marketing-backup
  teams-cli teams sync "Engineering" ./engineering-files`,
		Args: cobra.RangeArgs(1, 2),
		Run: func(cmd *cobra.Command, args []string) {
			teamName := args[0]
			targetDir := "."
			if len(args) > 1 {
				targetDir = args[1]
			}

			maxFileSizeMB, _ := cmd.Flags().GetInt64("max-size-mb")
			maxFileSize := maxFileSizeMB * 1024 * 1024 // Convert MB to bytes

			cfg := loadConfig()
			if cfg == nil {
				fmt.Fprintln(os.Stderr, "Not authenticated. Run 'teams-cli auth login' first.")
				os.Exit(1)
			}

			fmt.Printf("Syncing team '%s' to: %s\n\n", teamName, targetDir)

			m365Client := m365.NewClientFromConfig(cfg, m365.WithVerbose(verbose))
			teamSvc := teamservice.NewService()
			spSvc := sharepoint.NewService()
			ctx := context.Background()

			// Find the team
			teamList, err := teamSvc.ListTeams(ctx, m365Client)
			if err != nil {
				fmt.Fprintf(os.Stderr, "Error listing teams: %v\n", err)
				os.Exit(1)
			}

			var targetTeam *teams.Team
			for _, t := range teamList {
				if strings.EqualFold(t.DisplayName, teamName) || t.ID == teamName {
					targetTeam = &t
					break
				}
			}

			if targetTeam == nil {
				fmt.Fprintf(os.Stderr, "Team not found: %s\n", teamName)
				os.Exit(1)
			}

			fmt.Printf("Found team: %s (ID: %s)\n", targetTeam.DisplayName, targetTeam.ID)

			// Get channels
			channels, err := teamSvc.ListChannels(ctx, m365Client, targetTeam.ID)
			if err != nil {
				fmt.Fprintf(os.Stderr, "Error listing channels: %v\n", err)
				os.Exit(1)
			}

			fmt.Printf("Found %d channels\n\n", len(channels))

			// Create target directory
			if err := os.MkdirAll(targetDir, 0755); err != nil {
				fmt.Fprintf(os.Stderr, "Error creating directory: %v\n", err)
				os.Exit(1)
			}

			stats := &syncStats{}

			// Sync each channel's files
			for _, channel := range channels {
				channelDir := filepath.Join(targetDir, sanitizeFilename(channel.DisplayName))

				// Get the files folder for this channel
				folder, err := teamSvc.GetFilesFolder(ctx, m365Client, targetTeam.ID, channel.ID)
				if err != nil {
					// Some channels might not have files folders (private channels, etc.)
					fmt.Printf("Channel '%s': No files folder (skipping)\n", channel.DisplayName)
					continue
				}

				fmt.Printf("Channel '%s': Syncing files...\n", channel.DisplayName)

				// Get site ID from folder URL
				siteID, err := extractSiteIDFromURL(ctx, m365Client, folder.WebUrl)
				if err != nil {
					fmt.Printf("  Warning: Could not extract site ID from URL: %v\n", err)
					continue
				}

				// Recursively sync all files and folders from this channel
				if err := syncFolderRecursive(ctx, spSvc, m365Client, siteID, "", channelDir, stats, maxFileSize); err != nil {
					fmt.Fprintf(os.Stderr, "  Error syncing channel '%s': %v\n", channel.DisplayName, err)
					stats.totalErrors++
					continue
				}

				fmt.Printf("  Synced to: %s\n", channelDir)
			}

			fmt.Printf("\n✓ Sync completed!\n")
			fmt.Printf("  Total folders: %d\n", stats.totalFolders)
			fmt.Printf("  Total files: %d\n", stats.totalFiles)
			if stats.totalErrors > 0 {
				fmt.Printf("  Errors: %d\n", stats.totalErrors)
			}
		},
	}

	syncCmd.Flags().Int64("max-size-mb", 0, "Maximum file size to sync in MB (0 = no limit)")

	cmd.AddCommand(syncCmd)
}

// syncFolderRecursive recursively syncs all files and folders from SharePoint to local
func syncFolderRecursive(ctx context.Context, spSvc *sharepoint.Service, client teams.Client, siteID, remotePath, localPath string, stats *syncStats, maxFileSize int64) error {
	// Create local directory
	if err := os.MkdirAll(localPath, 0755); err != nil {
		return fmt.Errorf("failed to create directory %s: %w", localPath, err)
	}

	// List all items in this folder
	items, err := spSvc.ListFiles(ctx, client, siteID, remotePath)
	if err != nil {
		return fmt.Errorf("failed to list files in %s: %w", remotePath, err)
	}

	for _, item := range items {
		localItemPath := filepath.Join(localPath, sanitizeFilename(item.Name))

		if item.IsFolder {
			// Recursively sync subdirectory
			subRemotePath := remotePath
			if subRemotePath != "" {
				subRemotePath = subRemotePath + "/" + item.Name
			} else {
				subRemotePath = item.Name
			}

			stats.totalFolders++
			if err := syncFolderRecursive(ctx, spSvc, client, siteID, subRemotePath, localItemPath, stats, maxFileSize); err != nil {
				return err
			}
		} else {
			// Check file size limit
			if maxFileSize > 0 && item.Size > maxFileSize {
				fmt.Printf("    Skipping %s (size %d exceeds limit %d)\n", item.Name, item.Size, maxFileSize)
				continue
			}

			// Download file
			data, err := spSvc.DownloadFile(ctx, client, siteID, item.ID)
			if err != nil {
				fmt.Fprintf(os.Stderr, "    Error downloading %s: %v\n", item.Name, err)
				stats.totalErrors++
				continue
			}

			if err := os.WriteFile(localItemPath, data, 0644); err != nil {
				fmt.Fprintf(os.Stderr, "    Error saving %s: %v\n", item.Name, err)
				stats.totalErrors++
				continue
			}

			stats.totalFiles++
		}
	}

	return nil
}

func sanitizeFilename(name string) string {
	replacer := strings.NewReplacer(
		"/", "-",
		"\\", "-",
		":", "-",
		"*", "-",
		"?", "-",
		"\"", "-",
		"<", "-",
		">", "-",
		"|", "-",
	)
	return replacer.Replace(name)
}

func extractSiteIDFromURL(ctx context.Context, client teams.Client, webURL string) (string, error) {
	// SharePoint URLs look like:
	// https://tenant.sharepoint.com/sites/site-name/Shared Documents/Channel
	// We need to extract the site URL part and look it up via Graph API

	// Parse the URL to get the site path
	// Format: https://hostname/sites/site-name/...
	parts := strings.Split(webURL, "/")
	if len(parts) < 5 {
		return "", fmt.Errorf("invalid SharePoint URL format: %s", webURL)
	}

	// Find the "sites" part
	siteIdx := -1
	for i, part := range parts {
		if strings.ToLower(part) == "sites" && i+1 < len(parts) {
			siteIdx = i
			break
		}
	}

	if siteIdx == -1 {
		return "", fmt.Errorf("could not find 'sites' in URL: %s", webURL)
	}

	// Extract hostname and site path
	hostname := parts[2]                                    // e.g., tenant.sharepoint.com
	sitePath := strings.Join(parts[siteIdx:siteIdx+2], "/") // e.g., sites/site-name

	// Call Graph API to get site by path
	endpoint := fmt.Sprintf("sites/%s:/%s", hostname, sitePath)
	data, err := client.GraphRequest(ctx, "GET", endpoint, nil)
	if err != nil {
		return "", fmt.Errorf("failed to lookup site by URL: %w", err)
	}

	var resp struct {
		ID string `json:"id"`
	}
	if err := json.Unmarshal(data, &resp); err != nil {
		return "", fmt.Errorf("failed to parse site response: %w", err)
	}

	return resp.ID, nil
}
