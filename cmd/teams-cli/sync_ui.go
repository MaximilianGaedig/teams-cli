package main

import (
	"context"
	"fmt"
	"os"
	"time"

	"github.com/MaximilianGaedig/teams"
	"github.com/MaximilianGaedig/teams/m365"
	"github.com/MaximilianGaedig/teams/sharepoint"
	tea "github.com/charmbracelet/bubbletea"
	"github.com/spf13/cobra"
)

// syncMsg represents a progress message during sync
type syncMsg struct {
	file string
	done bool
	err  error
}

type syncProgressModel struct {
	siteName  string
	targetDir string
	direction string
	files     []string
	errors    []error
	done      bool
}

func initialSyncModel(siteName, targetDir, direction string) syncProgressModel {
	return syncProgressModel{
		siteName:  siteName,
		targetDir: targetDir,
		direction: direction,
		files:     []string{},
		errors:    []error{},
	}
}

func (m syncProgressModel) Init() tea.Cmd {
	return tea.Tick(time.Second/10, func(t time.Time) tea.Msg {
		return syncMsg{file: "Starting sync...", done: false}
	})
}

func (m syncProgressModel) Update(msg tea.Msg) (tea.Model, tea.Cmd) {
	switch msg := msg.(type) {
	case tea.KeyMsg:
		if msg.String() == "q" || msg.String() == "ctrl+c" {
			return m, tea.Quit
		}

	case syncMsg:
		if msg.err != nil {
			m.errors = append(m.errors, msg.err)
		}
		if msg.file != "" {
			m.files = append(m.files, msg.file)
		}
		if msg.done {
			m.done = true
			return m, tea.Quit
		}
		return m, nil

	default:
		return m, nil
	}

	return m, nil
}

func (m syncProgressModel) View() string {
	var s string
	s += fmt.Sprintf("\n  SharePoint Sync: %s\n\n", m.siteName)
	s += fmt.Sprintf("  Direction: %s\n", m.direction)
	s += fmt.Sprintf("  Target: %s\n\n", m.targetDir)
	s += fmt.Sprintf("  Files processed: %d\n", len(m.files))
	if len(m.errors) > 0 {
		s += fmt.Sprintf("  Errors: %d\n", len(m.errors))
	}
	if m.done {
		s += "\n  Sync completed!\n\n"
	} else {
		s += "\n  Press Ctrl+C to cancel\n\n"
	}
	return s
}

func addSyncCommand(cmd *cobra.Command) {
	syncCmd := &cobra.Command{
		Use:   "sync [site] [directory]",
		Short: "Sync SharePoint files to local directory",
		Long: `Synchronizes SharePoint files to a local directory.

Direction options:
  --direction=down    Download from SharePoint to local (default)
  --direction=up      Upload from local to SharePoint
  --direction=both    Bidirectional sync (not yet implemented)

Examples:
  teams-cli sharepoint sync "Documents" ./backup
  teams-cli sharepoint sync "Documents" ./backup --direction=up`,
		Args: cobra.RangeArgs(1, 2),
		Run: func(cmd *cobra.Command, args []string) {
			siteName := args[0]
			targetDir := "."
			if len(args) > 1 {
				targetDir = args[1]
			}

			direction, _ := cmd.Flags().GetString("direction")
			if direction == "" {
				direction = teams.SyncDirectionDown
			}

			maxFileSizeMB, _ := cmd.Flags().GetInt64("max-size-mb")
			maxFileSize := maxFileSizeMB * 1024 * 1024 // Convert MB to bytes

			cfg := loadConfig()
			if cfg == nil {
				fmt.Fprintln(os.Stderr, "Not authenticated. Run 'teams-cli auth login' first.")
				os.Exit(1)
			}

			// For now, run sync directly without TUI for simplicity
			fmt.Printf("Syncing SharePoint site '%s' to: %s\n", siteName, targetDir)
			fmt.Printf("Direction: %s\n\n", direction)

			m365Client := m365.NewClientFromConfig(cfg, m365.WithVerbose(verbose))
			spSvc := sharepoint.NewService()
			ctx := context.Background()

			siteID, err := spSvc.GetSiteID(ctx, m365Client, siteName)
			if err != nil {
				fmt.Fprintf(os.Stderr, "Error: %v\n", err)
				os.Exit(1)
			}

			incremental, _ := cmd.Flags().GetBool("incremental")

			opts := &teams.SyncOptions{
				Direction:    direction,
				ExcludePaths: []string{},
				MaxFileSize:  maxFileSize,
				Incremental:  incremental,
			}

			if err := spSvc.SyncFiles(ctx, m365Client, siteID, targetDir, opts); err != nil {
				fmt.Fprintf(os.Stderr, "Error: %v\n", err)
				os.Exit(1)
			}

			fmt.Println("\nSync completed successfully!")
		},
	}

	syncCmd.Flags().String("direction", "down", "Sync direction: down, up, or both")
	syncCmd.Flags().Int64("max-size-mb", 0, "Maximum file size to sync in MB (0 = no limit)")
	syncCmd.Flags().Bool("incremental", false, "Only download files changed since last sync")

	cmd.AddCommand(syncCmd)
}
