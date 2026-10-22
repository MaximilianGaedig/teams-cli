package main

import (
	"context"
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"

	"github.com/MaximilianGaedig/teams"
	"github.com/MaximilianGaedig/teams/auth"
	"github.com/MaximilianGaedig/teams/m365"
	"github.com/MaximilianGaedig/teams/sharepoint"
	teamservice "github.com/MaximilianGaedig/teams/teamservice"
	"github.com/MaximilianGaedig/teams/users"
	tea "github.com/charmbracelet/bubbletea"
	"github.com/charmbracelet/lipgloss"
	"github.com/spf13/cobra"
)

var (
	version = "0.1.0"
	verbose bool
)

var (
	titleStyle = lipgloss.NewStyle().
		Bold(true).
		Foreground(lipgloss.Color("#7D56F4")).
		MarginLeft(2)
)

func configPath() string {
	home, _ := os.UserHomeDir()
	return filepath.Join(home, ".config", "teams-cli", "config.json")
}

func loadConfig() *teams.Config {
	path := configPath()
	data, err := os.ReadFile(path)
	if err != nil {
		return nil
	}
	var cfg teams.Config
	if err := json.Unmarshal(data, &cfg); err != nil {
		return nil
	}
	return &cfg
}

func saveConfig(cfg *teams.Config) error {
	path := configPath()
	dir := filepath.Dir(path)
	if err := os.MkdirAll(dir, 0755); err != nil {
		return err
	}
	data, err := json.MarshalIndent(cfg, "", "  ")
	if err != nil {
		return err
	}
	return os.WriteFile(path, data, 0644)
}

func main() {
	rootCmd := &cobra.Command{
		Use:   "teams-cli",
		Short: "Microsoft Teams CLI",
		Long: titleStyle.Render("teams-cli") + `

A CLI for Microsoft Teams and SharePoint management.`,
		Version: version,
	}

	rootCmd.PersistentFlags().BoolVarP(&verbose, "verbose", "v", false, "verbose output")

	rootCmd.AddCommand(authCommands())
	rootCmd.AddCommand(userCommands())
	rootCmd.AddCommand(teamCommands())
	rootCmd.AddCommand(spCommands())

	if err := rootCmd.Execute(); err != nil {
		fmt.Fprintln(os.Stderr, err)
		os.Exit(1)
	}
}

func authCommands() *cobra.Command {
	cmd := &cobra.Command{Use: "auth", Short: "Authentication commands"}

	cmd.AddCommand(&cobra.Command{
		Use:   "login",
		Short: "Login with Microsoft 365",
		Run: func(cmd *cobra.Command, args []string) {
			// Re-login keeps the saved tenant and SharePoint settings.
			cfg := loadConfig()
			if cfg == nil {
				cfg = teams.DefaultConfig()
			}
			m365Client := m365.NewClientFromConfig(cfg, m365.WithVerbose(verbose))

			p := tea.NewProgram(initialLoginModel(m365Client))
			final, err := p.Run()
			if err != nil {
				fmt.Fprintf(os.Stderr, "Error: %v\n", err)
				os.Exit(1)
			}
			if err := loginOutcome(final); err != nil {
				fmt.Fprintf(os.Stderr, "Login failed: %v\n", err)
				os.Exit(1)
			}

			if err := saveConfig(cfg); err != nil {
				fmt.Fprintln(os.Stderr, "Warning: failed to save config:", err)
			}

			fmt.Println("\n  Discovering SharePoint tenant...")
			ctx := context.Background()
			m365Client = m365.NewClientFromConfig(cfg, m365.WithVerbose(verbose))
			svc := auth.NewService()

			spTenant, err := svc.DiscoverSharePointTenant(ctx, m365Client)
			if err != nil {
				fmt.Fprintf(os.Stderr, "  Warning: could not auto-discover SharePoint tenant: %v\n", err)
				fmt.Print("  Enter SharePoint tenant name manually (e.g., 'contoso'): ")
				if _, err := fmt.Scanln(&spTenant); err != nil {
					spTenant = ""
				}
			} else {
				fmt.Printf("  ✓ Auto-discovered SharePoint tenant: %s\n", spTenant)
			}

			cfg.SharePointTenant = spTenant

			if err := saveConfig(cfg); err != nil {
				fmt.Fprintln(os.Stderr, "  Warning: failed to save config:", err)
			}

			fmt.Println()
		},
	})

	cmd.AddCommand(&cobra.Command{
		Use:   "status",
		Short: "Check authentication status",
		Run: func(cmd *cobra.Command, args []string) {
			cfg := loadConfig()
			if cfg == nil {
				fmt.Println("Not configured")
				return
			}

			fmt.Println("=== Configuration ===")
			fmt.Printf("Client ID: %s\n", cfg.ClientID)
			fmt.Printf("Tenant ID: %s\n", cfg.TenantID)
			fmt.Printf("SharePoint Tenant: %s\n", cfg.SharePointTenant)

			m365Client := m365.NewClientFromConfig(cfg, m365.WithVerbose(verbose))

			fmt.Println("\n=== SharePoint Endpoints ===")
			fmt.Printf("SharePoint URL: %s\n", m365Client.GetSharePointURL())
			fmt.Printf("SharePoint My URL: %s\n", m365Client.GetSharePointMyURL())

			fmt.Println("\n=== Authentication Status ===")
			if m365Client.IsAuthenticated() {
				fmt.Println("Authenticated: yes")
				if m365Client.IsTokenValid() {
					fmt.Println("Token valid: yes")
				} else {
					fmt.Println("Token valid: no (need to re-authenticate)")
				}
			} else {
				fmt.Println("Authenticated: no")
			}
		},
	})

	cmd.AddCommand(&cobra.Command{
		Use:   "discover",
		Short: "Rediscover SharePoint tenant",
		Run: func(cmd *cobra.Command, args []string) {
			cfg := loadConfig()
			if cfg == nil {
				fmt.Println("Not configured")
				return
			}

			m365Client := m365.NewClientFromConfig(cfg, m365.WithVerbose(verbose))
			svc := auth.NewService()
			spTenant, err := svc.DiscoverSharePointTenant(context.Background(), m365Client)
			if err != nil {
				fmt.Fprintf(os.Stderr, "Error: %v\n", err)
				os.Exit(1)
			}

			cfg.SharePointTenant = spTenant
			if err := saveConfig(cfg); err != nil {
				fmt.Fprintf(os.Stderr, "Error saving config: %v\n", err)
				os.Exit(1)
			}

			fmt.Printf("Discovered SharePoint tenant: %s\n", spTenant)
		},
	})

	cmd.AddCommand(&cobra.Command{
		Use:   "logout",
		Short: "Logout and clear credentials",
		Run: func(cmd *cobra.Command, args []string) {
			cfg := loadConfig()
			if cfg == nil {
				fmt.Println("Not authenticated")
				return
			}
			m365Client := m365.NewClientFromConfig(cfg, m365.WithVerbose(verbose))
			if err := m365Client.ClearCredentials(); err != nil {
				fmt.Fprintln(os.Stderr, "Error:", err)
				return
			}
			fmt.Println("Logged out successfully")
		},
	})

	// Teams browser auth command
	cmd.AddCommand(&cobra.Command{
		Use:   "teams-browser-auth [authorization-code]",
		Short: "Authenticate using Teams web client browser flow",
		Long: `Authenticate using the Teams web client OAuth flow.
		
This uses Microsoft's official Teams web client credentials to obtain tokens
that can access the internal Teams API (teams.cloud.microsoft).

To use:
1. Open browser and go to: https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=5e3ce6c0-2b1f-4285-8d4b-75ee78787346&response_type=code&redirect_uri=https://teams.cloud.microsoft/v2&scope=https://api.spaces.skype.com/.default%20openid%20profile%20offline_access&response_mode=query
2. Login with your Microsoft account
3. Copy the 'code' parameter from the redirect URL
4. Run: teams-cli auth teams-browser-auth <code>`,
		Args: cobra.ExactArgs(1),
		Run: func(cmd *cobra.Command, args []string) {
			cfg := teams.DefaultConfig()
			m365Client := m365.NewClientFromConfig(cfg, m365.WithVerbose(verbose))
			svc := auth.NewService()

			ctx := context.Background()
			accessToken, refreshToken, err := svc.TeamsAuthCodeFlow(ctx, m365Client, args[0])
			if err != nil {
				fmt.Fprintf(os.Stderr, "Teams browser auth failed: %v\n", err)
				os.Exit(1)
			}

			// Save refresh token
			if err := m365Client.SaveCredentials(refreshToken); err != nil {
				fmt.Fprintln(os.Stderr, "Warning: failed to save credentials:", err)
			}

			if err := saveConfig(cfg); err != nil {
				fmt.Fprintln(os.Stderr, "Warning: failed to save config:", err)
			}

			fmt.Println("✓ Teams browser authentication successful")
			fmt.Printf("Access token acquired (length: %d)\n", len(accessToken))
		},
	})

	return cmd
}

func userCommands() *cobra.Command {
	cmd := &cobra.Command{Use: "users", Short: "User management commands"}

	cmd.AddCommand(&cobra.Command{
		Use:   "list",
		Short: "List users",
		Run: func(cmd *cobra.Command, args []string) {
			cfg := loadConfig()
			if cfg == nil {
				fmt.Fprintln(os.Stderr, "Not authenticated. Run 'teams-cli auth login' first.")
				os.Exit(1)
			}
			m365Client := m365.NewClientFromConfig(cfg, m365.WithVerbose(verbose))
			usersSvc := users.NewService()

			ctx := context.Background()
			list, err := usersSvc.ListUsers(ctx, m365Client)
			if err != nil {
				fmt.Fprintln(os.Stderr, "Error:", err)
				return
			}
			for _, u := range list {
				fmt.Printf("%s\t%s\t%s\n", u.ID, u.DisplayName, u.Mail)
			}
		},
	})

	cmd.AddCommand(&cobra.Command{
		Use:   "search [query]",
		Short: "Search users",
		Args:  cobra.ExactArgs(1),
		Run: func(cmd *cobra.Command, args []string) {
			cfg := loadConfig()
			if cfg == nil {
				fmt.Fprintln(os.Stderr, "Not authenticated. Run 'teams-cli auth login' first.")
				os.Exit(1)
			}
			m365Client := m365.NewClientFromConfig(cfg, m365.WithVerbose(verbose))
			usersSvc := users.NewService()

			ctx := context.Background()
			list, err := usersSvc.SearchUsers(ctx, m365Client, args[0])
			if err != nil {
				fmt.Fprintln(os.Stderr, "Error:", err)
				return
			}
			for _, u := range list {
				fmt.Printf("%s\t%s\t%s\n", u.ID, u.DisplayName, u.Mail)
			}
		},
	})

	return cmd
}

func teamCommands() *cobra.Command {
	cmd := &cobra.Command{Use: "teams", Short: "Teams management commands"}

	cmd.AddCommand(&cobra.Command{
		Use:   "list",
		Short: "List teams",
		Run: func(cmd *cobra.Command, args []string) {
			cfg := loadConfig()
			if cfg == nil {
				fmt.Fprintln(os.Stderr, "Not authenticated. Run 'teams-cli auth login' first.")
				os.Exit(1)
			}
			m365Client := m365.NewClientFromConfig(cfg, m365.WithVerbose(verbose))
			teamsSvc := teamservice.NewService()

			ctx := context.Background()
			list, err := teamsSvc.ListTeams(ctx, m365Client)
			if err != nil {
				fmt.Fprintln(os.Stderr, "Error:", err)
				return
			}
			for _, t := range list {
				fmt.Printf("%s\t%s\n", t.ID, t.DisplayName)
			}
		},
	})

	cmd.AddCommand(&cobra.Command{
		Use:   "members [team-id]",
		Short: "List team members",
		Args:  cobra.ExactArgs(1),
		Run: func(cmd *cobra.Command, args []string) {
			cfg := loadConfig()
			if cfg == nil {
				fmt.Fprintln(os.Stderr, "Not authenticated. Run 'teams-cli auth login' first.")
				os.Exit(1)
			}
			m365Client := m365.NewClientFromConfig(cfg, m365.WithVerbose(verbose))
			teamsSvc := teamservice.NewService()

			ctx := context.Background()
			members, err := teamsSvc.ListMembers(ctx, m365Client, args[0])
			if err != nil {
				fmt.Fprintln(os.Stderr, "Error:", err)
				return
			}
			for _, m := range members {
				fmt.Printf("%s\t%s\t%s\t%s\n", m.ID, m.UserID, m.Name, m.Email)
			}
		},
	})

	// Add member command
	cmd.AddCommand(&cobra.Command{
		Use:   "add-member [team-id] [user-id]",
		Short: "Add a member to a team",
		Args:  cobra.ExactArgs(2),
		Run: func(cmd *cobra.Command, args []string) {
			cfg := loadConfig()
			if cfg == nil {
				fmt.Fprintln(os.Stderr, "Not authenticated. Run 'teams-cli auth login' first.")
				os.Exit(1)
			}
			m365Client := m365.NewClientFromConfig(cfg, m365.WithVerbose(verbose))
			teamsSvc := teamservice.NewService()

			ctx := context.Background()
			member := &teams.AddMemberRequest{
				UserID: args[1],
				Roles:  []string{"member"},
			}
			result, err := teamsSvc.AddMember(ctx, m365Client, args[0], member)
			if err != nil {
				fmt.Fprintln(os.Stderr, "Error:", err)
				return
			}
			fmt.Printf("Added member: %s\n", result.ID)
		},
	})

	addTeamSyncCommand(cmd)

	channelsCmd := &cobra.Command{
		Use:   "channels [team-id]",
		Short: "List team channels with SharePoint URLs",
		Args:  cobra.ExactArgs(1),
		Run: func(cmd *cobra.Command, args []string) {
			cfg := loadConfig()
			if cfg == nil {
				fmt.Fprintln(os.Stderr, "Not authenticated. Run 'teams-cli auth login' first.")
				os.Exit(1)
			}
			m365Client := m365.NewClientFromConfig(cfg, m365.WithVerbose(verbose))
			teamsSvc := teamservice.NewService()

			ctx := context.Background()
			channels, err := teamsSvc.ListChannels(ctx, m365Client, args[0])
			if err != nil {
				fmt.Fprintln(os.Stderr, "Error:", err)
				return
			}

			showSP, _ := cmd.Flags().GetBool("sharepoint")
			for _, c := range channels {
				if showSP {
					folder, err := teamsSvc.GetFilesFolder(ctx, m365Client, args[0], c.ID)
					if err != nil {
						fmt.Printf("%s\t%s\t(sp: error: %v)\n", c.ID, c.DisplayName, err)
					} else {
						fmt.Printf("%s\t%s\t%s\n", c.ID, c.DisplayName, folder.WebUrl)
					}
				} else {
					fmt.Printf("%s\t%s\n", c.ID, c.DisplayName)
				}
			}
		},
	}
	channelsCmd.Flags().Bool("sharepoint", false, "Show SharePoint URLs for each channel")
	cmd.AddCommand(channelsCmd)

	// Channel members command
	cmd.AddCommand(&cobra.Command{
		Use:   "channel-members [team-id] [channel-id]",
		Short: "List channel members (requires admin consent)",
		Args:  cobra.ExactArgs(2),
		Run: func(cmd *cobra.Command, args []string) {
			cfg := loadConfig()
			if cfg == nil {
				fmt.Fprintln(os.Stderr, "Not authenticated. Run 'teams-cli auth login' first.")
				os.Exit(1)
			}
			m365Client := m365.NewClientFromConfig(cfg, m365.WithVerbose(verbose))
			teamsSvc := teamservice.NewService()

			ctx := context.Background()
			members, err := teamsSvc.ListChannelMembers(ctx, m365Client, args[0], args[1])
			if err != nil {
				fmt.Fprintln(os.Stderr, "Error:", err)
				return
			}
			for _, m := range members {
				fmt.Printf("%s\t%s\t%s\t%s\t%s\n", m.ID, m.UserID, m.DisplayName, m.Email, m.Roles)
			}
		},
	})

	return cmd
}

func spCommands() *cobra.Command {
	cmd := &cobra.Command{
		Use:     "sharepoint",
		Short:   "SharePoint commands",
		Aliases: []string{"sp"},
	}

	cmd.AddCommand(&cobra.Command{
		Use:   "sites [query]",
		Short: "Search SharePoint sites",
		Args:  cobra.ExactArgs(1),
		Run: func(cmd *cobra.Command, args []string) {
			cfg := loadConfig()
			if cfg == nil {
				fmt.Fprintln(os.Stderr, "Not authenticated. Run 'teams-cli auth login' first.")
				os.Exit(1)
			}
			m365Client := m365.NewClientFromConfig(cfg, m365.WithVerbose(verbose))
			spSvc := sharepoint.NewService()

			ctx := context.Background()
			sites, err := spSvc.SearchSites(ctx, m365Client, args[0])
			if err != nil {
				fmt.Fprintln(os.Stderr, "Error:", err)
				return
			}
			for _, s := range sites {
				fmt.Printf("%s\t%s\t%s\n", s.Name, s.DisplayName, s.WebUrl)
			}
		},
	})

	cmd.AddCommand(&cobra.Command{
		Use:   "files [site] [path]",
		Short: "List SharePoint files (use 'sites' to find site names)",
		Args:  cobra.RangeArgs(1, 2),
		Run: func(cmd *cobra.Command, args []string) {
			cfg := loadConfig()
			if cfg == nil {
				fmt.Fprintln(os.Stderr, "Not authenticated. Run 'teams-cli auth login' first.")
				os.Exit(1)
			}
			m365Client := m365.NewClientFromConfig(cfg, m365.WithVerbose(verbose))
			spSvc := sharepoint.NewService()

			ctx := context.Background()
			siteID, err := spSvc.GetSiteID(ctx, m365Client, args[0])
			if err != nil {
				fmt.Fprintln(os.Stderr, "Error:", err)
				return
			}
			path := ""
			if len(args) > 1 {
				path = args[1]
			}
			files, err := spSvc.ListFiles(ctx, m365Client, siteID, path)
			if err != nil {
				fmt.Fprintln(os.Stderr, "Error:", err)
				return
			}
			for _, f := range files {
				fmt.Printf("%s\t%d\n", f.Name, f.Size)
			}
		},
	})

	cmd.AddCommand(&cobra.Command{
		Use:   "users [site-name-or-id]",
		Short: "List SharePoint site users",
		Args:  cobra.ExactArgs(1),
		Run: func(cmd *cobra.Command, args []string) {
			cfg := loadConfig()
			if cfg == nil {
				fmt.Fprintln(os.Stderr, "Not authenticated. Run 'teams-cli auth login' first.")
				os.Exit(1)
			}
			m365Client := m365.NewClientFromConfig(cfg, m365.WithVerbose(verbose))
			spSvc := sharepoint.NewService()

			ctx := context.Background()
			siteID, err := spSvc.GetSiteID(ctx, m365Client, args[0])
			if err != nil {
				fmt.Fprintln(os.Stderr, "Error:", err)
				return
			}
			users, err := spSvc.ListSiteUsers(ctx, m365Client, siteID)
			if err != nil {
				fmt.Fprintln(os.Stderr, "Error:", err)
				return
			}
			for _, u := range users {
				fmt.Printf("%s\t%s\t%s\t%v\n", u.UserID, u.DisplayName, u.Email, u.Roles)
			}
		},
	})

	addSyncCommand(cmd)

	cmd.AddCommand(&cobra.Command{
		Use:   "download [site] [remote-path] [local-path]",
		Short: "Download a file from SharePoint",
		Args:  cobra.RangeArgs(2, 3),
		Run: func(cmd *cobra.Command, args []string) {
			cfg := loadConfig()
			if cfg == nil {
				fmt.Fprintln(os.Stderr, "Not authenticated. Run 'teams-cli auth login' first.")
				os.Exit(1)
			}
			m365Client := m365.NewClientFromConfig(cfg, m365.WithVerbose(verbose))
			spSvc := sharepoint.NewService()

			ctx := context.Background()
			siteName := args[0]
			remotePath := args[1]
			localPath := args[2]
			if len(args) <= 2 {
				localPath = filepath.Base(remotePath)
			}

			siteID, err := spSvc.GetSiteID(ctx, m365Client, siteName)
			if err != nil {
				fmt.Fprintf(os.Stderr, "Error: %v\n", err)
				os.Exit(1)
			}

			fileInfo, err := spSvc.GetFileByPath(ctx, m365Client, siteID, remotePath)
			if err != nil {
				fmt.Fprintf(os.Stderr, "Error: %v\n", err)
				os.Exit(1)
			}

			if fileInfo.IsFolder {
				fmt.Fprintln(os.Stderr, "Error: remote path is a folder, not a file")
				os.Exit(1)
			}

			data, err := spSvc.DownloadFile(ctx, m365Client, siteID, fileInfo.ID)
			if err != nil {
				fmt.Fprintf(os.Stderr, "Error: %v\n", err)
				os.Exit(1)
			}

			if err := os.WriteFile(localPath, data, 0644); err != nil {
				fmt.Fprintf(os.Stderr, "Error: %v\n", err)
				os.Exit(1)
			}

			fmt.Printf("Downloaded: %s -> %s (%d bytes)\n", remotePath, localPath, len(data))
		},
	})

	return cmd
}
