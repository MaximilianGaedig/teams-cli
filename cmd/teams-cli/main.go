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
			cfg := teams.DefaultConfig()
			m365Client := m365.NewClientFromConfig(cfg, m365.WithVerbose(verbose))

			model := initialLoginModel(m365Client)
			p := tea.NewProgram(model)
			if _, err := p.Run(); err != nil {
				fmt.Fprintf(os.Stderr, "Error: %v\n", err)
				os.Exit(1)
			}

			if model.err != nil {
				fmt.Fprintf(os.Stderr, "Login failed: %v\n", model.err)
				os.Exit(1)
			}

			if !model.done {
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

	addTeamSyncCommand(cmd)

	return cmd
}

func spCommands() *cobra.Command {
	cmd := &cobra.Command{
		Use:     "sharepoint",
		Short:   "SharePoint commands",
		Aliases: []string{"sp"},
	}

	cmd.AddCommand(&cobra.Command{
		Use:   "files [site]",
		Short: "List SharePoint files",
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
			files, err := spSvc.ListFiles(ctx, m365Client, siteID, "")
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

	return cmd
}
