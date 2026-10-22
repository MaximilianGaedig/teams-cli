package main

import (
	"context"
	"fmt"
	"os/exec"
	"runtime"
	"strings"
	"time"

	"github.com/MaximilianGaedig/teams/auth"
	"github.com/MaximilianGaedig/teams/m365"
	"github.com/charmbracelet/bubbles/spinner"
	tea "github.com/charmbracelet/bubbletea"
	"github.com/charmbracelet/lipgloss"
)

var (
	codeStyle = lipgloss.NewStyle().
			Bold(true).
			Background(lipgloss.Color("#7D56F4")).
			Foreground(lipgloss.Color("#FFFFFF")).
			Padding(1, 3)

	loginTitleStyle = lipgloss.NewStyle().
			Bold(true).
			Foreground(lipgloss.Color("#7D56F4"))

	successStyle = lipgloss.NewStyle().
			Foreground(lipgloss.Color("#00FF00"))

	loginErrorStyle = lipgloss.NewStyle().
			Foreground(lipgloss.Color("#FF0000"))
)

type loginModel struct {
	spinner    spinner.Model
	client     *m365.Client
	svc        *auth.Service
	deviceCode *auth.DeviceCodeResponse
	stage      string
	err        error
	done       bool
	token      string
}

type deviceCodeMsg *auth.DeviceCodeResponse
type tokenMsg string
type errMsg error
type tickMsg time.Time

func initialLoginModel(client *m365.Client) loginModel {
	s := spinner.New()
	s.Spinner = spinner.Dot
	s.Style = lipgloss.NewStyle().Foreground(lipgloss.Color("#7D56F4"))

	return loginModel{
		spinner: s,
		client:  client,
		svc:     auth.NewService(),
		stage:   "initializing",
	}
}

func (m loginModel) Init() tea.Cmd {
	return tea.Batch(
		m.spinner.Tick,
		m.getDeviceCode(),
	)
}

func (m loginModel) Update(msg tea.Msg) (tea.Model, tea.Cmd) {
	switch msg := msg.(type) {
	case tea.KeyMsg:
		if msg.String() == "q" || msg.String() == "ctrl+c" {
			return m, tea.Quit
		}

	case deviceCodeMsg:
		m.deviceCode = msg
		m.stage = "waiting"
		return m, tea.Batch(
			m.spinner.Tick,
			m.pollForToken(),
			m.tickCmd(),
		)

	case tokenMsg:
		m.token = string(msg)
		if err := m.client.SaveCredentials(m.token); err != nil {
			m.err = fmt.Errorf("failed to save credentials: %w", err)
			m.stage = "error"
			return m, tea.Quit
		}
		m.stage = "success"
		m.done = true
		return m, tea.Quit

	case errMsg:
		m.err = msg
		m.stage = "error"
		return m, tea.Quit

	case tickMsg:
		return m, m.tickCmd()

	default:
		var cmd tea.Cmd
		m.spinner, cmd = m.spinner.Update(msg)
		return m, cmd
	}

	return m, nil
}

func (m loginModel) View() string {
	if m.err != nil {
		return loginErrorStyle.Render(fmt.Sprintf("\n  Error: %v\n\n", m.err))
	}

	switch m.stage {
	case "initializing":
		return fmt.Sprintf("\n  %s Initializing...\n\n", m.spinner.View())

	case "waiting":
		var b strings.Builder
		b.WriteString("\n")
		b.WriteString(loginTitleStyle.Render("  Microsoft Teams Login"))
		b.WriteString("\n\n")

		// Show verification URL prominently
		if m.deviceCode.VerificationURI != "" {
			b.WriteString("  Open this link in your browser:\n")
			b.WriteString("  " + lipgloss.NewStyle().Foreground(lipgloss.Color("#00AAFF")).Underline(true).Render(m.deviceCode.VerificationURI))
			b.WriteString("\n\n")
		} else {
			b.WriteString("  Waiting for verification URL...\n\n")
		}

		b.WriteString("  And enter this code:\n")
		b.WriteString(codeStyle.Render(m.deviceCode.UserCode))
		b.WriteString("\n\n")
		fmt.Fprintf(&b, "  %s Waiting for authentication...\n\n", m.spinner.View())
		b.WriteString("  Press Ctrl+C to cancel\n\n")
		return b.String()

	case "success":
		return successStyle.Render("\n  ✓ Successfully authenticated!\n\n")

	default:
		return "\n"
	}
}

// openBrowser opens the URL in the default browser
func openBrowser(url string) error {
	var cmd string
	var args []string

	switch runtime.GOOS {
	case "windows":
		cmd = "cmd"
		args = []string{"/c", "start"}
	case "darwin":
		cmd = "open"
	default: // linux and others
		cmd = "xdg-open"
	}

	args = append(args, url)
	return exec.Command(cmd, args...).Start()
}

func (m loginModel) getDeviceCode() tea.Cmd {
	return func() tea.Msg {
		dc, err := m.svc.StartDeviceLogin(context.Background(), m.client)
		if err != nil {
			return errMsg(err)
		}

		// Try to open browser automatically
		if dc.VerificationURI != "" {
			if err := openBrowser(dc.VerificationURI); err != nil {
				// Browser open failed, that's ok - user can click the link
				fmt.Printf("  Could not open browser automatically: %v\n", err)
			}
		}

		return deviceCodeMsg(dc)
	}
}

func (m loginModel) pollForToken() tea.Cmd {
	return func() tea.Msg {
		ctx, cancel := context.WithTimeout(context.Background(), 5*time.Minute)
		defer cancel()

		for {
			select {
			case <-ctx.Done():
				return errMsg(fmt.Errorf("authentication timeout"))
			default:
			}

			_, refreshToken, err := m.svc.GetToken(ctx, m.client, m.deviceCode.DeviceCode)
			if err != nil {
				if pendingErr, ok := err.(*auth.ErrAuthorizationPending); ok {
					time.Sleep(pendingErr.Interval)
					continue
				}
				return errMsg(err)
			}

			// Only a refresh token can be persisted: an access token saved in its
			// place makes every later refresh fail with AADSTS9002313.
			if refreshToken == "" {
				return errMsg(fmt.Errorf("sign-in returned no refresh token, cannot stay logged in"))
			}
			return tokenMsg(refreshToken)
		}
	}
}

func (m loginModel) tickCmd() tea.Cmd {
	return tea.Tick(time.Second, func(t time.Time) tea.Msg {
		return tickMsg(t)
	})
}
