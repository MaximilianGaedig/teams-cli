package auth

import (
	"bytes"
	"context"
	"encoding/json"
	"fmt"
	"io"
	"net/http"
	"net/url"
	"strings"
	"time"

	"github.com/MaximilianGaedig/teams"
)

type Service struct{}

func NewService() *Service {
	return &Service{}
}

type DeviceCodeResponse struct {
	DeviceCode      string `json:"device_code"`
	UserCode        string `json:"user_code"`
	VerificationURI string `json:"verification_uri"`
	ExpiresIn       string `json:"expires_in"`
	Interval        string `json:"interval"`
	Message         string `json:"message"`
}

const teamsResource = "https://api.spaces.skype.com"
const graphResource = "https://graph.microsoft.com"

func (s *Service) StartDeviceLogin(ctx context.Context, client teams.Client) (*DeviceCodeResponse, error) {
	data := url.Values{}
	data.Set("client_id", client.GetClientID())
	data.Set("resource", teamsResource)

	resp, err := client.HTTPClient().PostForm(client.GetDeviceCodeURL(), data)
	if err != nil {
		return nil, err
	}
	defer resp.Body.Close() //nolint:errcheck

	var dc DeviceCodeResponse
	if err := json.NewDecoder(resp.Body).Decode(&dc); err != nil {
		return nil, err
	}

	// Set default verification URI if not provided
	if dc.VerificationURI == "" {
		dc.VerificationURI = "https://microsoft.com/devicelogin"
	}

	return &dc, nil
}

// ErrAuthorizationPending is returned when the user hasn't completed authorization yet
type ErrAuthorizationPending struct {
	Interval time.Duration
}

func (e *ErrAuthorizationPending) Error() string {
	return "authorization_pending"
}

// GetToken attempts to exchange the device code for tokens.
// Returns ErrAuthorizationPending if the user hasn't authorized yet.
// Caller should poll repeatedly until success or other error.
func (s *Service) GetToken(ctx context.Context, client teams.Client, deviceCode string) (accessToken, refreshToken string, err error) {
	data := url.Values{}
	data.Set("client_id", client.GetClientID())
	data.Set("grant_type", "urn:ietf:params:oauth:grant-type:device_code")
	data.Set("code", deviceCode)

	resp, err := client.HTTPClient().PostForm(client.GetTokenURL(), data)
	if err != nil {
		return "", "", err
	}

	token, refresh, err := parseOAuthError(resp)
	if err != nil {
		if strings.Contains(err.Error(), "authorization_pending") {
			return "", "", &ErrAuthorizationPending{Interval: 5 * time.Second}
		}
		if strings.Contains(err.Error(), "slow_down") {
			return "", "", &ErrAuthorizationPending{Interval: 10 * time.Second}
		}
		return "", "", err
	}
	return token, refresh, nil
}

func (s *Service) DiscoverSharePointTenant(ctx context.Context, client teams.Client) (string, error) {
	if client.GetSharePointURL() != "" && client.GetSharePointURL() != "https://.sharepoint.com" {
		return extractTenantFromURL(client.GetSharePointURL()), nil
	}

	// Try Teams user aggregate settings API - the proper way to discover SharePoint tenant
	tenant, err := s.discoverFromTeamsAPI(ctx, client)
	if err != nil {
		return "", fmt.Errorf("teams API discovery failed: %w", err)
	}

	return tenant, nil
}

func (s *Service) discoverFromTeamsAPI(ctx context.Context, client teams.Client) (string, error) {
	// Use Teams user aggregate settings API to get the SharePoint web URL
	// This is the proper way Teams discovers the SharePoint tenant
	url := "https://teams.microsoft.com/api/mt/part/emea-02/beta/users/useraggregatesettings"

	reqBody := map[string]interface{}{
		"tenantSiteUrl": true,
	}
	reqBodyBytes, _ := json.Marshal(reqBody)

	req, err := http.NewRequestWithContext(ctx, "POST", url, bytes.NewReader(reqBodyBytes))
	if err != nil {
		return "", err
	}

	// Get a Teams token for this request - the user_aggregate_settings API needs Teams scope
	token, err := s.getTeamsToken(ctx, client)
	if err != nil {
		return "", fmt.Errorf("failed to get Teams token: %w", err)
	}

	req.Header.Set("Authorization", "Bearer "+token)
	req.Header.Set("Content-Type", "application/json")

	resp, err := client.HTTPClient().Do(req)
	if err != nil {
		return "", err
	}
	defer resp.Body.Close() //nolint:errcheck

	body, _ := io.ReadAll(resp.Body)

	var result struct {
		TenantSiteUrl struct {
			Value struct {
				WebUrl string `json:"webUrl"`
			} `json:"value"`
		} `json:"tenantSiteUrl"`
	}

	if err := json.Unmarshal(body, &result); err != nil {
		return "", err
	}

	webUrl := result.TenantSiteUrl.Value.WebUrl
	if webUrl == "" {
		return "", fmt.Errorf("tenantSiteUrl not found in response")
	}

	// Parse URL like "https://contoso.sharepoint.com/_layouts/15/sharepoint.aspx"
	// Extract the tenant name (contoso)
	return extractTenantFromURL(webUrl), nil
}

func extractTenantFromURL(webUrl string) string {
	parts := strings.Split(webUrl, ".sharepoint.com")
	if len(parts) > 0 {
		httpsParts := strings.Split(parts[0], "https://")
		if len(httpsParts) > 1 {
			return httpsParts[1]
		}
	}
	return ""
}

func (s *Service) getAccessToken(ctx context.Context, client teams.Client, resource string) (string, error) {
	if client.GetRefreshToken() == "" {
		return "", fmt.Errorf("no refresh token available")
	}

	body := fmt.Sprintf("client_id=%s&grant_type=refresh_token&refresh_token=%s&resource=%s", client.GetClientID(), client.GetRefreshToken(), resource)

	resp, err := doOAuthPOST(ctx, client, client.GetTokenURL(), body, "https://teams.microsoft.com")
	if err != nil {
		return "", err
	}
	defer resp.Body.Close() //nolint:errcheck

	token, _, err := parseOAuthError(resp)
	return token, err
}

func (s *Service) getTeamsToken(ctx context.Context, client teams.Client) (string, error) {
	return s.getAccessToken(ctx, client, teamsResource)
}

func doOAuthPOST(ctx context.Context, client teams.Client, url string, body string, origin string) (*http.Response, error) {
	req, err := http.NewRequestWithContext(ctx, "POST", url, strings.NewReader(body))
	if err != nil {
		return nil, err
	}
	req.Header.Set("Content-Type", "application/x-www-form-urlencoded")
	if origin != "" {
		req.Header.Set("Origin", origin)
	}
	return client.HTTPClient().Do(req)
}

func parseOAuthError(resp *http.Response) (token, refresh string, err error) {
	var result struct {
		AccessToken      string `json:"access_token"`
		RefreshToken     string `json:"refresh_token"`
		Error            string `json:"error"`
		ErrorDescription string `json:"error_description"`
	}

	if err := json.NewDecoder(resp.Body).Decode(&result); err != nil {
		return "", "", fmt.Errorf("failed to decode response: %w", err)
	}
	_ = resp.Body.Close()

	if result.Error != "" {
		return "", "", fmt.Errorf("auth failed: %s - %s", result.Error, result.ErrorDescription)
	}
	return result.AccessToken, result.RefreshToken, nil
}
