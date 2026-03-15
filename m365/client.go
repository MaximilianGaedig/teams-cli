package m365

import (
	"context"
	"encoding/json"
	"fmt"
	"io"
	"log"
	"net/http"
	"net/url"
	"os"
	"path/filepath"
	"strings"
	"time"

	"github.com/MaximilianGaedig/teams"
)

const (
	GraphResource = "https://graph.microsoft.com"
)

type Client struct {
	CredentialsDir  string
	ClientID        string
	TenantID        string
	TokenURL        string
	DeviceCodeURL   string
	SharePointURL   string
	SharePointMyURL string
	RefreshToken    string
	tokenCache      map[string]*teams.TokenInfo
	cookieCache     string
	cookieExpiry    time.Time
	httpClient      *http.Client
	verbose         bool

	rateLimitCheck bool
	retryCount     int
}

type ClientOption func(*Client)

func WithHTTPClient(hc *http.Client) ClientOption {
	return func(c *Client) {
		c.httpClient = hc
	}
}

func WithCredentialsDir(dir string) ClientOption {
	return func(c *Client) {
		c.CredentialsDir = dir
	}
}

func WithClientID(id string) ClientOption {
	return func(c *Client) {
		c.ClientID = id
	}
}

func WithTenantID(id string) ClientOption {
	return func(c *Client) {
		c.TenantID = id
		c.TokenURL = fmt.Sprintf("https://login.microsoftonline.com/%s/oauth2/token", id)
		c.DeviceCodeURL = fmt.Sprintf("https://login.microsoftonline.com/%s/oauth2/devicecode", id)
	}
}

func WithSharePointTenant(tenant string) ClientOption {
	return func(c *Client) {
		c.SharePointURL = fmt.Sprintf("https://%s.sharepoint.com", tenant)
		c.SharePointMyURL = fmt.Sprintf("https://%s-my.sharepoint.com", tenant)
	}
}

func WithVerbose(v bool) ClientOption {
	return func(c *Client) {
		c.verbose = v
	}
}

func WithTimeout(timeout time.Duration) ClientOption {
	return func(c *Client) {
		c.httpClient.Timeout = timeout
	}
}

func WithRetry(count int) ClientOption {
	return func(c *Client) {
		c.retryCount = count
	}
}

func WithRateLimitCheck(enabled bool) ClientOption {
	return func(c *Client) {
		c.rateLimitCheck = enabled
	}
}

func NewClient(clientID, tenantID, sharePointTenant string, opts ...ClientOption) *Client {
	home, err := os.UserHomeDir()
	if err != nil {
		panic(fmt.Sprintf("failed to get home directory: %v", err))
	}
	credDir := filepath.Join(home, ".config", "teams-cli")
	if err := os.MkdirAll(credDir, 0750); err != nil {
		panic(fmt.Sprintf("failed to create credentials directory: %v", err))
	}

	c := &Client{
		CredentialsDir:  credDir,
		ClientID:        clientID,
		TenantID:        tenantID,
		TokenURL:        fmt.Sprintf("https://login.microsoftonline.com/%s/oauth2/token", tenantID),
		DeviceCodeURL:   fmt.Sprintf("https://login.microsoftonline.com/%s/oauth2/devicecode", tenantID),
		SharePointURL:   fmt.Sprintf("https://%s.sharepoint.com", sharePointTenant),
		SharePointMyURL: fmt.Sprintf("https://%s-my.sharepoint.com", sharePointTenant),
		tokenCache:      make(map[string]*teams.TokenInfo),
		httpClient:      &http.Client{},
	}

	for _, opt := range opts {
		opt(c)
	}

	c.loadCredentials()
	return c
}

func NewClientFromConfig(cfg *teams.Config, opts ...ClientOption) *Client {
	return NewClient(cfg.ClientID, cfg.TenantID, cfg.SharePointTenant, opts...)
}

func (c *Client) HTTPClient() *http.Client {
	return c.httpClient
}

func (c *Client) GetDeviceCodeURL() string {
	return c.DeviceCodeURL
}

func (c *Client) GetSharePointURL() string {
	return c.SharePointURL
}

func (c *Client) GetSharePointMyURL() string {
	return c.SharePointMyURL
}

func (c *Client) GetClientID() string {
	return c.ClientID
}

func (c *Client) GetTokenURL() string {
	return c.TokenURL
}

func (c *Client) GetRefreshToken() string {
	return c.RefreshToken
}

func (c *Client) IsAuthenticated() bool {
	return c.RefreshToken != ""
}

func (c *Client) IsTokenValid() bool {
	if c.RefreshToken == "" {
		return false
	}
	// Validate by attempting to acquire a token
	_, err := c.AcquireAccessToken(context.Background(), GraphResource)
	return err == nil
}

func (c *Client) ValidateAndRefreshToken(ctx context.Context) error {
	if c.RefreshToken == "" {
		return fmt.Errorf("not authenticated")
	}
	_, err := c.AcquireAccessToken(ctx, GraphResource)
	return err
}

func (c *Client) loadCredentials() {
	credFile := filepath.Join(c.CredentialsDir, "credentials.json")

	if !isPathWithinDir(credFile, c.CredentialsDir) {
		return
	}

	data, err := os.ReadFile(credFile)
	if err != nil {
		return
	}

	var creds struct {
		ClientID     string `json:"client_id"`
		TenantID     string `json:"tenant_id"`
		RefreshToken string `json:"refresh_token"`
		SavedAt      int64  `json:"saved_at"`
	}

	if err := json.Unmarshal(data, &creds); err != nil {
		return
	}

	c.RefreshToken = creds.RefreshToken
}

func (c *Client) SaveCredentials(refreshToken string) error {
	credFile := filepath.Join(c.CredentialsDir, "credentials.json")

	creds := map[string]interface{}{
		"client_id":     c.ClientID,
		"tenant_id":     c.TenantID,
		"refresh_token": refreshToken,
		"saved_at":      time.Now().Unix(),
	}

	data, err := json.MarshalIndent(creds, "", "  ")
	if err != nil {
		return err
	}

	return os.WriteFile(credFile, data, 0600)
}

func (c *Client) ClearCredentials() error {
	credFile := filepath.Join(c.CredentialsDir, "credentials.json")
	c.RefreshToken = ""
	c.tokenCache = make(map[string]*teams.TokenInfo)
	return os.Remove(credFile)
}

func (c *Client) AcquireAccessToken(ctx context.Context, resource string) (string, error) {
	if cached, ok := c.tokenCache[resource]; ok && time.Now().Before(cached.ExpiresAt.Add(-5*time.Minute)) {
		return cached.AccessToken, nil
	}

	if c.RefreshToken == "" {
		return "", fmt.Errorf("not authenticated")
	}

	data := url.Values{}
	data.Set("client_id", c.ClientID)
	data.Set("grant_type", "refresh_token")
	data.Set("refresh_token", c.RefreshToken)
	data.Set("resource", resource)

	resp, err := c.httpClient.PostForm(c.TokenURL, data)
	if err != nil {
		return "", err
	}
	defer resp.Body.Close() //nolint:errcheck

	respBody, err := io.ReadAll(resp.Body)
	if err != nil {
		return "", fmt.Errorf("failed to read response body: %w", err)
	}

	token, refresh, expiresIn, err := parseTokenResponse(respBody)
	if err != nil {
		return "", err
	}

	if refresh != "" && refresh != c.RefreshToken {
		c.RefreshToken = refresh
		_ = c.SaveCredentials(refresh)
	}

	c.tokenCache[resource] = &teams.TokenInfo{
		AccessToken: token,
		ExpiresAt:   time.Now().Add(time.Duration(expiresIn) * time.Second),
	}

	return token, nil
}

func (c *Client) GetGraphToken() (string, error) {
	return c.AcquireAccessToken(context.Background(), GraphResource)
}

func (c *Client) GetSharePointToken() (string, error) {
	return c.AcquireAccessToken(context.Background(), c.SharePointURL)
}

func (c *Client) AcquireSharePointCookie(ctx context.Context) (string, time.Time, error) {
	if c.cookieCache != "" && time.Now().Before(c.cookieExpiry) {
		c.logf("Using cached SharePoint cookie")
		return c.cookieCache, c.cookieExpiry, nil
	}

	token, err := c.GetSharePointToken()
	if err != nil {
		return "", time.Time{}, err
	}

	url := c.SharePointMyURL + "/_api/SP.OAuth.NativeClient/Authenticate"
	req, _ := http.NewRequestWithContext(ctx, "POST", url, strings.NewReader(""))
	req.Header.Set("Authorization", "Bearer "+token)
	req.Header.Set("Content-Type", "application/json")

	c.logf("POST %s", url)

	resp, err := c.httpClient.Do(req)
	if err != nil {
		return "", time.Time{}, err
	}
	defer resp.Body.Close() //nolint:errcheck

	cookie, err := extractSPOJar(resp)
	if err != nil {
		return "", time.Time{}, err
	}

	expiry := time.Now().Add(1 * time.Hour)
	c.cookieCache = cookie
	c.cookieExpiry = expiry

	return cookie, expiry, nil
}

func (c *Client) GraphRequest(ctx context.Context, method, endpoint string, body []byte) ([]byte, error) {
	token, err := c.GetGraphToken()
	if err != nil {
		return nil, err
	}

	url := GraphResource + "/v1.0/" + endpoint
	req, err := buildRequest(ctx, c.httpClient, method, url, token, body, "bearer")
	if err != nil {
		return nil, err
	}

	c.logf("%s %s", method, url)

	resp, err := c.doRequestWithRetry(req)
	if err != nil {
		return nil, err
	}
	defer resp.Body.Close() //nolint:errcheck

	if resp.StatusCode == 401 {
		delete(c.tokenCache, GraphResource)
		token, err = c.GetGraphToken()
		if err != nil {
			return nil, err
		}
		req.Header.Set("Authorization", "Bearer "+token)
		resp, err = c.doRequestWithRetry(req)
		if err != nil {
			return nil, err
		}
		defer resp.Body.Close() //nolint:errcheck
	}

	return handleResponse(resp, c.verbose, "%s %s", method, endpoint)
}

func (c *Client) SharePointRequest(ctx context.Context, method, endpoint string, body []byte) ([]byte, error) {
	cookie, _, err := c.AcquireSharePointCookie(ctx)
	if err != nil {
		return nil, err
	}

	url := c.SharePointMyURL + "/" + endpoint
	req, err := buildRequest(ctx, c.httpClient, method, url, cookie, body, "cookie")
	if err != nil {
		return nil, err
	}

	c.logf("%s %s", method, url)

	resp, err := c.doRequestWithRetry(req)
	if err != nil {
		return nil, err
	}
	defer resp.Body.Close() //nolint:errcheck

	return handleResponse(resp, c.verbose, "%s %s", method, endpoint)
}

func (c *Client) DownloadSharePointFile(ctx context.Context, serverRelativeUrl string) ([]byte, error) {
	token, err := c.GetSharePointToken()
	if err != nil {
		return nil, err
	}

	parts := strings.Split(serverRelativeUrl, "/sites/")
	if len(parts) < 2 {
		return nil, fmt.Errorf("invalid SharePoint URL format")
	}

	siteBase := parts[0]
	siteAndPath := parts[1]
	siteName := strings.Split(siteAndPath, "/")[0]
	siteUrl := siteBase + "/sites/" + siteName

	apiUrl := fmt.Sprintf("%s/_api/Web/GetFileByServerRelativeUrl('/%s')/$value", siteUrl, url.PathEscape(siteAndPath))

	req, _ := http.NewRequestWithContext(ctx, "GET", apiUrl, nil)
	req.Header.Set("Authorization", "Bearer "+token)

	c.logf("Downloading: %s", apiUrl)

	resp, err := c.doRequestWithRetry(req)
	if err != nil {
		return nil, err
	}
	defer resp.Body.Close() //nolint:errcheck

	if resp.StatusCode != 200 {
		return nil, fmt.Errorf("HTTP %d", resp.StatusCode)
	}

	return io.ReadAll(resp.Body)
}

type APIType int

const (
	APITypeGraph APIType = iota
	APITypeSharePoint
)

func (c *Client) RawRequest(ctx context.Context, apiType APIType, method, path string, body []byte) (*Response, error) {
	var resp *http.Response
	var err error

	switch apiType {
	case APITypeGraph:
		resp, err = c.rawGraphRequest(ctx, method, path, body)
	case APITypeSharePoint:
		resp, err = c.rawSharePointRequest(ctx, method, path, body)
	default:
		return nil, fmt.Errorf("unknown API type: %d", apiType)
	}

	if err != nil {
		return nil, err
	}

	return newResponse(resp), nil
}

func (c *Client) rawGraphRequest(ctx context.Context, method, path string, body []byte) (*http.Response, error) {
	token, err := c.GetGraphToken()
	if err != nil {
		return nil, err
	}

	url := GraphResource + "/v1.0/" + path
	req, err := buildRequest(ctx, c.httpClient, method, url, token, body, "bearer")
	if err != nil {
		return nil, err
	}

	c.logf("%s %s", method, url)

	resp, err := c.doRequestWithRetry(req)
	if err != nil {
		return nil, err
	}

	if err := CheckResponse(resp); err != nil {
		return nil, err
	}

	return resp, nil
}

func (c *Client) rawSharePointRequest(ctx context.Context, method, path string, body []byte) (*http.Response, error) {
	cookie, _, err := c.AcquireSharePointCookie(ctx)
	if err != nil {
		return nil, err
	}

	url := c.SharePointMyURL + "/" + path
	req, err := buildRequest(ctx, c.httpClient, method, url, cookie, body, "cookie")
	if err != nil {
		return nil, err
	}

	c.logf("%s %s", method, url)

	resp, err := c.doRequestWithRetry(req)
	if err != nil {
		return nil, err
	}

	if err := CheckResponse(resp); err != nil {
		return nil, err
	}

	return resp, nil
}

func isPathWithinDir(path, dir string) bool {
	absPath, err := filepath.Abs(path)
	if err != nil {
		return false
	}
	absDir, err := filepath.Abs(dir)
	if err != nil {
		return false
	}
	return strings.HasPrefix(absPath, absDir+string(filepath.Separator))
}

func (c *Client) logf(format string, args ...interface{}) {
	if c.verbose {
		log.Printf(format, args...)
	}
}

func buildRequest(ctx context.Context, client *http.Client, method, url, token string, body []byte, authType string) (*http.Request, error) {
	var req *http.Request
	reader := strings.NewReader(string(body))
	if body == nil {
		req, _ = http.NewRequestWithContext(ctx, method, url, nil)
	} else {
		req, _ = http.NewRequestWithContext(ctx, method, url, reader)
	}

	switch authType {
	case "bearer":
		req.Header.Set("Authorization", "Bearer "+token)
	case "cookie":
		req.Header.Set("Cookie", token)
		req.Header.Set("Accept", "application/json;odata=verbose")
	}

	if body != nil {
		req.Header.Set("Content-Type", "application/json")
	}
	return req, nil
}

func handleResponse(resp *http.Response, verbose bool, format string, args ...interface{}) ([]byte, error) {
	if verbose {
		log.Printf(format, args...)
	}

	if resp.StatusCode == 401 {
		return nil, fmt.Errorf("unauthorized")
	}

	if resp.StatusCode < 200 || resp.StatusCode > 299 {
		return nil, fmt.Errorf(format+" failed: HTTP %d", append(args, resp.StatusCode)...)
	}

	// For 204 No Content, return empty body
	if resp.StatusCode == 204 {
		return []byte{}, nil
	}

	body, err := io.ReadAll(resp.Body)
	if err != nil {
		return nil, fmt.Errorf("failed to read response body: %w", err)
	}

	if verbose {
		log.Printf("Response body: %s", string(body))
	}

	return body, nil
}

func extractSPOJar(resp *http.Response) (string, error) {
	cookies := resp.Header.Values("Set-Cookie")
	for _, cookie := range cookies {
		if strings.Contains(cookie, "SPOIDCRL=") {
			return strings.Split(cookie, ";")[0], nil
		}
	}
	return "", fmt.Errorf("failed to get SharePoint cookie")
}

func parseTokenResponse(body []byte) (token, refresh string, expiresIn int, err error) {
	var result struct {
		AccessToken  string      `json:"access_token"`
		RefreshToken string      `json:"refresh_token"`
		ExpiresIn    interface{} `json:"expires_in"`
		Error        string      `json:"error"`
		Description  string      `json:"error_description"`
	}

	if err = json.Unmarshal(body, &result); err != nil {
		return "", "", 0, err
	}

	if result.Error != "" {
		return "", "", 0, fmt.Errorf("%s: %s", result.Error, result.Description)
	}

	expiresIn = 3600
	if result.ExpiresIn != nil {
		switch v := result.ExpiresIn.(type) {
		case float64:
			expiresIn = int(v)
		case string:
			_, _ = fmt.Sscanf(v, "%d", &expiresIn)
		}
	}

	return result.AccessToken, result.RefreshToken, expiresIn, nil
}

func (c *Client) doRequestWithRetry(req *http.Request) (*http.Response, error) {
	var resp *http.Response
	var err error

	for attempt := 0; attempt <= c.retryCount; attempt++ {
		if attempt > 0 {
			// Calculate backoff with exponential increase and jitter
			backoff := time.Duration(1<<uint(attempt-1)) * time.Second
			if backoff > 30*time.Second {
				backoff = 30 * time.Second
			}
			c.logf("Retry %d/%d after %v", attempt, c.retryCount, backoff)
			time.Sleep(backoff)
		}

		resp, err = c.httpClient.Do(req)
		if err != nil {
			return nil, err
		}

		// Check for rate limiting (429) or server errors (5xx)
		if resp.StatusCode == http.StatusTooManyRequests {
			retryAfter := parseRetryAfter(resp.Header.Get("Retry-After"))
			if retryAfter > 0 {
				c.logf("Rate limited, waiting %v", retryAfter)
				time.Sleep(retryAfter)
				resp.Body.Close()
				continue
			}
			resp.Body.Close()
			if attempt < c.retryCount {
				continue
			}
		} else if resp.StatusCode >= 500 && resp.StatusCode < 600 {
			// Server error, may be transient
			resp.Body.Close()
			if attempt < c.retryCount {
				continue
			}
		}

		return resp, nil
	}

	return resp, err
}
