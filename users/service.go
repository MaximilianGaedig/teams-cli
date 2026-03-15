package users

import (
	"context"
	"encoding/json"
	"fmt"
	"net/url"
	"strings"
	"time"

	"github.com/MaximilianGaedig/teams"
)

type Service struct{}

func NewService() *Service {
	return &Service{}
}

// ListUsers returns all users by automatically paginating through all pages.
// Note: This may make multiple API calls for large directories.
func (s *Service) ListUsers(ctx context.Context, client teams.Client) ([]teams.User, error) {
	return s.fetchAllUsers(ctx, client, "users")
}

// ListUsersFirstPage returns only the first page of users (default page size from Graph API).
// Use the returned nextLink with ListUsersNextPage to fetch additional pages.
func (s *Service) ListUsersFirstPage(ctx context.Context, client teams.Client) ([]teams.User, string, error) {
	return s.fetchUserPage(ctx, client, "users")
}

// ListUsersNextPage fetches the next page of users using the nextLink from a previous response.
func (s *Service) ListUsersNextPage(ctx context.Context, client teams.Client, nextLink string) ([]teams.User, string, error) {
	// Extract just the endpoint part from the full URL
	endpoint := nextLink
	if idx := strings.Index(nextLink, "/v1.0/"); idx != -1 {
		endpoint = nextLink[idx+6:]
	}
	return s.fetchUserPage(ctx, client, endpoint)
}

func (s *Service) SearchUsers(ctx context.Context, client teams.Client, query string) ([]teams.User, error) {
	return s.fetchAllUsers(ctx, client, "users?$filter=startsWith(displayName,'"+url.QueryEscape(query)+"')")
}

// fetchAllUsers automatically paginates through all pages and returns all users.
func (s *Service) fetchAllUsers(ctx context.Context, client teams.Client, endpoint string) ([]teams.User, error) {
	var allUsers []teams.User
	nextLink := endpoint

	for nextLink != "" {
		users, newNextLink, err := s.fetchUserPage(ctx, client, nextLink)
		if err != nil {
			return nil, err
		}
		allUsers = append(allUsers, users...)
		nextLink = newNextLink
	}

	return allUsers, nil
}

// fetchUserPage fetches a single page of users and returns the users + next link.
func (s *Service) fetchUserPage(ctx context.Context, client teams.Client, endpoint string) ([]teams.User, string, error) {
	data, err := client.GraphRequest(ctx, "GET", endpoint, nil)
	if err != nil {
		return nil, "", err
	}

	var resp struct {
		Value []struct {
			ID                string `json:"id"`
			DisplayName       string `json:"displayName"`
			Mail              string `json:"mail"`
			UserPrincipalName string `json:"userPrincipalName"`
		} `json:"value"`
		NextLink string `json:"@odata.nextLink"`
	}
	if err := json.Unmarshal(data, &resp); err != nil {
		return nil, "", err
	}

	users := make([]teams.User, len(resp.Value))
	for i, u := range resp.Value {
		users[i] = teams.User{ID: u.ID, DisplayName: u.DisplayName, Mail: u.Mail, UserPrincipalName: u.UserPrincipalName}
	}

	// Extract next link from @odata.nextLink
	nextLink := ""
	if resp.NextLink != "" {
		// Extract just the endpoint part after /v1.0/
		if idx := strings.Index(resp.NextLink, "/v1.0/"); idx != -1 {
			nextLink = resp.NextLink[idx+6:]
		} else {
			nextLink = resp.NextLink
		}
	}

	return users, nextLink, nil
}

func (s *Service) GetUser(ctx context.Context, client teams.Client, userID string) (*teams.User, error) {
	data, err := client.GraphRequest(ctx, "GET", "users/"+userID, nil)
	if err != nil {
		return nil, err
	}

	var u struct {
		ID                string `json:"id"`
		DisplayName       string `json:"displayName"`
		Mail              string `json:"mail"`
		UserPrincipalName string `json:"userPrincipalName"`
	}
	if err := json.Unmarshal(data, &u); err != nil {
		return nil, err
	}

	return &teams.User{ID: u.ID, DisplayName: u.DisplayName, Mail: u.Mail, UserPrincipalName: u.UserPrincipalName}, nil
}

func (s *Service) AddUser(ctx context.Context, client teams.Client, userID, role string) error {
	body := fmt.Sprintf(`{"@odata.id":"https://graph.microsoft.com/v1.0/directoryObjects/%s"}`, userID)
	_, err := client.GraphRequest(ctx, "POST", fmt.Sprintf("directoryRoles/roleTemplateId=%s/members/$ref", teams.RoleUserAdministrator), []byte(body))
	return err
}

// GetUserPresence gets the presence status of a user
func (s *Service) GetUserPresence(ctx context.Context, client teams.Client, userID string) (*teams.Presence, error) {
	data, err := client.GraphRequest(ctx, "GET", fmt.Sprintf("users/%s/presence", userID), nil)
	if err != nil {
		return nil, err
	}

	var p struct {
		ID               string    `json:"id"`
		Availability     string    `json:"availability"`
		Activity         string    `json:"activity"`
		LastSeenDateTime time.Time `json:"lastSeenDateTime"`
	}
	if err := json.Unmarshal(data, &p); err != nil {
		return nil, err
	}

	return &teams.Presence{
		ID:               p.ID,
		Availability:     p.Availability,
		Activity:         p.Activity,
		LastSeenDateTime: p.LastSeenDateTime,
	}, nil
}

// GetUserPhoto gets the URL of user's profile photo
func (s *Service) GetUserPhoto(ctx context.Context, client teams.Client, userID string) (*teams.Photo, error) {
	data, err := client.GraphRequest(ctx, "GET", fmt.Sprintf("users/%s/photo", userID), nil)
	if err != nil {
		return nil, err
	}

	var p struct {
		ID     string `json:"id"`
		Height int    `json:"height"`
		Width  int    `json:"width"`
	}
	if err := json.Unmarshal(data, &p); err != nil {
		return nil, err
	}

	// Build the photo URL
	photoURL := fmt.Sprintf("https://graph.microsoft.com/v1.0/users/%s/photo/$value", userID)

	return &teams.Photo{
		ID:     p.ID,
		Height: p.Height,
		Width:  p.Width,
		URL:    photoURL,
	}, nil
}

// GetUserManager gets the manager of a user
func (s *Service) GetUserManager(ctx context.Context, client teams.Client, userID string) (*teams.User, error) {
	data, err := client.GraphRequest(ctx, "GET", fmt.Sprintf("users/%s/manager", userID), nil)
	if err != nil {
		return nil, err
	}

	var u struct {
		ID                string `json:"id"`
		DisplayName       string `json:"displayName"`
		Mail              string `json:"mail"`
		UserPrincipalName string `json:"userPrincipalName"`
	}
	if err := json.Unmarshal(data, &u); err != nil {
		return nil, err
	}

	return &teams.User{
		ID:                u.ID,
		DisplayName:       u.DisplayName,
		Mail:              u.Mail,
		UserPrincipalName: u.UserPrincipalName,
	}, nil
}
