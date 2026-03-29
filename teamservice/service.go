package teamservice

import (
	"context"
	"encoding/json"
	"fmt"
	"time"

	"github.com/MaximilianGaedig/teams"
)

type Service struct{}

func NewService() *Service {
	return &Service{}
}

// ListTeams returns the teams the user is a member of
func (s *Service) ListTeams(ctx context.Context, client teams.Client) ([]teams.Team, error) {
	data, err := client.GraphRequest(ctx, "GET", "me/joinedTeams", nil)
	if err != nil {
		return nil, err
	}

	var resp struct {
		Value []struct {
			ID          string `json:"id"`
			DisplayName string `json:"displayName"`
			Description string `json:"description"`
		} `json:"value"`
	}
	if err := json.Unmarshal(data, &resp); err != nil {
		return nil, err
	}

	result := make([]teams.Team, len(resp.Value))
	for i, t := range resp.Value {
		result[i] = teams.Team{ID: t.ID, DisplayName: t.DisplayName, Description: t.Description}
	}
	return result, nil
}

// GetTeamDetails gets detailed information about a team
func (s *Service) GetTeamDetails(ctx context.Context, client teams.Client, teamID string) (*teams.TeamDetails, error) {
	data, err := client.GraphRequest(ctx, "GET", fmt.Sprintf("teams/%s", teamID), nil)
	if err != nil {
		return nil, err
	}

	var resp struct {
		ID                string                 `json:"id"`
		DisplayName       string                 `json:"displayName"`
		Description       string                 `json:"description"`
		InternalID        string                 `json:"internalId"`
		Classification    string                 `json:"classification"`
		Specialization    string                 `json:"specialization"`
		Visibility        string                 `json:"visibility"`
		WebUrl            string                 `json:"webUrl"`
		IsArchived        bool                   `json:"isArchived"`
		MemberSettings    map[string]interface{} `json:"memberSettings"`
		GuestSettings     map[string]interface{} `json:"guestSettings"`
		MessagingSettings map[string]interface{} `json:"messagingSettings"`
		FunSettings       map[string]interface{} `json:"funSettings"`
		DiscoverySettings map[string]interface{} `json:"discoverySettings"`
	}
	if err := json.Unmarshal(data, &resp); err != nil {
		return nil, err
	}

	details := &teams.TeamDetails{
		ID:             resp.ID,
		DisplayName:    resp.DisplayName,
		Description:    resp.Description,
		InternalID:     resp.InternalID,
		Classification: resp.Classification,
		Specialization: resp.Specialization,
		Visibility:     resp.Visibility,
		WebUrl:         resp.WebUrl,
		IsArchived:     resp.IsArchived,
	}

	// Parse member settings
	if resp.MemberSettings != nil {
		details.MemberSettings = &teams.MemberSettings{
			AllowCreateUpdateChannels:         getBool(resp.MemberSettings, "allowCreateUpdateChannels"),
			AllowDeleteChannels:               getBool(resp.MemberSettings, "allowDeleteChannels"),
			AllowAddRemoveApps:                getBool(resp.MemberSettings, "allowAddRemoveApps"),
			AllowCreateUpdateRemoveConnectors: getBool(resp.MemberSettings, "allowCreateUpdateRemoveConnectors"),
			AllowCreateUpdateRemoveTabs:       getBool(resp.MemberSettings, "allowCreateUpdateRemoveTabs"),
		}
	}

	// Parse guest settings
	if resp.GuestSettings != nil {
		details.GuestSettings = &teams.GuestSettings{
			AllowCreateUpdateChannels: getBool(resp.GuestSettings, "allowCreateUpdateChannels"),
			AllowDeleteChannels:       getBool(resp.GuestSettings, "allowDeleteChannels"),
		}
	}

	// Parse messaging settings
	if resp.MessagingSettings != nil {
		details.MessagingSettings = &teams.MessagingSettings{
			AllowUserEditMessages:    getBool(resp.MessagingSettings, "allowUserEditMessages"),
			AllowUserDeleteMessages:  getBool(resp.MessagingSettings, "allowUserDeleteMessages"),
			AllowOwnerDeleteMessages: getBool(resp.MessagingSettings, "allowOwnerDeleteMessages"),
			AllowTeamMentions:        getBool(resp.MessagingSettings, "allowTeamMentions"),
			AllowChannelMentions:     getBool(resp.MessagingSettings, "allowChannelMentions"),
		}
	}

	// Parse fun settings
	if resp.FunSettings != nil {
		details.FunSettings = &teams.FunSettings{
			AllowGiphy:            getBool(resp.FunSettings, "allowGiphy"),
			GiphyContentRating:    getString(resp.FunSettings, "giphyContentRating"),
			AllowStickersAndMemes: getBool(resp.FunSettings, "allowStickersAndMemes"),
			AllowCustomMemes:      getBool(resp.FunSettings, "allowCustomMemes"),
		}
	}

	// Parse discovery settings
	if resp.DiscoverySettings != nil {
		details.DiscoverySettings = &teams.DiscoverySettings{
			ShowInTeamsSearchAndSuggestions: getBool(resp.DiscoverySettings, "showInTeamsSearchAndSuggestions"),
		}
	}

	return details, nil
}

// CreateTeam creates a new team
func (s *Service) CreateTeam(ctx context.Context, client teams.Client, team *teams.CreateTeamRequest) (*teams.Team, error) {
	body := map[string]interface{}{
		"displayName": team.DisplayName,
		"description": team.Description,
		"visibility":  team.Visibility,
	}

	if team.Template != "" {
		body["template@odata.bind"] = fmt.Sprintf("https://graph.microsoft.com/v1.0/teamsTemplates('%s')", team.Template)
	}

	if len(team.Owners) > 0 {
		owners := make([]map[string]string, len(team.Owners))
		for i, owner := range team.Owners {
			owners[i] = map[string]string{"@odata.type": "#microsoft.graph.aadUserConversationMember", "user@odata.bind": fmt.Sprintf("https://graph.microsoft.com/v1.0/users('%s')", owner)}
		}
		body["owners@odata.bind"] = owners
	}

	if len(team.Members) > 0 {
		members := make([]map[string]string, len(team.Members))
		for i, member := range team.Members {
			members[i] = map[string]string{"@odata.type": "#microsoft.graph.aadUserConversationMember", "user@odata.bind": fmt.Sprintf("https://graph.microsoft.com/v1.0/users('%s')", member)}
		}
		body["members@odata.bind"] = members
	}

	jsonBody, err := json.Marshal(body)
	if err != nil {
		return nil, err
	}

	_, err = client.GraphRequest(ctx, "POST", "teams", jsonBody)
	if err != nil {
		return nil, err
	}

	return &teams.Team{
		DisplayName: team.DisplayName,
		Description: team.Description,
	}, nil
}

// DeleteTeam deletes a team
func (s *Service) DeleteTeam(ctx context.Context, client teams.Client, teamID string) error {
	_, err := client.GraphRequest(ctx, "DELETE", fmt.Sprintf("groups/%s", teamID), nil)
	return err
}

// CloneTeam creates a copy of a team
func (s *Service) CloneTeam(ctx context.Context, client teams.Client, teamID string, request *teams.CloneTeamRequest) (*teams.Team, error) {
	body := map[string]interface{}{
		"displayName":  request.DisplayName,
		"description":  request.Description,
		"visibility":   request.Visibility,
		"partsToClone": request.PartsToClone,
	}

	jsonBody, err := json.Marshal(body)
	if err != nil {
		return nil, err
	}

	_, err = client.GraphRequest(ctx, "POST", fmt.Sprintf("teams/%s/clone", teamID), jsonBody)
	if err != nil {
		return nil, err
	}

	return &teams.Team{
		DisplayName: request.DisplayName,
		Description: request.Description,
	}, nil
}

// ArchiveTeam archives a team
func (s *Service) ArchiveTeam(ctx context.Context, client teams.Client, teamID string) error {
	_, err := client.GraphRequest(ctx, "POST", fmt.Sprintf("teams/%s/archive", teamID), []byte(`{}`))
	return err
}

// RestoreTeam restores an archived team
func (s *Service) RestoreTeam(ctx context.Context, client teams.Client, teamID string) error {
	_, err := client.GraphRequest(ctx, "POST", fmt.Sprintf("teams/%s/unarchive", teamID), []byte(`{}`))
	return err
}

// ListChannels returns the channels in a team
func (s *Service) ListChannels(ctx context.Context, client teams.Client, teamID string) ([]teams.Channel, error) {
	data, err := client.GraphRequest(ctx, "GET", fmt.Sprintf("teams/%s/channels", teamID), nil)
	if err != nil {
		return nil, err
	}

	var resp struct {
		Value []struct {
			ID             string `json:"id"`
			DisplayName    string `json:"displayName"`
			Description    string `json:"description"`
			WebUrl         string `json:"webUrl"`
			MembershipType string `json:"membershipType"`
		} `json:"value"`
	}
	if err := json.Unmarshal(data, &resp); err != nil {
		return nil, err
	}

	channels := make([]teams.Channel, len(resp.Value))
	for i, c := range resp.Value {
		channels[i] = teams.Channel{
			ID:             c.ID,
			DisplayName:    c.DisplayName,
			Description:    c.Description,
			WebUrl:         c.WebUrl,
			MembershipType: c.MembershipType,
		}
	}
	return channels, nil
}

// GetChannel gets details of a specific channel
func (s *Service) GetChannel(ctx context.Context, client teams.Client, teamID, channelID string) (*teams.Channel, error) {
	data, err := client.GraphRequest(ctx, "GET", fmt.Sprintf("teams/%s/channels/%s", teamID, channelID), nil)
	if err != nil {
		return nil, err
	}

	var resp struct {
		ID             string `json:"id"`
		DisplayName    string `json:"displayName"`
		Description    string `json:"description"`
		WebUrl         string `json:"webUrl"`
		MembershipType string `json:"membershipType"`
	}
	if err := json.Unmarshal(data, &resp); err != nil {
		return nil, err
	}

	return &teams.Channel{
		ID:             resp.ID,
		DisplayName:    resp.DisplayName,
		Description:    resp.Description,
		WebUrl:         resp.WebUrl,
		MembershipType: resp.MembershipType,
	}, nil
}

// CreateChannel creates a new channel in a team
func (s *Service) CreateChannel(ctx context.Context, client teams.Client, teamID string, channel *teams.CreateChannelRequest) (*teams.Channel, error) {
	body := map[string]interface{}{
		"displayName":    channel.DisplayName,
		"description":    channel.Description,
		"membershipType": channel.MembershipType,
	}

	if len(channel.Members) > 0 {
		members := make([]map[string]string, len(channel.Members))
		for i, member := range channel.Members {
			members[i] = map[string]string{"@odata.type": "#microsoft.graph.aadUserConversationMember", "user@odata.bind": fmt.Sprintf("https://graph.microsoft.com/v1.0/users('%s')", member)}
		}
		body["members"] = members
	}

	jsonBody, err := json.Marshal(body)
	if err != nil {
		return nil, err
	}

	data, err := client.GraphRequest(ctx, "POST", fmt.Sprintf("teams/%s/channels", teamID), jsonBody)
	if err != nil {
		return nil, err
	}

	var resp teams.Channel
	if err := json.Unmarshal(data, &resp); err != nil {
		return nil, err
	}

	return &resp, nil
}

// DeleteChannel deletes a channel from a team
func (s *Service) DeleteChannel(ctx context.Context, client teams.Client, teamID, channelID string) error {
	_, err := client.GraphRequest(ctx, "DELETE", fmt.Sprintf("teams/%s/channels/%s", teamID, channelID), nil)
	return err
}

// GetFilesFolder gets the SharePoint folder for a channel
func (s *Service) GetFilesFolder(ctx context.Context, client teams.Client, teamID, channelID string) (*teams.DriveItem, error) {
	// Try Graph API first
	data, err := client.GraphRequest(ctx, "GET", fmt.Sprintf("teams/%s/channels/%s/filesFolder", teamID, channelID), nil)
	if err != nil {
		// Fallback to Teams internal API (emea regional endpoint)
		return s.getFilesFolderViaTeamsAPI(ctx, client, channelID)
	}

	var resp struct {
		ID     string `json:"id"`
		Name   string `json:"name"`
		WebUrl string `json:"webUrl"`
		Size   int64  `json:"size"`
		Folder struct {
			ChildCount int `json:"childCount"`
		} `json:"folder"`
		File struct {
			MimeType string `json:"mimeType"`
		} `json:"file"`
	}
	if err := json.Unmarshal(data, &resp); err != nil {
		return nil, err
	}

	item := &teams.DriveItem{
		ID:     resp.ID,
		Name:   resp.Name,
		WebUrl: resp.WebUrl,
		Size:   resp.Size,
	}

	if resp.Folder.ChildCount > 0 || (resp.File.MimeType == "" && resp.Folder.ChildCount == 0) {
		item.Folder = &teams.Folder{ChildCount: resp.Folder.ChildCount}
	}

	if resp.File.MimeType != "" {
		item.File = &teams.File{MimeType: resp.File.MimeType}
	}

	return item, nil
}

// getFilesFolderViaTeamsAPI uses the Teams internal EMEA API to get document library info
func (s *Service) getFilesFolderViaTeamsAPI(ctx context.Context, client teams.Client, channelID string) (*teams.DriveItem, error) {
	// Use Teams EMEA regional API
	endpoint := fmt.Sprintf("mt/part/emea-02/beta/channels/%s/documentlibraries", channelID)

	data, err := client.TeamsRequest(ctx, "GET", endpoint, nil)
	if err != nil {
		return nil, fmt.Errorf("Teams API document libraries failed: %w", err)
	}

	var resp struct {
		Value []struct {
			ID      string `json:"id"`
			Name    string `json:"name"`
			WebUrl  string `json:"webUrl"`
			DriveID string `json:"driveId"`
		} `json:"value"`
	}
	if err := json.Unmarshal(data, &resp); err != nil {
		return nil, err
	}

	if len(resp.Value) == 0 {
		return nil, fmt.Errorf("no document libraries found")
	}

	// Return the first document library
	return &teams.DriveItem{
		ID:     resp.Value[0].ID,
		Name:   resp.Value[0].Name,
		WebUrl: resp.Value[0].WebUrl,
	}, nil
}

// ListMembers returns the members of a team
func (s *Service) ListMembers(ctx context.Context, client teams.Client, teamID string) ([]teams.Member, error) {
	// Try teams endpoint first, fall back to groups endpoint on 403
	data, err := client.GraphRequest(ctx, "GET", fmt.Sprintf("teams/%s/members", teamID), nil)
	if err != nil {
		// Try groups endpoint as fallback (uses different permissions)
		data, err = client.GraphRequest(ctx, "GET", fmt.Sprintf("groups/%s/members", teamID), nil)
		if err != nil {
			return nil, err
		}
	}

	var resp struct {
		Value []struct {
			ID                string   `json:"id"`
			DisplayName       string   `json:"displayName"`
			Email             string   `json:"email"`
			Mail              string   `json:"mail"`
			UserPrincipalName string   `json:"userPrincipalName"`
			UserID            string   `json:"userId"`
			Roles             []string `json:"roles"`
		} `json:"value"`
	}
	if err := json.Unmarshal(data, &resp); err != nil {
		return nil, err
	}

	members := make([]teams.Member, len(resp.Value))
	for i, m := range resp.Value {
		// Try both "email" (teams endpoint) and "mail" (groups endpoint)
		email := m.Email
		if email == "" {
			email = m.Mail
		}
		if email == "" {
			email = m.UserPrincipalName
		}

		userID := m.UserID
		if userID == "" {
			userID = m.ID
		}

		members[i] = teams.Member{
			ID:     m.ID,
			Name:   m.DisplayName,
			Email:  email,
			UserID: userID,
			Roles:  m.Roles,
		}
	}
	return members, nil
}

// AddMember adds a member to a team
func (s *Service) AddMember(ctx context.Context, client teams.Client, teamID string, member *teams.AddMemberRequest) (*teams.Member, error) {
	body := map[string]interface{}{
		"@odata.type":     "#microsoft.graph.aadUserConversationMember",
		"roles":           member.Roles,
		"user@odata.bind": fmt.Sprintf("https://graph.microsoft.com/v1.0/users('%s')", member.UserID),
	}

	jsonBody, err := json.Marshal(body)
	if err != nil {
		return nil, err
	}

	data, err := client.GraphRequest(ctx, "POST", fmt.Sprintf("teams/%s/members", teamID), jsonBody)
	if err != nil {
		return nil, err
	}

	var resp struct {
		ID     string   `json:"id"`
		UserID string   `json:"userId"`
		Roles  []string `json:"roles"`
	}
	if err := json.Unmarshal(data, &resp); err != nil {
		return nil, err
	}

	return &teams.Member{
		ID:     resp.ID,
		UserID: resp.UserID,
		Roles:  resp.Roles,
	}, nil
}

// RemoveMember removes a member from a team
func (s *Service) RemoveMember(ctx context.Context, client teams.Client, teamID, memberID string) error {
	_, err := client.GraphRequest(ctx, "DELETE", fmt.Sprintf("teams/%s/members/%s", teamID, memberID), nil)
	return err
}

// ListChannelMembers gets channel members
// Note: Requires 'ChannelMember.Read.All' permission with admin consent in Azure AD
func (s *Service) ListChannelMembers(ctx context.Context, client teams.Client, teamID, channelID string) ([]teams.ChannelMember, error) {
	// Use Graph API (requires admin consent for ChannelMember.Read.All)
	data, err := client.GraphRequest(ctx, "GET", fmt.Sprintf("teams/%s/channels/%s/members", teamID, channelID), nil)
	if err != nil {
		return nil, fmt.Errorf("channel members requires 'ChannelMember.Read.All' permission with admin consent: %w", err)
	}

	var resp struct {
		Value []struct {
			ID          string   `json:"id"`
			DisplayName string   `json:"displayName"`
			Email       string   `json:"email"`
			UserID      string   `json:"userId"`
			Roles       []string `json:"roles"`
			TenantID    string   `json:"tenantId"`
		} `json:"value"`
	}
	if err := json.Unmarshal(data, &resp); err != nil {
		return nil, err
	}

	members := make([]teams.ChannelMember, len(resp.Value))
	for i, m := range resp.Value {
		members[i] = teams.ChannelMember{
			ID:          m.ID,
			DisplayName: m.DisplayName,
			Email:       m.Email,
			UserID:      m.UserID,
			Roles:       m.Roles,
			TenantID:    m.TenantID,
		}
	}
	return members, nil
}

// AddChannelMember adds a member to a channel
func (s *Service) AddChannelMember(ctx context.Context, client teams.Client, teamID, channelID string, member *teams.AddChannelMemberRequest) (*teams.ChannelMember, error) {
	body := map[string]interface{}{
		"@odata.type":     "#microsoft.graph.aadUserConversationMember",
		"user@odata.bind": fmt.Sprintf("https://graph.microsoft.com/v1.0/users('%s')", member.UserID),
		"roles":           member.Roles,
	}

	bodyJSON, err := json.Marshal(body)
	if err != nil {
		return nil, err
	}

	data, err := client.GraphRequest(ctx, "POST", fmt.Sprintf("teams/%s/channels/%s/members", teamID, channelID), bodyJSON)
	if err != nil {
		return nil, err
	}

	var resp struct {
		ID          string   `json:"id"`
		DisplayName string   `json:"displayName"`
		Email       string   `json:"email"`
		UserID      string   `json:"userId"`
		Roles       []string `json:"roles"`
		TenantID    string   `json:"tenantId"`
	}
	if err := json.Unmarshal(data, &resp); err != nil {
		return nil, err
	}

	return &teams.ChannelMember{
		ID:          resp.ID,
		DisplayName: resp.DisplayName,
		Email:       resp.Email,
		UserID:      resp.UserID,
		Roles:       resp.Roles,
		TenantID:    resp.TenantID,
	}, nil
}

// RemoveChannelMember removes a member from a channel
func (s *Service) RemoveChannelMember(ctx context.Context, client teams.Client, teamID, channelID, memberID string) error {
	_, err := client.GraphRequest(ctx, "DELETE", fmt.Sprintf("teams/%s/channels/%s/members/%s", teamID, channelID, memberID), nil)
	return err
}

// GetChannelMessages gets messages from a channel
func (s *Service) GetChannelMessages(ctx context.Context, client teams.Client, teamID, channelID string) ([]teams.Message, error) {
	data, err := client.GraphRequest(ctx, "GET", fmt.Sprintf("teams/%s/channels/%s/messages", teamID, channelID), nil)
	if err != nil {
		return nil, err
	}

	var resp struct {
		Value []struct {
			ID                   string     `json:"id"`
			ReplyToID            string     `json:"replyToId"`
			CreatedDateTime      time.Time  `json:"createdDateTime"`
			LastModifiedDateTime time.Time  `json:"lastModifiedDateTime"`
			DeletedDateTime      *time.Time `json:"deletedDateTime"`
			Subject              string     `json:"subject"`
			Summary              string     `json:"summary"`
			Body                 struct {
				ContentType string `json:"contentType"`
				Content     string `json:"content"`
			} `json:"body"`
			From struct {
				User struct {
					ID          string `json:"id"`
					DisplayName string `json:"displayName"`
				} `json:"user"`
				Application struct {
					ID          string `json:"id"`
					DisplayName string `json:"displayName"`
				} `json:"application"`
			} `json:"from"`
			Attachments []struct {
				ID          string `json:"id"`
				Name        string `json:"name"`
				ContentType string `json:"contentType"`
				ContentUrl  string `json:"contentUrl"`
			} `json:"attachments"`
			Reactions []struct {
				ReactionType    string    `json:"reactionType"`
				CreatedDateTime time.Time `json:"createdDateTime"`
				User            struct {
					User struct {
						ID          string `json:"id"`
						DisplayName string `json:"displayName"`
					} `json:"user"`
				} `json:"user"`
			} `json:"reactions"`
			Mentions []struct {
				ID          int    `json:"id"`
				MentionText string `json:"mentionText"`
				Mentioned   struct {
					User struct {
						ID          string `json:"id"`
						DisplayName string `json:"displayName"`
					} `json:"user"`
				} `json:"mentioned"`
			} `json:"mentions"`
		} `json:"value"`
	}
	if err := json.Unmarshal(data, &resp); err != nil {
		return nil, err
	}

	messages := make([]teams.Message, len(resp.Value))
	for i, m := range resp.Value {
		msg := teams.Message{
			ID:                   m.ID,
			ReplyToID:            m.ReplyToID,
			CreatedDateTime:      m.CreatedDateTime,
			LastModifiedDateTime: m.LastModifiedDateTime,
			DeletedDateTime:      m.DeletedDateTime,
			Subject:              m.Subject,
			Summary:              m.Summary,
			Body: &teams.MessageBody{
				ContentType: m.Body.ContentType,
				Content:     m.Body.Content,
			},
		}

		// Parse sender
		if m.From.User.ID != "" {
			msg.From = &teams.IdentitySet{
				User: &teams.Identity{
					ID:          m.From.User.ID,
					DisplayName: m.From.User.DisplayName,
				},
			}
		} else if m.From.Application.ID != "" {
			msg.From = &teams.IdentitySet{
				Application: &teams.Identity{
					ID:          m.From.Application.ID,
					DisplayName: m.From.Application.DisplayName,
				},
			}
		}

		// Parse attachments
		if len(m.Attachments) > 0 {
			msg.Attachments = make([]teams.Attachment, len(m.Attachments))
			for j, a := range m.Attachments {
				msg.Attachments[j] = teams.Attachment{
					ID:          a.ID,
					Name:        a.Name,
					ContentType: a.ContentType,
					ContentUrl:  a.ContentUrl,
				}
			}
		}

		// Parse reactions
		if len(m.Reactions) > 0 {
			msg.Reactions = make([]teams.Reaction, len(m.Reactions))
			for j, r := range m.Reactions {
				msg.Reactions[j] = teams.Reaction{
					ReactionType:    r.ReactionType,
					CreatedDateTime: r.CreatedDateTime,
					User: &teams.IdentitySet{
						User: &teams.Identity{
							ID:          r.User.User.ID,
							DisplayName: r.User.User.DisplayName,
						},
					},
				}
			}
		}

		// Parse mentions
		if len(m.Mentions) > 0 {
			msg.Mentions = make([]teams.Mention, len(m.Mentions))
			for j, mnt := range m.Mentions {
				msg.Mentions[j] = teams.Mention{
					ID:          mnt.ID,
					MentionText: mnt.MentionText,
					Mentioned: &teams.IdentitySet{
						User: &teams.Identity{
							ID:          mnt.Mentioned.User.ID,
							DisplayName: mnt.Mentioned.User.DisplayName,
						},
					},
				}
			}
		}

		messages[i] = msg
	}
	return messages, nil
}

// SendChannelMessage sends a message to a channel
func (s *Service) SendChannelMessage(ctx context.Context, client teams.Client, teamID, channelID string, message *teams.MessageRequest) (*teams.Message, error) {
	body := map[string]interface{}{
		"body": map[string]string{
			"contentType": message.Body.ContentType,
			"content":     message.Body.Content,
		},
	}

	if message.Subject != "" {
		body["subject"] = message.Subject
	}

	jsonBody, err := json.Marshal(body)
	if err != nil {
		return nil, err
	}

	data, err := client.GraphRequest(ctx, "POST", fmt.Sprintf("teams/%s/channels/%s/messages", teamID, channelID), jsonBody)
	if err != nil {
		return nil, err
	}

	var resp teams.Message
	if err := json.Unmarshal(data, &resp); err != nil {
		return nil, err
	}

	return &resp, nil
}

// ListInstalledApps lists apps installed in a team
func (s *Service) ListInstalledApps(ctx context.Context, client teams.Client, teamID string) ([]teams.InstalledApp, error) {
	data, err := client.GraphRequest(ctx, "GET", fmt.Sprintf("teams/%s/installedApps?$expand=teamsAppDefinition", teamID), nil)
	if err != nil {
		return nil, err
	}

	var resp struct {
		Value []struct {
			ID       string `json:"id"`
			TeamsApp struct {
				ID                 string `json:"id"`
				ExternalID         string `json:"externalId"`
				DisplayName        string `json:"displayName"`
				DistributionMethod string `json:"distributionMethod"`
			} `json:"teamsApp"`
			TeamsAppDefinition struct {
				ID          string `json:"id"`
				TeamsAppID  string `json:"teamsAppId"`
				DisplayName string `json:"displayName"`
				Version     string `json:"version"`
			} `json:"teamsAppDefinition"`
		} `json:"value"`
	}
	if err := json.Unmarshal(data, &resp); err != nil {
		return nil, err
	}

	apps := make([]teams.InstalledApp, len(resp.Value))
	for i, app := range resp.Value {
		apps[i] = teams.InstalledApp{
			ID: app.ID,
			TeamsApp: &teams.TeamsApp{
				ID:                 app.TeamsApp.ID,
				ExternalID:         app.TeamsApp.ExternalID,
				DisplayName:        app.TeamsApp.DisplayName,
				DistributionMethod: app.TeamsApp.DistributionMethod,
			},
			TeamsAppDefinition: &teams.TeamsAppDefinition{
				ID:          app.TeamsAppDefinition.ID,
				TeamsAppID:  app.TeamsAppDefinition.TeamsAppID,
				DisplayName: app.TeamsAppDefinition.DisplayName,
				Version:     app.TeamsAppDefinition.Version,
			},
		}
	}
	return apps, nil
}

// Helper functions
func getBool(m map[string]interface{}, key string) bool {
	if v, ok := m[key]; ok {
		if b, ok := v.(bool); ok {
			return b
		}
	}
	return false
}

func getString(m map[string]interface{}, key string) string {
	if v, ok := m[key]; ok {
		if s, ok := v.(string); ok {
			return s
		}
	}
	return ""
}
