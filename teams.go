package teams

import (
	"context"
	"net/http"
	"time"
)

type Config struct {
	ClientID         string
	TenantID         string
	SharePointTenant string
	DirectoryRoleID  string
}

const (
	RoleUserAdministrator = "3e2f6e2e-64c4-4a1c-b3f2-6c9f1f7c3e1c"
	RoleGlobalAdmin       = "62e90394-69f5-4237-9120-304932e14d84"

	// DefaultOAuthClientID is the official Microsoft Teams client ID for device code flow
	DefaultOAuthClientID = "1fec8e78-bce4-4aaf-ab1b-5451cc387264"
	// DefaultOAuthTenantID is the common tenant ID for Microsoft authentication ("common" works for any tenant)
	DefaultOAuthTenantID = "common"

	// TeamsWebOrigin hosts the Teams web client and its internal APIs. Microsoft
	// redirects teams.microsoft.com here, mandatorily from 2027-01-01 (MC1465764).
	TeamsWebOrigin = "https://teams.cloud.microsoft"
)

func DefaultConfig() *Config {
	return &Config{
		ClientID:        DefaultOAuthClientID,
		TenantID:        DefaultOAuthTenantID,
		DirectoryRoleID: RoleUserAdministrator,
	}
}

type TokenInfo struct {
	AccessToken string
	ExpiresAt   time.Time
}

type Client interface {
	IsAuthenticated() bool
	IsTokenValid() bool
	AcquireAccessToken(ctx context.Context, resource string) (string, error)
	AcquireSharePointCookie(ctx context.Context) (string, time.Time, error)
	GraphRequest(ctx context.Context, method, url string, body []byte) ([]byte, error)
	BetaGraphRequest(ctx context.Context, method, url string, body []byte) ([]byte, error)
	SharePointRequest(ctx context.Context, method, url string, body []byte) ([]byte, error)
	TeamsRequest(ctx context.Context, method, url string, body []byte) ([]byte, error)
	TeamsEMEARequest(ctx context.Context, method, url string, body []byte) ([]byte, error)
	HTTPClient() *http.Client
	GetDeviceCodeURL() string
	GetTokenURL() string
	GetSharePointURL() string
	GetRefreshToken() string
	GetClientID() string
}

type TeamsService interface {
	// Team management
	ListTeams(ctx context.Context, client Client) ([]Team, error)
	GetTeamDetails(ctx context.Context, client Client, teamID string) (*TeamDetails, error)
	CreateTeam(ctx context.Context, client Client, team *CreateTeamRequest) (*Team, error)
	DeleteTeam(ctx context.Context, client Client, teamID string) error
	CloneTeam(ctx context.Context, client Client, teamID string, request *CloneTeamRequest) (*Team, error)
	ArchiveTeam(ctx context.Context, client Client, teamID string) error
	RestoreTeam(ctx context.Context, client Client, teamID string) error

	// Channel management
	ListChannels(ctx context.Context, client Client, teamID string) ([]Channel, error)
	GetChannel(ctx context.Context, client Client, teamID, channelID string) (*Channel, error)
	CreateChannel(ctx context.Context, client Client, teamID string, channel *CreateChannelRequest) (*Channel, error)
	DeleteChannel(ctx context.Context, client Client, teamID, channelID string) error
	GetFilesFolder(ctx context.Context, client Client, teamID, channelID string) (*DriveItem, error)

	// Members management
	ListMembers(ctx context.Context, client Client, teamID string) ([]Member, error)
	AddMember(ctx context.Context, client Client, teamID string, member *AddMemberRequest) (*Member, error)
	RemoveMember(ctx context.Context, client Client, teamID, memberID string) error

	// Channel members management
	ListChannelMembers(ctx context.Context, client Client, teamID, channelID string) ([]ChannelMember, error)
	AddChannelMember(ctx context.Context, client Client, teamID, channelID string, member *AddChannelMemberRequest) (*ChannelMember, error)
	RemoveChannelMember(ctx context.Context, client Client, teamID, channelID, memberID string) error

	// Messages
	GetChannelMessages(ctx context.Context, client Client, teamID, channelID string) ([]Message, error)
	SendChannelMessage(ctx context.Context, client Client, teamID, channelID string, message *MessageRequest) (*Message, error)

	// Apps
	ListInstalledApps(ctx context.Context, client Client, teamID string) ([]InstalledApp, error)
}

type UsersService interface {
	// ListUsers returns all users by automatically paginating through all pages.
	// Note: This may make multiple API calls for large directories.
	ListUsers(ctx context.Context, client Client) ([]User, error)
	// ListUsersFirstPage returns only the first page of users (default page size from Graph API).
	// Use the returned nextLink with ListUsersNextPage to fetch additional pages.
	ListUsersFirstPage(ctx context.Context, client Client) ([]User, string, error)
	// ListUsersNextPage fetches the next page of users using the nextLink from a previous response.
	ListUsersNextPage(ctx context.Context, client Client, nextLink string) ([]User, string, error)
	// SearchUsers searches for users matching the query
	SearchUsers(ctx context.Context, client Client, query string) ([]User, error)
	// GetUser gets a specific user by ID
	GetUser(ctx context.Context, client Client, userID string) (*User, error)
	// AddUser adds a user to a directory role
	AddUser(ctx context.Context, client Client, userID, role string) error
	// GetUserPresence gets the presence status of a user
	GetUserPresence(ctx context.Context, client Client, userID string) (*Presence, error)
	// GetUserPhoto gets the URL of user's profile photo
	GetUserPhoto(ctx context.Context, client Client, userID string) (*Photo, error)
	// GetUserManager gets the manager of a user
	GetUserManager(ctx context.Context, client Client, userID string) (*User, error)
}

type SharePointService interface {
	ListFiles(ctx context.Context, client Client, siteID, path string) ([]FileInfo, error)
	DownloadFile(ctx context.Context, client Client, siteID, fileID string) ([]byte, error)
	UploadFile(ctx context.Context, client Client, siteID, folderID, fileName string, content []byte) (*FileInfo, error)
	CreateFolder(ctx context.Context, client Client, siteID, parentFolderID, folderName string) (*FileInfo, error)
	DeleteFile(ctx context.Context, client Client, siteID, fileID string) error
	MoveFile(ctx context.Context, client Client, siteID, fileID, destinationFolderID string) (*FileInfo, error)
	CopyFile(ctx context.Context, client Client, siteID, fileID, destinationFolderID string) (*FileInfo, error)
	SyncFiles(ctx context.Context, client Client, siteID, localPath string, options *SyncOptions) error
	GetSiteID(ctx context.Context, client Client, siteName string) (string, error)
	ListSiteUsers(ctx context.Context, client Client, siteID string) ([]SiteUser, error)
}

// Team types
type Team struct {
	ID          string
	DisplayName string
	Description string
}

type TeamDetails struct {
	ID                string
	DisplayName       string
	Description       string
	InternalID        string
	Classification    string
	Specialization    string
	Visibility        string
	WebUrl            string
	IsArchived        bool
	MemberSettings    *MemberSettings
	GuestSettings     *GuestSettings
	MessagingSettings *MessagingSettings
	FunSettings       *FunSettings
	DiscoverySettings *DiscoverySettings
}

type MemberSettings struct {
	AllowCreateUpdateChannels         bool
	AllowDeleteChannels               bool
	AllowAddRemoveApps                bool
	AllowCreateUpdateRemoveConnectors bool
	AllowCreateUpdateRemoveTabs       bool
}

type GuestSettings struct {
	AllowCreateUpdateChannels bool
	AllowDeleteChannels       bool
}

type MessagingSettings struct {
	AllowUserEditMessages    bool
	AllowUserDeleteMessages  bool
	AllowOwnerDeleteMessages bool
	AllowTeamMentions        bool
	AllowChannelMentions     bool
}

type FunSettings struct {
	AllowGiphy            bool
	GiphyContentRating    string
	AllowStickersAndMemes bool
	AllowCustomMemes      bool
}

type DiscoverySettings struct {
	ShowInTeamsSearchAndSuggestions bool
}

type CreateTeamRequest struct {
	DisplayName string
	Description string
	Visibility  string   // "public" or "private"
	Owners      []string // User IDs
	Members     []string // User IDs
	Template    string   // Template ID
}

type CloneTeamRequest struct {
	DisplayName  string
	Description  string
	Visibility   string
	PartsToClone string // "apps,tabs,settings,channels,members"
}

// Channel types
type Channel struct {
	ID             string
	DisplayName    string
	Description    string
	WebUrl         string
	MembershipType string // "standard", "private", "shared"
}

type CreateChannelRequest struct {
	DisplayName    string
	Description    string
	MembershipType string   // "standard" or "private"
	Members        []string // User IDs (for private channels)
}

// Member types
type Member struct {
	ID     string
	Name   string
	Email  string
	UserID string
	Roles  []string // "owner", "member"
}

type AddMemberRequest struct {
	UserID string
	Roles  []string // "owner", "member"
}

// Channel Member types
type ChannelMember struct {
	ID          string
	DisplayName string
	Email       string
	UserID      string
	Roles       []string // "owner", "member"
	TenantID    string
}

type AddChannelMemberRequest struct {
	UserID string
	Roles  []string // "owner", "member"
}

// Message types
type Message struct {
	ID                   string
	ReplyToID            string
	From                 *IdentitySet
	Body                 *MessageBody
	CreatedDateTime      time.Time
	LastModifiedDateTime time.Time
	DeletedDateTime      *time.Time
	Subject              string
	Summary              string
	Attachments          []Attachment
	Reactions            []Reaction
	Mentions             []Mention
}

type MessageRequest struct {
	Body    *MessageBody
	Subject string
}

type MessageBody struct {
	ContentType string // "text" or "html"
	Content     string
}

type IdentitySet struct {
	User        *Identity
	Application *Identity
	Device      *Identity
}

type Identity struct {
	ID          string
	DisplayName string
}

type Attachment struct {
	ID          string
	Name        string
	ContentType string
	ContentUrl  string
}

type Reaction struct {
	ReactionType    string
	CreatedDateTime time.Time
	User            *IdentitySet
}

type Mention struct {
	ID          int
	Mentioned   *IdentitySet
	MentionText string
}

// App types
type InstalledApp struct {
	ID                 string
	TeamsApp           *TeamsApp
	TeamsAppDefinition *TeamsAppDefinition
}

type TeamsApp struct {
	ID                 string
	ExternalID         string
	DisplayName        string
	DistributionMethod string
}

type TeamsAppDefinition struct {
	ID          string
	TeamsAppID  string
	DisplayName string
	Version     string
}

// Drive item types
type DriveItem struct {
	ID     string
	Name   string
	WebUrl string
	Size   int64
	Folder *Folder
	File   *File
}

type Folder struct {
	ChildCount int
}

type File struct {
	MimeType string
}

// User types
type User struct {
	ID                string
	DisplayName       string
	Mail              string
	UserPrincipalName string
}

// SiteUser represents a user with access to a SharePoint site
type SiteUser struct {
	ID          string
	UserID      string
	DisplayName string
	Email       string
	Roles       []string // "write", "read", "owner"
}

type Presence struct {
	ID               string
	Availability     string // "Available", "Busy", "Away", "DoNotDisturb", "Offline"
	Activity         string // "Available", "InACall", "InAMeeting", etc.
	LastSeenDateTime time.Time
}

type Photo struct {
	ID     string
	Height int
	Width  int
	URL    string
}

// File types
type FileInfo struct {
	ID           string
	Name         string
	Path         string
	Size         int64
	LastModified time.Time
	IsFolder     bool
	Hash         string
}

// Sync direction constants
const (
	SyncDirectionDown = "down" // SharePoint -> Local
	SyncDirectionUp   = "up"   // Local -> SharePoint
	SyncDirectionBoth = "both" // Bidirectional (not yet implemented)
)

type SyncOptions struct {
	Direction    string // "down", "up", or "both"
	Parallel     int
	ExcludePaths []string
	MaxFileSize  int64 // Maximum file size in bytes (0 = no limit)
	Incremental  bool  // Only download files changed since last sync
}

type SyncState struct {
	LastSync   string              `json:"lastSync"`
	Site       string              `json:"site"`
	DeltaToken string              `json:"deltaToken,omitempty"`
	Files      map[string]FileMeta `json:"files"`
}

type FileMeta struct {
	ID       string `json:"id"`
	Size     int64  `json:"size"`
	Modified string `json:"modified"`
	Hash     string `json:"hash,omitempty"`
}

// ListOptions provides pagination options for list operations
type ListOptions struct {
	// Top specifies the number of items to return per page (max 999)
	Top int
	// NextLink is the URL for the next page (from previous response's @odata.nextLink)
	NextLink string
}

// ListResult represents a paginated list response with pagination info
type ListResult struct {
	// Items contains the actual items (type depends on the list operation)
	Items interface{}
	// NextLink is the URL for the next page, empty if no more pages
	NextLink string
	// Count is the total number of items returned in this page
	Count int
}

// SiteInfo represents a SharePoint site
type SiteInfo struct {
	ID          string
	Name        string
	DisplayName string
	WebUrl      string
}
