package sharepoint

import (
	"context"
	"encoding/json"
	"fmt"
	"net/url"
	"os"
	"path/filepath"
	"strings"
	"time"

	"github.com/MaximilianGaedig/teams"
)

type Service struct{}

func NewService() *Service {
	return &Service{}
}

func (s *Service) ListFiles(ctx context.Context, client teams.Client, siteID, path string) ([]teams.FileInfo, error) {
	endpoint := fmt.Sprintf("sites/%s/drive/root/children", siteID)
	if path != "" {
		// URL-encode the path to handle special characters (Polish letters, spaces, #, etc.)
		encodedPath := url.PathEscape(path)
		endpoint = fmt.Sprintf("sites/%s/drive/root:/%s:/children", siteID, encodedPath)
	}

	data, err := client.GraphRequest(ctx, "GET", endpoint, nil)
	if err != nil {
		return nil, err
	}

	var resp struct {
		Value []struct {
			ID           string    `json:"id"`
			Name         string    `json:"name"`
			Size         int64     `json:"size"`
			LastModified time.Time `json:"lastModifiedDateTime"`
			Folder       *struct {
				ChildCount int `json:"childCount"`
			} `json:"folder,omitempty"`
		} `json:"value"`
	}
	if err := json.Unmarshal(data, &resp); err != nil {
		return nil, err
	}

	files := make([]teams.FileInfo, len(resp.Value))
	for i, item := range resp.Value {
		isFolder := item.Folder != nil
		files[i] = teams.FileInfo{
			ID:           item.ID,
			Name:         item.Name,
			Path:         path,
			Size:         item.Size,
			LastModified: item.LastModified,
			IsFolder:     isFolder,
		}
	}
	return files, nil
}

func (s *Service) DownloadFile(ctx context.Context, client teams.Client, siteID, fileID string) ([]byte, error) {
	return client.GraphRequest(ctx, "GET", fmt.Sprintf("sites/%s/drive/items/%s/content", siteID, fileID), nil)
}

func (s *Service) GetSiteID(ctx context.Context, client teams.Client, siteName string) (string, error) {
	data, err := client.GraphRequest(ctx, "GET", "sites?search="+siteName, nil)
	if err != nil {
		return "", err
	}

	var resp struct {
		Value []struct {
			ID string `json:"id"`
		} `json:"value"`
	}
	if err := json.Unmarshal(data, &resp); err != nil {
		return "", err
	}

	if len(resp.Value) == 0 {
		return "", fmt.Errorf("site not found: %s", siteName)
	}
	return resp.Value[0].ID, nil
}

func (s *Service) SyncFiles(ctx context.Context, client teams.Client, siteID, localPath string, options *teams.SyncOptions) error {
	switch options.Direction {
	case teams.SyncDirectionDown, "":
		return s.syncDown(ctx, client, siteID, "", localPath, options)
	case teams.SyncDirectionUp:
		return s.syncUp(ctx, client, siteID, "", localPath, options)
	case teams.SyncDirectionBoth:
		return fmt.Errorf("bidirectional sync not yet implemented")
	default:
		return fmt.Errorf("invalid sync direction: %s", options.Direction)
	}
}

// syncDown downloads files from SharePoint to local directory
func (s *Service) syncDown(ctx context.Context, client teams.Client, siteID, remotePath, localPath string, options *teams.SyncOptions) error {
	// Create local directory if it doesn't exist
	if err := os.MkdirAll(localPath, 0755); err != nil {
		return fmt.Errorf("failed to create directory %s: %w", localPath, err)
	}

	// List files in remote directory
	files, err := s.ListFiles(ctx, client, siteID, remotePath)
	if err != nil {
		return fmt.Errorf("failed to list files in %s: %w", remotePath, err)
	}

	for _, file := range files {
		localFilePath := filepath.Join(localPath, file.Name)

		// Check if path should be excluded
		if s.shouldExclude(localFilePath, options.ExcludePaths) {
			continue
		}

		if file.IsFolder {
			// Recursively sync subdirectory
			subRemotePath := remotePath
			if subRemotePath != "" {
				subRemotePath = subRemotePath + "/" + file.Name
			} else {
				subRemotePath = file.Name
			}
			if err := s.syncDown(ctx, client, siteID, subRemotePath, localFilePath, options); err != nil {
				return err
			}
		} else {
			// Check file size limit
			if options.MaxFileSize > 0 && file.Size > options.MaxFileSize {
				fmt.Printf("Skipping %s (size %d exceeds limit %d)\n", localFilePath, file.Size, options.MaxFileSize)
				continue
			}

			// Download file
			fmt.Printf("Downloading: %s\n", localFilePath)
			data, err := s.DownloadFile(ctx, client, siteID, file.ID)
			if err != nil {
				return fmt.Errorf("failed to download %s: %w", file.Name, err)
			}
			if err := os.WriteFile(localFilePath, data, 0644); err != nil {
				return fmt.Errorf("failed to write %s: %w", localFilePath, err)
			}
		}
	}

	return nil
}

// syncUp uploads files from local directory to SharePoint with full folder mirroring
func (s *Service) syncUp(ctx context.Context, client teams.Client, siteID, remotePath, localPath string, options *teams.SyncOptions) error {
	// First, get or create the parent folder in SharePoint
	var parentFolderID string
	if remotePath == "" {
		// We're at root, get the root folder ID
		rootFiles, err := s.ListFiles(ctx, client, siteID, "")
		if err != nil {
			return fmt.Errorf("failed to get root folder: %w", err)
		}
		for _, f := range rootFiles {
			if f.IsFolder {
				parentFolderID = f.ID
				break
			}
		}
		if parentFolderID == "" {
			return fmt.Errorf("could not find root folder")
		}
	} else {
		// We need to find or create the folder at remotePath
		// For simplicity, we'll list the parent and look for our folder
		parentRemotePath := filepath.Dir(remotePath)
		folderName := filepath.Base(remotePath)

		if parentRemotePath == "." {
			parentRemotePath = ""
		}

		parentFiles, err := s.ListFiles(ctx, client, siteID, parentRemotePath)
		if err != nil {
			return fmt.Errorf("failed to list parent folder: %w", err)
		}

		// Look for existing folder
		for _, f := range parentFiles {
			if f.IsFolder && f.Name == folderName {
				parentFolderID = f.ID
				break
			}
		}

		// If folder doesn't exist, create it
		if parentFolderID == "" {
			// Get grandparent folder ID
			var grandparentID string
			if parentRemotePath == "" {
				// Get root folder ID
				rootFiles, err := s.ListFiles(ctx, client, siteID, "")
				if err != nil {
					return fmt.Errorf("failed to get root folder: %w", err)
				}
				for _, f := range rootFiles {
					if f.IsFolder {
						grandparentID = f.ID
						break
					}
				}
			}

			newFolder, err := s.CreateFolder(ctx, client, siteID, grandparentID, folderName)
			if err != nil {
				return fmt.Errorf("failed to create folder %s: %w", folderName, err)
			}
			parentFolderID = newFolder.ID
			fmt.Printf("Created folder: %s\n", remotePath)
		}
	}

	entries, err := os.ReadDir(localPath)
	if err != nil {
		return fmt.Errorf("failed to read directory %s: %w", localPath, err)
	}

	for _, entry := range entries {
		localFilePath := filepath.Join(localPath, entry.Name())

		// Check if path should be excluded
		if s.shouldExclude(localFilePath, options.ExcludePaths) {
			continue
		}

		if entry.IsDir() {
			// Recursively sync subdirectory
			subRemotePath := remotePath
			if subRemotePath != "" {
				subRemotePath = subRemotePath + "/" + entry.Name()
			} else {
				subRemotePath = entry.Name()
			}

			fmt.Printf("Uploading directory: %s\n", subRemotePath)
			if err := s.syncUp(ctx, client, siteID, subRemotePath, localFilePath, options); err != nil {
				return err
			}
		} else {
			// Upload file to the correct parent folder
			fmt.Printf("Uploading: %s -> %s/%s\n", localFilePath, remotePath, entry.Name())
			data, err := os.ReadFile(localFilePath)
			if err != nil {
				return fmt.Errorf("failed to read %s: %w", localFilePath, err)
			}

			_, err = s.UploadFile(ctx, client, siteID, parentFolderID, entry.Name(), data)
			if err != nil {
				return fmt.Errorf("failed to upload %s: %w", entry.Name(), err)
			}
		}
	}

	return nil
}

func (s *Service) shouldExclude(path string, excludePaths []string) bool {
	for _, exclude := range excludePaths {
		if matched, _ := filepath.Match(exclude, path); matched {
			return true
		}
		// Also check if the path contains the exclude pattern
		if strings.Contains(path, exclude) {
			return true
		}
	}
	return false
}

// UploadFile uploads a file to a folder in SharePoint
func (s *Service) UploadFile(ctx context.Context, client teams.Client, siteID, folderID, fileName string, content []byte) (*teams.FileInfo, error) {
	endpoint := fmt.Sprintf("sites/%s/drive/items/%s:/%s:/content", siteID, folderID, fileName)

	data, err := client.GraphRequest(ctx, "PUT", endpoint, content)
	if err != nil {
		return nil, err
	}

	var resp struct {
		ID                   string `json:"id"`
		Name                 string `json:"name"`
		Size                 int64  `json:"size"`
		LastModifiedDateTime string `json:"lastModifiedDateTime"`
		Folder               *struct {
			ChildCount int `json:"childCount"`
		} `json:"folder,omitempty"`
	}
	if err := json.Unmarshal(data, &resp); err != nil {
		return nil, err
	}

	return &teams.FileInfo{
		ID:       resp.ID,
		Name:     resp.Name,
		Path:     fileName,
		Size:     resp.Size,
		IsFolder: resp.Folder != nil,
	}, nil
}

// CreateFolder creates a new folder in SharePoint
func (s *Service) CreateFolder(ctx context.Context, client teams.Client, siteID, parentFolderID, folderName string) (*teams.FileInfo, error) {
	body := map[string]interface{}{
		"name":                              folderName,
		"folder":                            map[string]interface{}{},
		"@microsoft.graph.conflictBehavior": "rename",
	}

	jsonBody, err := json.Marshal(body)
	if err != nil {
		return nil, err
	}

	endpoint := fmt.Sprintf("sites/%s/drive/items/%s/children", siteID, parentFolderID)
	data, err := client.GraphRequest(ctx, "POST", endpoint, jsonBody)
	if err != nil {
		return nil, err
	}

	var resp struct {
		ID                   string `json:"id"`
		Name                 string `json:"name"`
		Size                 int64  `json:"size"`
		LastModifiedDateTime string `json:"lastModifiedDateTime"`
		Folder               *struct {
			ChildCount int `json:"childCount"`
		} `json:"folder,omitempty"`
	}
	if err := json.Unmarshal(data, &resp); err != nil {
		return nil, err
	}

	return &teams.FileInfo{
		ID:       resp.ID,
		Name:     resp.Name,
		Path:     folderName,
		Size:     resp.Size,
		IsFolder: true,
	}, nil
}

// DeleteFile deletes a file or folder from SharePoint (moves to recycle bin)
func (s *Service) DeleteFile(ctx context.Context, client teams.Client, siteID, fileID string) error {
	endpoint := fmt.Sprintf("sites/%s/drive/items/%s", siteID, fileID)
	_, err := client.GraphRequest(ctx, "DELETE", endpoint, nil)
	return err
}

// MoveFile moves a file to a different folder
func (s *Service) MoveFile(ctx context.Context, client teams.Client, siteID, fileID, destinationFolderID string) (*teams.FileInfo, error) {
	body := map[string]interface{}{
		"parentReference": map[string]string{
			"id": destinationFolderID,
		},
	}

	jsonBody, err := json.Marshal(body)
	if err != nil {
		return nil, err
	}

	endpoint := fmt.Sprintf("sites/%s/drive/items/%s", siteID, fileID)
	data, err := client.GraphRequest(ctx, "PATCH", endpoint, jsonBody)
	if err != nil {
		return nil, err
	}

	var resp struct {
		ID                   string `json:"id"`
		Name                 string `json:"name"`
		Size                 int64  `json:"size"`
		LastModifiedDateTime string `json:"lastModifiedDateTime"`
		Folder               *struct {
			ChildCount int `json:"childCount"`
		} `json:"folder,omitempty"`
	}
	if err := json.Unmarshal(data, &resp); err != nil {
		return nil, err
	}

	return &teams.FileInfo{
		Name:     resp.Name,
		Size:     resp.Size,
		IsFolder: resp.Folder != nil,
	}, nil
}

// CopyFile copies a file to a different folder
func (s *Service) CopyFile(ctx context.Context, client teams.Client, siteID, fileID, destinationFolderID string) (*teams.FileInfo, error) {
	body := map[string]interface{}{
		"parentReference": map[string]string{
			"id": destinationFolderID,
		},
	}

	jsonBody, err := json.Marshal(body)
	if err != nil {
		return nil, err
	}

	endpoint := fmt.Sprintf("sites/%s/drive/items/%s/copy", siteID, fileID)
	_, err = client.GraphRequest(ctx, "POST", endpoint, jsonBody)
	if err != nil {
		return nil, err
	}

	// Copy is async, return nil for now
	return nil, nil
}

// ListSiteUsers lists the users who have access to a SharePoint site
func (s *Service) ListSiteUsers(ctx context.Context, client teams.Client, siteID string) ([]teams.SiteUser, error) {
	return s.fetchAllSiteUsers(ctx, client, fmt.Sprintf("sites/%s/drive/root/permissions", siteID))
}

func (s *Service) fetchAllSiteUsers(ctx context.Context, client teams.Client, endpoint string) ([]teams.SiteUser, error) {
	var allUsers []teams.SiteUser
	nextLink := endpoint

	for nextLink != "" {
		users, newNextLink, err := s.fetchSiteUserPage(ctx, client, nextLink)
		if err != nil {
			return nil, err
		}
		allUsers = append(allUsers, users...)
		nextLink = newNextLink
	}

	return allUsers, nil
}

func (s *Service) fetchSiteUserPage(ctx context.Context, client teams.Client, endpoint string) ([]teams.SiteUser, string, error) {
	data, err := client.GraphRequest(ctx, "GET", endpoint, nil)
	if err != nil {
		return nil, "", err
	}

	var resp struct {
		Value []struct {
			ID          string   `json:"id"`
			Roles       []string `json:"roles"`
			GrantedToV2 struct {
				User struct {
					ID          string `json:"id"`
					DisplayName string `json:"displayName"`
					Email       string `json:"email"`
				} `json:"user"`
			} `json:"grantedToV2"`
			GrantedTo struct {
				User struct {
					ID          string `json:"id"`
					DisplayName string `json:"displayName"`
					Email       string `json:"email"`
				} `json:"user"`
			} `json:"grantedTo"`
		} `json:"value"`
		NextLink string `json:"@odata.nextLink"`
	}
	if err := json.Unmarshal(data, &resp); err != nil {
		return nil, "", err
	}

	users := make([]teams.SiteUser, len(resp.Value))
	for i, p := range resp.Value {
		// Try grantedToV2 first (newer format), fall back to grantedTo
		user := p.GrantedToV2.User
		if user.ID == "" {
			user = p.GrantedTo.User
		}

		users[i] = teams.SiteUser{
			ID:          p.ID,
			UserID:      user.ID,
			DisplayName: user.DisplayName,
			Email:       user.Email,
			Roles:       p.Roles,
		}
	}

	nextLink := ""
	if resp.NextLink != "" {
		if idx := len("https://graph.microsoft.com/v1.0/"); len(resp.NextLink) > idx {
			if strings.HasPrefix(resp.NextLink, "https://graph.microsoft.com/v1.0/") {
				nextLink = resp.NextLink[idx:]
			} else {
				nextLink = resp.NextLink
			}
		}
	}

	return users, nextLink, nil
}
