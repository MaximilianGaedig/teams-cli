# teams

[![Go Reference](https://pkg.go.dev/badge/github.com/MaximilianGaedig/teams.svg)](https://pkg.go.dev/github.com/MaximilianGaedig/teams)

A Go library and CLI for Microslop Teams and SharePoint.

## features

- **library**: Go client for Microslop Graph API (Teams, SharePoint, Users)
- **cli**: Command-line interface for common operations and syncing of sharepoint files to a local directory
- **authentication**: Device code flow (OAuth 2.0)
- **api coverage**: Teams, Channels, Messages, SharePoint Files, Users

## installation

```bash
go install github.com/MaximilianGaedig/teams/cmd/teams-cli@latest
```

### usage

```bash
# Login (opens browser for device code)
teams-cli auth login

# List your teams
teams-cli teams list

# List users
teams-cli users list

# Search users
teams-cli users search "john"

# List SharePoint files
teams-cli sharepoint files "Documents"
```

### library usage

also see the cli as a reference implementation

```go
package main

import (
    "context"
    "fmt"
    "log"
    "time"

    "github.com/MaximilianGaedig/teams"
    "github.com/MaximilianGaedig/teams/auth"
    "github.com/MaximilianGaedig/teams/m365"
    "github.com/MaximilianGaedig/teams/teamservice"
)

func main() {
    // Create config and client
    cfg := teams.DefaultConfig()
    client := m365.NewClientFromConfig(cfg)
    
    // Check if authentication needed
    if !client.IsAuthenticated() {
        // Start device code flow
        svc := auth.NewService()
        ctx := context.Background()
        
        dc, err := svc.StartDeviceLogin(ctx, client)
        if err != nil {
            log.Fatal(err)
        }
        
        fmt.Printf("Visit: %s\nEnter code: %s\n", dc.VerificationURI, dc.UserCode)
        
        // Poll for token (with external loop for UI control)
        var refreshToken string
        for {
            _, token, err := svc.GetToken(ctx, client, dc.DeviceCode)
            if err == nil {
                refreshToken = token
                break
            }
            if pendingErr, ok := err.(*auth.ErrAuthorizationPending); ok {
                fmt.Println("Waiting for authorization...")
                time.Sleep(pendingErr.Interval)
                continue
            }
            log.Fatal(err)
        }
        
        // Save credentials
        if err := client.SaveCredentials(refreshToken); err != nil {
            log.Fatal(err)
        }
    }
    
    // Now make API calls
    svc := teamservice.NewService()
    ctx := context.Background()
    teamList, err := svc.ListTeams(ctx, client)
    if err != nil {
        log.Fatal(err)
    }
    
    for _, t := range teamList {
        fmt.Printf("%s: %s\n", t.ID, t.DisplayName)
    }
}
```

## authentication

This library uses **device code flow** authentication. When you first run the CLI or use the library:

1. A device code URL and code will be displayed
2. Open the URL in your browser
3. Enter the code
4. Authenticate with your Microslop account
5. The refresh token is saved locally for future use

Credentials are stored in `~/.config/teams-cli/credentials.json`.

## api coverage

### teams (`github.com/MaximilianGaedig/teams/teamservice`)
- `ListTeams` - List joined teams
- `GetTeamDetails` - Get team info
- `ListChannels` - List team channels
- `GetChannel` - Get channel details
- `ListMembers` - List team members
- `GetChannelMessages` - Get channel messages
- `SendChannelMessage` - Post message
- `GetFilesFolder` - Get channel files location
- `ListInstalledApps` - List team apps

### sharepoint (`github.com/MaximilianGaedig/teams/sharepoint`)
- `ListFiles` - List files in site
- `DownloadFile` - Download file content
- `UploadFile` - Upload file
- `CreateFolder` - Create folder
- `DeleteFile` - Delete file/folder
- `MoveFile` - Move file
- `CopyFile` - Copy file
- `SyncFiles` - Sync local directory

### users (`github.com/MaximilianGaedig/teams/users`)
- `ListUsers` - List directory users
- `ListUsersFirstPage` - First page with pagination
- `ListUsersNextPage` - Next page
- `SearchUsers` - Search by name
- `GetUser` - Get user details
- `GetUserPresence` - Get availability
- `GetUserPhoto` - Get profile photo
- `GetUserManager` - Get manager

## pagination

List methods support pagination:

```go
import "github.com/MaximilianGaedig/teams/users"

// Get first page
users, nextLink, err := svc.ListUsersFirstPage(ctx, client)

// Get subsequent pages
for nextLink != "" {
    users, nextLink, err = svc.ListUsersNextPage(ctx, client, nextLink)
    // Process users...
}
```

Or use `ListUsers()` to automatically fetch all pages.

## error handling

The library provides typed errors:

```go
import "github.com/MaximilianGaedig/teams/m365"

users, err := svc.ListUsers(ctx, client)
if err != nil {
    if rateErr, ok := m365.IsRateLimitError(err); ok {
        fmt.Printf("Rate limited, retry after: %v\n", rateErr.RetryAfter)
    }
    // Handle other errors...
}
```

## rate limiting

The client automatically handles rate limiting with exponential backoff. Configure with:

```go
import "github.com/MaximilianGaedig/teams/m365"

client := m365.NewClientFromConfig(cfg, m365.WithRetry(3))
```

## testing

Run tests with:

```bash
go test ./...
```

Integration tests require authentication. Tests use a dedicated SharePoint test site and clean up after themselves.
