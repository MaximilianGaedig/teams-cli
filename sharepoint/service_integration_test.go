package sharepoint_test

import (
	"context"
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"
	"testing"
	"time"

	"github.com/MaximilianGaedig/teams"
	"github.com/MaximilianGaedig/teams/m365"
	"github.com/MaximilianGaedig/teams/sharepoint"
)

// testSiteID should be set via SHAREPOINT_TEST_SITE_ID environment variable
var testSiteID = os.Getenv("SHAREPOINT_TEST_SITE_ID")

func loadConfigForTest(t *testing.T) teams.Client {
	home, err := os.UserHomeDir()
	if err != nil {
		t.Fatalf("Cannot determine home directory: %v", err)
	}
	path := filepath.Join(home, ".config", "teams-cli", "config.json")
	data, err := os.ReadFile(path)
	if err != nil {
		t.Fatalf("No config file: %v", err)
	}
	var cfg teams.Config
	if err := json.Unmarshal(data, &cfg); err != nil {
		t.Fatalf("Cannot parse config: %v", err)
	}

	client := m365.NewClientFromConfig(&cfg)
	if !client.IsAuthenticated() {
		t.Fatal("Not authenticated")
	}
	return client
}

func TestService_ListFiles_API(t *testing.T) {
	client := loadConfigForTest(t)
	ctx := context.Background()
	svc := sharepoint.NewService()

	files, err := svc.ListFiles(ctx, client, testSiteID, "")
	if err != nil {
		t.Fatalf("ListFiles failed: %v", err)
	}

	t.Logf("Found %d items in test site", len(files))
}

func TestService_GetSiteID_API(t *testing.T) {
	client := loadConfigForTest(t)
	ctx := context.Background()
	svc := sharepoint.NewService()

	files, err := svc.ListFiles(ctx, client, testSiteID, "")
	if err != nil {
		t.Fatalf("Cannot access test site: %v", err)
	}

	t.Logf("Test site accessible, found %d items", len(files))
}

func TestService_DownloadFile_API(t *testing.T) {
	client := loadConfigForTest(t)
	ctx := context.Background()
	svc := sharepoint.NewService()

	files, err := svc.ListFiles(ctx, client, testSiteID, "")
	if err != nil {
		t.Fatalf("ListFiles failed: %v", err)
	}

	if len(files) == 0 {
		t.Skip("No files in test site to test DownloadFile")
	}

	var fileID string
	for _, f := range files {
		if !f.IsFolder {
			fileID = f.ID
			break
		}
	}
	if fileID == "" {
		t.Skip("No files in test site (only folders)")
	}

	data, err := svc.DownloadFile(ctx, client, testSiteID, fileID)
	if err != nil {
		t.Fatalf("DownloadFile failed: %v", err)
	}

	t.Logf("Downloaded file from test site, size: %d bytes", len(data))
}

func TestService_SyncFiles_API(t *testing.T) {
	client := loadConfigForTest(t)
	ctx := context.Background()
	svc := sharepoint.NewService()

	options := &teams.SyncOptions{
		Direction: teams.SyncDirectionDown,
	}

	err := svc.SyncFiles(ctx, client, testSiteID, "", options)
	if err != nil {
		t.Fatalf("SyncFiles failed: %v", err)
	}

	t.Logf("SyncFiles completed successfully on test site")
}

func TestService_CreateAndDeleteFolder_API(t *testing.T) {
	client := loadConfigForTest(t)
	ctx := context.Background()
	svc := sharepoint.NewService()

	testFolderName := fmt.Sprintf("test_folder_%d", time.Now().Unix())

	t.Logf("Creating test folder: %s in test site", testFolderName)
	folder, err := svc.CreateFolder(ctx, client, testSiteID, "root", testFolderName)
	if err != nil {
		t.Fatalf("CreateFolder failed: %v", err)
	}

	t.Logf("Successfully created folder: %s (ID: %s)", folder.Name, folder.ID)

	t.Logf("Deleting test folder: %s", folder.ID)
	err = svc.DeleteFile(ctx, client, testSiteID, folder.ID)
	if err != nil {
		t.Fatalf("PermanentDeleteFile failed: %v", err)
	}

	t.Logf("Successfully created and deleted test folder: %s", testFolderName)
}

func TestService_UploadAndDeleteFile_API(t *testing.T) {
	client := loadConfigForTest(t)
	ctx := context.Background()
	svc := sharepoint.NewService()

	testFolderName := fmt.Sprintf("teams_cli_test_%d", time.Now().Unix())
	t.Logf("Creating test folder: %s", testFolderName)

	folder, err := svc.CreateFolder(ctx, client, testSiteID, "root", testFolderName)
	if err != nil {
		t.Fatalf("CreateFolder failed: %v", err)
	}

	t.Logf("Created test folder: %s (ID: %s)", folder.Name, folder.ID)

	testFileName := fmt.Sprintf("test_file_%d.txt", time.Now().Unix())
	testContent := []byte("This is a test file created by teams-cli integration tests. Safe to delete.")

	t.Logf("Uploading test file: %s", testFileName)
	uploadedFile, err := svc.UploadFile(ctx, client, testSiteID, folder.ID, testFileName, testContent)
	if err != nil {
		_ = svc.DeleteFile(ctx, client, testSiteID, folder.ID)
		t.Fatalf("UploadFile failed: %v", err)
	}

	t.Logf("Successfully uploaded file: %s (ID: %s)", uploadedFile.Name, uploadedFile.ID)

	files, err := svc.ListFiles(ctx, client, testSiteID, testFolderName)
	if err != nil {
		_ = svc.DeleteFile(ctx, client, testSiteID, uploadedFile.ID)
		_ = svc.DeleteFile(ctx, client, testSiteID, folder.ID)
		t.Fatalf("ListFiles failed: %v", err)
	}

	found := false
	for _, f := range files {
		if f.Name == testFileName {
			found = true
			break
		}
	}
	if !found {
		_ = svc.DeleteFile(ctx, client, testSiteID, uploadedFile.ID)
		_ = svc.DeleteFile(ctx, client, testSiteID, folder.ID)
		t.Errorf("Uploaded file %s not found in file list", testFileName)
	}

	t.Logf("Deleting test file: %s", uploadedFile.ID)
	err = svc.DeleteFile(ctx, client, testSiteID, uploadedFile.ID)
	if err != nil {
		_ = svc.DeleteFile(ctx, client, testSiteID, folder.ID)
		t.Fatalf("DeleteFile (file) failed: %v", err)
	}

	t.Logf("Deleting test folder: %s", folder.ID)
	err = svc.DeleteFile(ctx, client, testSiteID, folder.ID)
	if err != nil {
		t.Fatalf("DeleteFile (folder) failed: %v", err)
	}

	t.Logf("Successfully uploaded file and cleaned up test folder")
}

func TestService_MoveFile_API(t *testing.T) {
	client := loadConfigForTest(t)
	ctx := context.Background()
	svc := sharepoint.NewService()

	sourceFolderName := fmt.Sprintf("move_test_source_%d", time.Now().Unix())
	sourceFolder, err := svc.CreateFolder(ctx, client, testSiteID, "root", sourceFolderName)
	if err != nil {
		t.Fatalf("CreateFolder (source) failed: %v", err)
	}

	destFolderName := fmt.Sprintf("move_test_dest_%d", time.Now().Unix())
	destFolder, err := svc.CreateFolder(ctx, client, testSiteID, "root", destFolderName)
	if err != nil {
		_ = svc.DeleteFile(ctx, client, testSiteID, sourceFolder.ID)
		t.Fatalf("CreateFolder (dest) failed: %v", err)
	}

	testFileName := fmt.Sprintf("file_to_move_%d.txt", time.Now().Unix())
	testContent := []byte("This file will be moved")
	uploadedFile, err := svc.UploadFile(ctx, client, testSiteID, sourceFolder.ID, testFileName, testContent)
	if err != nil {
		_ = svc.DeleteFile(ctx, client, testSiteID, sourceFolder.ID)
		_ = svc.DeleteFile(ctx, client, testSiteID, destFolder.ID)
		t.Fatalf("UploadFile failed: %v", err)
	}

	t.Logf("Moving file %s to folder %s", uploadedFile.Name, destFolder.Name)
	movedFile, err := svc.MoveFile(ctx, client, testSiteID, uploadedFile.ID, destFolder.ID)
	if err != nil {
		_ = svc.DeleteFile(ctx, client, testSiteID, uploadedFile.ID)
		_ = svc.DeleteFile(ctx, client, testSiteID, sourceFolder.ID)
		_ = svc.DeleteFile(ctx, client, testSiteID, destFolder.ID)
		t.Fatalf("MoveFile failed: %v", err)
	}

	t.Logf("Successfully moved file to: %s", movedFile.Name)

	_ = svc.DeleteFile(ctx, client, testSiteID, movedFile.ID)
	_ = svc.DeleteFile(ctx, client, testSiteID, sourceFolder.ID)
	_ = svc.DeleteFile(ctx, client, testSiteID, destFolder.ID)

	t.Logf("Successfully tested MoveFile and cleaned up")
}

func TestService_CopyFile_API(t *testing.T) {
	client := loadConfigForTest(t)
	ctx := context.Background()
	svc := sharepoint.NewService()

	sourceFolderName := fmt.Sprintf("copy_test_source_%d", time.Now().Unix())
	sourceFolder, err := svc.CreateFolder(ctx, client, testSiteID, "root", sourceFolderName)
	if err != nil {
		t.Fatalf("CreateFolder (source) failed: %v", err)
	}

	destFolderName := fmt.Sprintf("copy_test_dest_%d", time.Now().Unix())
	destFolder, err := svc.CreateFolder(ctx, client, testSiteID, "root", destFolderName)
	if err != nil {
		_ = svc.DeleteFile(ctx, client, testSiteID, sourceFolder.ID)
		t.Fatalf("CreateFolder (dest) failed: %v", err)
	}

	testFileName := fmt.Sprintf("file_to_copy_%d.txt", time.Now().Unix())
	testContent := []byte("This file will be copied")
	uploadedFile, err := svc.UploadFile(ctx, client, testSiteID, sourceFolder.ID, testFileName, testContent)
	if err != nil {
		_ = svc.DeleteFile(ctx, client, testSiteID, sourceFolder.ID)
		_ = svc.DeleteFile(ctx, client, testSiteID, destFolder.ID)
		t.Fatalf("UploadFile failed: %v", err)
	}

	t.Logf("Copying file %s to folder %s", uploadedFile.Name, destFolder.Name)
	_, err = svc.CopyFile(ctx, client, testSiteID, uploadedFile.ID, destFolder.ID)
	if err != nil {
		_ = svc.DeleteFile(ctx, client, testSiteID, uploadedFile.ID)
		_ = svc.DeleteFile(ctx, client, testSiteID, sourceFolder.ID)
		_ = svc.DeleteFile(ctx, client, testSiteID, destFolder.ID)
		t.Fatalf("CopyFile failed: %v", err)
	}

	t.Logf("Successfully copied file")

	_ = svc.DeleteFile(ctx, client, testSiteID, uploadedFile.ID)
	_ = svc.DeleteFile(ctx, client, testSiteID, sourceFolder.ID)
	_ = svc.DeleteFile(ctx, client, testSiteID, destFolder.ID)

	t.Logf("Successfully tested CopyFile and cleaned up")
}
