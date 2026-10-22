package main

import (
	"errors"
	"testing"

	"github.com/MaximilianGaedig/teams/auth"
	"github.com/MaximilianGaedig/teams/m365"
	tea "github.com/charmbracelet/bubbletea"
)

var errAuthFailed = errors.New("authentication failed")

func TestLoginModelInitialState(t *testing.T) {
	client := m365.NewClient("test-client-id", "test-tenant-id", "test-sharepoint-tenant")
	model := initialLoginModel(client)

	if model.stage != "initializing" {
		t.Errorf("Initial stage = %q, want %q", model.stage, "initializing")
	}
	if model.svc == nil {
		t.Error("Service should not be nil")
	}
	if model.client == nil {
		t.Error("Client should not be nil")
	}
	if model.done {
		t.Error("Initial done should be false")
	}
	if model.err != nil {
		t.Errorf("Initial err should be nil, got %v", model.err)
	}
}

func TestLoginModelView(t *testing.T) {
	client := m365.NewClient("test-client-id", "test-tenant-id", "test-sharepoint-tenant")
	model := initialLoginModel(client)

	view := model.View()
	if view == "" {
		t.Error("View should not be empty")
	}
	if len(view) < 10 {
		t.Errorf("View too short: %s", view)
	}
}

func TestLoginModelViewError(t *testing.T) {
	client := m365.NewClient("test-client-id", "test-tenant-id", "test-sharepoint-tenant")
	model := initialLoginModel(client)
	model.err = errAuthFailed
	model.stage = "error"

	view := model.View()
	if len(view) < 5 {
		t.Errorf("Error view too short: %s", view)
	}
}

func TestLoginModelViewSuccess(t *testing.T) {
	client := m365.NewClient("test-client-id", "test-tenant-id", "test-sharepoint-tenant")
	model := initialLoginModel(client)
	model.stage = "success"
	model.done = true

	view := model.View()
	if len(view) < 5 {
		t.Errorf("Success view too short: %s", view)
	}
}

func TestLoginModelViewWaiting(t *testing.T) {
	client := m365.NewClient("test-client-id", "test-tenant-id", "test-sharepoint-tenant")
	model := initialLoginModel(client)
	model.stage = "waiting"
	model.deviceCode = &auth.DeviceCodeResponse{
		VerificationURI: "https://login.microsoftonline.com/device",
		UserCode:        "ABC123",
	}

	view := model.View()
	if len(view) < 10 {
		t.Errorf("Waiting view too short: %s", view)
	}
}

func TestLoginModelUpdateKeyPressQuit(t *testing.T) {
	client := m365.NewClient("test-client-id", "test-tenant-id", "test-sharepoint-tenant")
	model := initialLoginModel(client)

	newModel, _ := model.Update(tea.KeyMsg{})

	if newModel.(loginModel).stage != "initializing" {
		t.Error("Empty key should not change stage")
	}
}

func TestLoginModelUpdateUnknownMsg(t *testing.T) {
	client := m365.NewClient("test-client-id", "test-tenant-id", "test-sharepoint-tenant")
	model := initialLoginModel(client)

	newModel, _ := model.Update("unknown")

	_, ok := newModel.(loginModel)
	if !ok {
		t.Error("Should return loginModel")
	}
}

func TestLoginModelTickCmd(t *testing.T) {
	client := m365.NewClient("test-client-id", "test-tenant-id", "test-sharepoint-tenant")
	model := initialLoginModel(client)

	cmd := model.tickCmd()
	if cmd == nil {
		t.Error("tickCmd should return a command")
	}
}

func TestLoginModelDeviceCodeMsg(t *testing.T) {
	client := m365.NewClient("test-client-id", "test-tenant-id", "test-sharepoint-tenant")
	model := initialLoginModel(client)

	dc := &auth.DeviceCodeResponse{
		DeviceCode:      "test-code",
		VerificationURI: "https://login.microsoftonline.com",
		UserCode:        "ABC123",
		ExpiresIn:       "900",
		Interval:        "5",
	}

	msg := deviceCodeMsg(dc)
	newModel, cmd := model.Update(msg)

	updated := newModel.(loginModel)
	if updated.stage != "waiting" {
		t.Errorf("stage = %q, want %q", updated.stage, "waiting")
	}
	if updated.deviceCode != dc {
		t.Error("deviceCode should be set")
	}
	if cmd == nil {
		t.Error("Should return command after device code received")
	}
}

func TestLoginModelTokenMsg(t *testing.T) {
	client := m365.NewClient("test-client-id", "test-tenant-id", "test-sharepoint-tenant")
	model := initialLoginModel(client)

	msg := tokenMsg("test-token")
	newModel, cmd := model.Update(msg)

	updated := newModel.(loginModel)
	if !updated.done {
		t.Error("done should be true after token received")
	}
	if cmd == nil {
		t.Error("Should return a command after successful token")
	}
}

func TestLoginModelTokenMsgSuccess(t *testing.T) {
	client := m365.NewClient("test-client-id", "test-tenant-id", "test-sharepoint-tenant")
	model := initialLoginModel(client)

	msg := tokenMsg("test-token")
	newModel, cmd := model.Update(msg)

	updated := newModel.(loginModel)
	if !updated.done {
		t.Error("done should be true after token received")
	}
	if cmd == nil {
		t.Error("Should return a command after successful token")
	}
	if updated.token != "test-token" {
		t.Errorf("token = %q, want %q", updated.token, "test-token")
	}
}

func TestLoginModelErrMsg(t *testing.T) {
	client := m365.NewClient("test-client-id", "test-tenant-id", "test-sharepoint-tenant")
	model := initialLoginModel(client)

	testErr := errAuthFailed
	msg := errMsg(testErr)
	newModel, cmd := model.Update(msg)

	updated := newModel.(loginModel)
	if updated.err != testErr {
		t.Error("err should be set from errMsg")
	}
	if updated.stage != "error" {
		t.Errorf("stage = %q, want %q", updated.stage, "error")
	}
	if cmd == nil {
		t.Error("Should return a command after error")
	}
}

func TestLoginModelInit(t *testing.T) {
	client := m365.NewClient("test-client-id", "test-tenant-id", "test-sharepoint-tenant")
	model := initialLoginModel(client)

	cmd := model.Init()
	if cmd == nil {
		t.Error("Init should return a command")
	}
}

func TestDeviceCodeMsgType(t *testing.T) {
	dc := &auth.DeviceCodeResponse{
		DeviceCode:      "test-code",
		VerificationURI: "https://login.microsoftonline.com",
		UserCode:        "ABC123",
	}

	msg := deviceCodeMsg(dc)
	if msg == nil {
		t.Error("deviceCodeMsg should not be nil")
	}
}

func TestTokenMsgType(t *testing.T) {
	msg := tokenMsg("test-token")
	if msg == "" {
		t.Error("tokenMsg should not be empty")
	}
}

func TestErrMsgType(t *testing.T) {
	testErr := errors.New("test error")
	msg := errMsg(testErr)
	if msg == nil {
		t.Error("errMsg should not be nil")
	}
}

func TestLoginOutcome(t *testing.T) {
	client := m365.NewClient("test-client-id", "test-tenant-id", "test-sharepoint-tenant")
	initial := initialLoginModel(client)

	if err := loginOutcome(initial); err == nil {
		t.Error("loginOutcome(initial model) = nil, want error: login has not finished")
	}

	succeeded, _ := initial.Update(tokenMsg("test-token"))
	if err := loginOutcome(succeeded); err != nil {
		t.Errorf("loginOutcome(after token) = %v, want nil", err)
	}

	failed, _ := initial.Update(errMsg(errAuthFailed))
	if err := loginOutcome(failed); err == nil {
		t.Error("loginOutcome(after error) = nil, want error")
	}
}
