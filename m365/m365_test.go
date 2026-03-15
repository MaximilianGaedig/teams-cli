package m365

import (
	"net/http"
	"net/http/httptest"
	"strings"
	"testing"
)

func TestErrorResponse(t *testing.T) {
	resp := &http.Response{
		StatusCode: http.StatusUnauthorized,
		Body:       http.NoBody,
	}

	err := CheckResponse(resp)
	if err == nil {
		t.Error("Expected error for 401")
	}
}

func TestCheckResponse_NoError(t *testing.T) {
	resp := &http.Response{
		StatusCode: http.StatusOK,
		Body:       http.NoBody,
	}

	err := CheckResponse(resp)
	if err != nil {
		t.Errorf("Expected no error for 200: %v", err)
	}
}

func TestCheckResponse_RateLimit(t *testing.T) {
	server := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		w.Header().Set("Retry-After", "60")
		w.WriteHeader(http.StatusTooManyRequests)
	}))
	defer server.Close()

	resp, _ := http.Get(server.URL)
	err := CheckResponse(resp)
	if err == nil {
		t.Error("Expected error for 429")
	}
}

func TestResponse_Decode(t *testing.T) {
	server := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		w.WriteHeader(http.StatusOK)
		w.Write([]byte(`{"value": "test"}`))
	}))
	defer server.Close()

	httpResp, _ := http.Get(server.URL)
	r := newResponse(httpResp)

	var result map[string]string
	err := r.Decode(&result)
	if err != nil {
		t.Errorf("Decode failed: %v", err)
	}
	if result["value"] != "test" {
		t.Errorf("Expected test, got %v", result)
	}
}

func TestResponse_GetBody(t *testing.T) {
	server := httptest.NewServer(http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		w.WriteHeader(http.StatusOK)
		w.Write([]byte("test body"))
	}))
	defer server.Close()

	httpResp, _ := http.Get(server.URL)
	r := newResponse(httpResp)

	body, err := r.GetBody()
	if err != nil {
		t.Errorf("GetBody failed: %v", err)
	}
	if string(body) != "test body" {
		t.Errorf("Expected 'test body', got %s", string(body))
	}
}

func TestResponse_HasNextPage(t *testing.T) {
	resp := &http.Response{
		StatusCode: http.StatusOK,
		Header:     http.Header{},
	}
	r := newResponse(resp)

	if r.HasNextPage() {
		t.Error("Expected no next page")
	}
}

func TestParseListResponse(t *testing.T) {
	body := []byte(`{"value":[{"id":"1"}],"@odata.nextLink":"https://graph.microsoft.com/users?$skiptoken=abc"}`)
	values, nextLink, err := parseListResponse(body)
	if err != nil {
		t.Errorf("parseListResponse failed: %v", err)
	}
	if len(values) == 0 {
		t.Error("Expected values")
	}
	if !strings.Contains(nextLink, "skiptoken") {
		t.Errorf("Expected nextLink with skiptoken, got %s", nextLink)
	}
}

func TestParseInt(t *testing.T) {
	tests := []struct {
		input    string
		fallback int
		expected int
	}{
		{"123", 0, 123},
		{"", 0, 0},
		{"abc", 0, 0},
		{"456", 100, 456},
	}

	for _, tt := range tests {
		result := parseInt(tt.input, tt.fallback)
		if result != tt.expected {
			t.Errorf("parseInt(%q, %d) = %d, want %d", tt.input, tt.fallback, result, tt.expected)
		}
	}
}

func TestGraphError(t *testing.T) {
	err := &GraphError{
		Code:    "400",
		Message: "Bad request",
	}
	if err.Error() == "" {
		t.Error("Expected error message")
	}
}
