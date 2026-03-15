package m365

import (
	"encoding/json"
	"errors"
	"fmt"
	"io"
	"net/http"
	"strconv"
	"time"
)

type ErrorResponse struct {
	Response *http.Response
	Err      *GraphError `json:"error"`
}

func (e *ErrorResponse) Error() string {
	if e.Err != nil {
		return e.Err.Message
	}
	return "unknown error"
}

type GraphError struct {
	Code       string      `json:"code"`
	Message    string      `json:"message"`
	InnerError *GraphError `json:"innerError"`
}

func (e *GraphError) Error() string {
	if e.InnerError != nil {
		return fmt.Sprintf("%s: %s (%s)", e.Message, e.InnerError.Message, e.InnerError.Code)
	}
	return fmt.Sprintf("%s (%s)", e.Message, e.Code)
}

type RateLimitError struct {
	Response   *http.Response
	RetryAfter time.Duration
	Message    string
}

func (e *RateLimitError) Error() string {
	if e.RetryAfter > 0 {
		return fmt.Sprintf("rate limited: %s (retry after %v)", e.Message, e.RetryAfter)
	}
	return fmt.Sprintf("rate limited: %s", e.Message)
}

func CheckResponse(resp *http.Response) error {
	if resp.StatusCode < 400 {
		return nil
	}

	body, err := io.ReadAll(resp.Body)
	if err != nil {
		return &ErrorResponse{
			Response: resp,
			Err:      &GraphError{Message: fmt.Sprintf("failed to read response body: %v", err)},
		}
	}

	switch resp.StatusCode {
	case http.StatusUnauthorized:
		return &ErrorResponse{
			Response: resp,
			Err:      &GraphError{Code: "Unauthorized", Message: "unauthorized"},
		}
	case http.StatusTooManyRequests:
		retryAfter := parseRetryAfter(resp.Header.Get("Retry-After"))
		return &RateLimitError{
			Response:   resp,
			RetryAfter: retryAfter,
			Message:    "too many requests",
		}
	case http.StatusForbidden:
		if remaining := resp.Header.Get("X-RateLimit-Remaining"); remaining == "0" {
			retryAfter := parseRetryAfter(resp.Header.Get("X-RateLimit-Reset"))
			return &RateLimitError{
				Response:   resp,
				RetryAfter: retryAfter,
				Message:    "rate limit exceeded",
			}
		}
		fallthrough
	default:
		var errResp ErrorResponse
		if json.Unmarshal(body, &errResp) == nil {
			errResp.Response = resp
			return &errResp
		}
		return &ErrorResponse{
			Response: resp,
			Err:      &GraphError{Message: fmt.Sprintf("HTTP %d: %s", resp.StatusCode, string(body))},
		}
	}
}

func parseRetryAfter(header string) time.Duration {
	if header == "" {
		return 0
	}
	if seconds, err := strconv.Atoi(header); err == nil {
		return time.Duration(seconds) * time.Second
	}
	if t, err := time.Parse(time.RFC1123, header); err == nil {
		return time.Until(t)
	}
	return 0
}

func IsErrorResponse(err error) (*ErrorResponse, bool) {
	var e *ErrorResponse
	if errors.As(err, &e) {
		return e, true
	}
	return nil, false
}

func IsRateLimitError(err error) (*RateLimitError, bool) {
	var e *RateLimitError
	if errors.As(err, &e) {
		return e, true
	}
	return nil, false
}

func (e *ErrorResponse) Unwrap() error {
	return e.Err
}

func (e *RateLimitError) Unwrap() error {
	return nil
}

func (e *ErrorResponse) Is(target error) bool {
	if t, ok := target.(*ErrorResponse); ok {
		return e.Response != nil && t.Response != nil && e.Response.StatusCode == t.Response.StatusCode
	}
	return false
}

func (e *RateLimitError) Is(target error) bool {
	if t, ok := target.(*RateLimitError); ok {
		return e.Response != nil && t.Response != nil && e.Response.StatusCode == t.Response.StatusCode
	}
	return false
}
