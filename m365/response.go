package m365

import (
	"encoding/json"
	"io"
	"net/http"
	"strconv"
	"strings"
	"time"
)

type Response struct {
	*http.Response

	Content []byte

	NextPage  int
	PrevPage  int
	FirstPage int
	LastPage  int

	NextLink string
}

type Rate struct {
	Limit     int
	Remaining int
	Reset     time.Time
	Used      int
}

func newResponse(resp *http.Response) *Response {
	r := &Response{
		Response: resp,
	}
	r.parsePagination()
	return r
}

func (r *Response) parsePagination() {
	linkHeader := r.Response.Header.Get("Link")
	if linkHeader == "" {
		return
	}

	links := strings.Split(linkHeader, ",")
	for _, link := range links {
		link = strings.TrimSpace(link)
		parts := strings.Split(link, ";")
		if len(parts) != 2 {
			continue
		}
		url := strings.Trim(parts[0], "<>")
		rel := strings.Trim(parts[1], "\"")

		page := 0
		if strings.Contains(rel, "next") {
			r.NextLink = url
		}
		if idx := strings.Index(rel, "page="); idx != -1 {
			page, _ = strconv.Atoi(rel[idx+5:])
		}

		switch {
		case strings.Contains(rel, "next"):
			r.NextPage = page
		case strings.Contains(rel, "prev"):
			r.PrevPage = page
		case strings.Contains(rel, "first"):
			r.FirstPage = page
		case strings.Contains(rel, "last"):
			r.LastPage = page
		}
	}
}

func (r *Response) Decode(v interface{}) error {
	if r.Content == nil {
		body, err := io.ReadAll(r.Response.Body)
		if err != nil {
			return err
		}
		r.Content = body
	}
	return json.Unmarshal(r.Content, v)
}

func (r *Response) GetBody() ([]byte, error) {
	if r.Content != nil {
		return r.Content, nil
	}
	return io.ReadAll(r.Response.Body)
}

func (r *Response) HasNextPage() bool {
	return r.NextLink != "" || r.NextPage > 0
}

type ODataError struct {
	Error *GraphError `json:"error"`
}

type ListResponse struct {
	Value    json.RawMessage `json:"value"`
	NextLink string          `json:"@odata.nextLink"`
}

func parseListResponse(body []byte) (json.RawMessage, string, error) {
	var resp ListResponse
	if err := json.Unmarshal(body, &resp); err != nil {
		return nil, "", err
	}
	return resp.Value, resp.NextLink, nil
}

func parseInt(s string, fallback int) int {
	if s == "" {
		return fallback
	}
	v, err := strconv.Atoi(s)
	if err != nil {
		return fallback
	}
	return v
}
