// Package m365 binds a Microsoft 365 user in as a target system: the agent's
// own mailbox, its calendar, and the SharePoint sites and OneDrive it is given
// — all through the Microsoft Graph API with an Entra ID app registration.
package m365

import (
	"bytes"
	"context"
	"encoding/json"
	"fmt"
	"io"
	"net/http"
	"net/url"
	"strings"
	"time"

	"github.com/benjaminLedel/covey-plugin-sdk/target"
)

// Client speaks Graph v1.0 for one mailbox.
type Client struct {
	Graph   string // base without /v1.0
	Mailbox string
	Token   string
	HTTP    *http.Client
}

func NewClient(cfg Config, token string) *Client {
	return &Client{
		Graph:   strings.TrimRight(cfg.GraphBase, "/"),
		Mailbox: cfg.Mailbox,
		Token:   token,
		HTTP:    target.Client("m365", 60*time.Second),
	}
}

// user is the path prefix of everything the agent owns.
func (c *Client) user() string {
	return c.Graph + "/v1.0/users/" + url.PathEscape(c.Mailbox)
}

// do sends one request. Ids are asked for as immutable ids: a message moved to
// another folder keeps its id, so a gate entry or a note naming it stays true.
// Bodies come as text, because the agent reads them and a model does not need
// the HTML of a mail client.
func (c *Client) do(ctx context.Context, method, u string, body any, out any) error {
	var rd io.Reader
	if body != nil {
		data, err := json.Marshal(body)
		if err != nil {
			return err
		}
		rd = bytes.NewReader(data)
	}
	req, err := http.NewRequestWithContext(ctx, method, u, rd)
	if err != nil {
		return err
	}
	req.Header.Set("Authorization", "Bearer "+c.Token)
	req.Header.Set("Prefer", `IdType="ImmutableId", outlook.body-content-type="text"`)
	if body != nil {
		req.Header.Set("Content-Type", "application/json")
	}
	resp, err := c.HTTP.Do(req)
	if err != nil {
		return err
	}
	defer resp.Body.Close()
	data, err := io.ReadAll(io.LimitReader(resp.Body, 8<<20))
	if err != nil {
		return err
	}
	if resp.StatusCode < 200 || resp.StatusCode >= 300 {
		return graphError(method, u, resp.StatusCode, data)
	}
	if out != nil && len(data) > 0 {
		return json.Unmarshal(data, out)
	}
	return nil
}

// stream fetches a body without reading it into memory (file contents,
// attachments). A download URL Graph redirects to (SharePoint's tempauth host)
// is followed by the HTTP client; that host carries its own token in the URL,
// and Go drops the Authorization header on a redirect to a different host.
func (c *Client) stream(ctx context.Context, u string) (io.ReadCloser, string, error) {
	req, err := http.NewRequestWithContext(ctx, http.MethodGet, u, nil)
	if err != nil {
		return nil, "", err
	}
	req.Header.Set("Authorization", "Bearer "+c.Token)
	req.Header.Set("Prefer", `IdType="ImmutableId"`)
	resp, err := c.HTTP.Do(req)
	if err != nil {
		return nil, "", err
	}
	if resp.StatusCode < 200 || resp.StatusCode >= 300 {
		data, _ := io.ReadAll(io.LimitReader(resp.Body, 64<<10))
		resp.Body.Close()
		return nil, "", graphError(http.MethodGet, u, resp.StatusCode, data)
	}
	return resp.Body, resp.Header.Get("Content-Type"), nil
}

// put uploads raw bytes (a file's content).
func (c *Client) put(ctx context.Context, u string, data []byte, out any) error {
	req, err := http.NewRequestWithContext(ctx, http.MethodPut, u, bytes.NewReader(data))
	if err != nil {
		return err
	}
	req.Header.Set("Authorization", "Bearer "+c.Token)
	req.Header.Set("Content-Type", "application/octet-stream")
	resp, err := c.HTTP.Do(req)
	if err != nil {
		return err
	}
	defer resp.Body.Close()
	body, err := io.ReadAll(io.LimitReader(resp.Body, 1<<20))
	if err != nil {
		return err
	}
	if resp.StatusCode < 200 || resp.StatusCode >= 300 {
		return graphError(http.MethodPut, u, resp.StatusCode, body)
	}
	if out != nil && len(body) > 0 {
		return json.Unmarshal(body, out)
	}
	return nil
}

// graphError turns a Graph error body into one sentence the agent can act on.
// The two that need explaining are the ones an operator fixes, not the agent.
func graphError(method, u string, status int, data []byte) error {
	var e struct {
		Error struct {
			Code    string `json:"code"`
			Message string `json:"message"`
		} `json:"error"`
	}
	_ = json.Unmarshal(data, &e)
	msg := strings.TrimSpace(e.Error.Code + ": " + e.Error.Message)
	if msg == ":" {
		msg = fmt.Sprintf("%.200s", data)
	}
	where := u
	if p, err := url.Parse(u); err == nil {
		where = p.Path
	}
	switch status {
	case http.StatusUnauthorized:
		return fmt.Errorf("graph %s %s: HTTP 401 — the app registration's token was refused (%s)", method, where, msg)
	case http.StatusForbidden:
		return fmt.Errorf("graph %s %s: HTTP 403 — the app registration lacks the permission for this, or the Exchange scope does not include this mailbox; an operator grants it, retrying does not (%s)", method, where, msg)
	case http.StatusNotFound:
		return fmt.Errorf("graph %s %s: HTTP 404 — no such item (%s)", method, where, msg)
	}
	return fmt.Errorf("graph %s %s: HTTP %d (%s)", method, where, status, msg)
}

// ---------------------------------------------------------------- MAIL

// Address is a mail address with its display name.
type Address struct {
	Name    string `json:"name,omitempty"`
	Address string `json:"address"`
}

type recipient struct {
	EmailAddress Address `json:"emailAddress"`
}

// Message is the part of a Graph message the agent needs.
type Message struct {
	ID               string      `json:"id"`
	Subject          string      `json:"subject"`
	From             *recipient  `json:"from,omitempty"`
	ToRecipients     []recipient `json:"toRecipients,omitempty"`
	CcRecipients     []recipient `json:"ccRecipients,omitempty"`
	ReceivedDateTime string      `json:"receivedDateTime,omitempty"`
	LastModified     string      `json:"lastModifiedDateTime,omitempty"`
	IsRead           bool        `json:"isRead"`
	IsDraft          bool        `json:"isDraft,omitempty"`
	HasAttachments   bool        `json:"hasAttachments"`
	BodyPreview      string      `json:"bodyPreview,omitempty"`
	ConversationID   string      `json:"conversationId,omitempty"`
	ParentFolderID   string      `json:"parentFolderId,omitempty"`
	WebLink          string      `json:"webLink,omitempty"`
	Body             *struct {
		ContentType string `json:"contentType"`
		Content     string `json:"content"`
	} `json:"body,omitempty"`
}

// folderPath names a mail folder: a well-known name (inbox, drafts, sentitems,
// deleteditems, archive, junkemail), or an id, or a display name looked up
// among the top-level folders.
func (c *Client) folderPath(ctx context.Context, folder string) (string, error) {
	f := strings.TrimSpace(folder)
	if f == "" {
		f = "inbox"
	}
	switch strings.ToLower(f) {
	case "inbox", "drafts", "sentitems", "deleteditems", "archive", "junkemail", "outbox":
		return c.user() + "/mailFolders/" + strings.ToLower(f), nil
	}
	if len(f) > 60 && !strings.ContainsAny(f, " /") {
		return c.user() + "/mailFolders/" + url.PathEscape(f), nil
	}
	folders, err := c.MailFolders(ctx)
	if err != nil {
		return "", err
	}
	for _, mf := range folders {
		if strings.EqualFold(mf.DisplayName, f) {
			return c.user() + "/mailFolders/" + url.PathEscape(mf.ID), nil
		}
	}
	return "", fmt.Errorf("no mail folder %q — list_mail_folders names the ones there are", f)
}

// MailFolder is one folder of the mailbox.
type MailFolder struct {
	ID          string `json:"id"`
	DisplayName string `json:"displayName"`
	Unread      int    `json:"unreadItemCount"`
	Total       int    `json:"totalItemCount"`
}

func (c *Client) MailFolders(ctx context.Context) ([]MailFolder, error) {
	var out struct {
		Value []MailFolder `json:"value"`
	}
	err := c.do(ctx, http.MethodGet, c.user()+"/mailFolders?$top=100", nil, &out)
	return out.Value, err
}

const listSelect = "id,subject,from,receivedDateTime,lastModifiedDateTime,isRead,hasAttachments,bodyPreview,conversationId"

// ListMessages lists a folder, newest first.
func (c *Client) ListMessages(ctx context.Context, folder string, unreadOnly bool, search string, limit int) ([]Message, error) {
	base, err := c.folderPath(ctx, folder)
	if err != nil {
		return nil, err
	}
	if limit <= 0 || limit > 50 {
		limit = 20
	}
	q := url.Values{"$top": {fmt.Sprint(limit)}, "$select": {listSelect}}
	if s := strings.TrimSpace(search); s != "" {
		// $search cannot be combined with $orderby; Graph sorts by relevance.
		q.Set("$search", `"`+strings.ReplaceAll(s, `"`, "")+`"`)
	} else {
		q.Set("$orderby", "receivedDateTime desc")
		if unreadOnly {
			// Exchange refuses a $filter whose properties do not lead with the
			// $orderby ones ("InefficientFilter"); the receivedDateTime clause
			// is there for that rule alone.
			q.Set("$filter", "receivedDateTime ge 1900-01-01T00:00:00Z and isRead eq false")
		}
	}
	var out struct {
		Value []Message `json:"value"`
	}
	err = c.do(ctx, http.MethodGet, base+"/messages?"+q.Encode(), nil, &out)
	return out.Value, err
}

func (c *Client) GetMessage(ctx context.Context, id string) (Message, error) {
	var m Message
	err := c.do(ctx, http.MethodGet, c.user()+"/messages/"+url.PathEscape(id), nil, &m)
	return m, err
}

// AttachmentInfo is an attachment without its bytes.
type AttachmentInfo struct {
	ID          string `json:"id"`
	Name        string `json:"name"`
	ContentType string `json:"contentType"`
	Size        int64  `json:"size"`
	IsInline    bool   `json:"isInline"`
	Type        string `json:"@odata.type,omitempty"`
}

func (c *Client) Attachments(ctx context.Context, id string) ([]AttachmentInfo, error) {
	var out struct {
		Value []AttachmentInfo `json:"value"`
	}
	err := c.do(ctx, http.MethodGet, c.user()+"/messages/"+url.PathEscape(id)+"/attachments?$select=id,name,contentType,size,isInline", nil, &out)
	return out.Value, err
}

func (c *Client) AttachmentContent(ctx context.Context, msgID, attID string) (io.ReadCloser, string, error) {
	return c.stream(ctx, c.user()+"/messages/"+url.PathEscape(msgID)+"/attachments/"+url.PathEscape(attID)+"/$value")
}

// fileAttachment is how a sandbox file travels in a message.
type fileAttachment struct {
	Type         string `json:"@odata.type"`
	Name         string `json:"name"`
	ContentType  string `json:"contentType,omitempty"`
	ContentBytes []byte `json:"contentBytes"`
}

// Outgoing is a message the agent writes.
type Outgoing struct {
	To          []string
	Cc          []string
	Subject     string
	Body        string
	Attachments []fileAttachment
}

func toRecipients(addrs []string) []recipient {
	out := make([]recipient, 0, len(addrs))
	for _, a := range addrs {
		if a = strings.TrimSpace(a); a != "" {
			out = append(out, recipient{EmailAddress: Address{Address: a}})
		}
	}
	return out
}

func (o Outgoing) message() map[string]any {
	m := map[string]any{
		"subject":      o.Subject,
		"body":         map[string]string{"contentType": "Text", "content": o.Body},
		"toRecipients": toRecipients(o.To),
	}
	if len(o.Cc) > 0 {
		m["ccRecipients"] = toRecipients(o.Cc)
	}
	if len(o.Attachments) > 0 {
		m["attachments"] = o.Attachments
	}
	return m
}

// Send sends a new message and keeps a copy in Sent Items.
func (c *Client) Send(ctx context.Context, o Outgoing) error {
	return c.do(ctx, http.MethodPost, c.user()+"/sendMail", map[string]any{"message": o.message(), "saveToSentItems": true}, nil)
}

// CreateDraft puts a message into Drafts and returns it — for a person to read
// and send.
func (c *Client) CreateDraft(ctx context.Context, o Outgoing) (Message, error) {
	var m Message
	err := c.do(ctx, http.MethodPost, c.user()+"/messages", o.message(), &m)
	return m, err
}

// Reply answers a message (or all of its recipients), with attachments when
// there are any: those need the two-step draft route, because /reply takes no
// attachments.
func (c *Client) Reply(ctx context.Context, id, body string, all bool, atts []fileAttachment) error {
	verb := "/reply"
	if all {
		verb = "/replyAll"
	}
	base := c.user() + "/messages/" + url.PathEscape(id)
	if len(atts) == 0 {
		return c.do(ctx, http.MethodPost, base+verb, map[string]any{"comment": body}, nil)
	}
	create := "/createReply"
	if all {
		create = "/createReplyAll"
	}
	var draft Message
	if err := c.do(ctx, http.MethodPost, base+create, map[string]any{"comment": body}, &draft); err != nil {
		return err
	}
	dbase := c.user() + "/messages/" + url.PathEscape(draft.ID)
	for _, a := range atts {
		if err := c.do(ctx, http.MethodPost, dbase+"/attachments", a, nil); err != nil {
			return fmt.Errorf("reply draft %s kept in Drafts, attachment %q failed: %w", draft.ID, a.Name, err)
		}
	}
	return c.do(ctx, http.MethodPost, dbase+"/send", nil, nil)
}

func (c *Client) Move(ctx context.Context, id, folder string) (Message, error) {
	dest, err := c.folderPath(ctx, folder)
	if err != nil {
		return Message{}, err
	}
	destID := dest[strings.LastIndex(dest, "/")+1:]
	var m Message
	err = c.do(ctx, http.MethodPost, c.user()+"/messages/"+url.PathEscape(id)+"/move", map[string]string{"destinationId": destID}, &m)
	return m, err
}

func (c *Client) MarkRead(ctx context.Context, id string, read bool) error {
	return c.do(ctx, http.MethodPatch, c.user()+"/messages/"+url.PathEscape(id), map[string]bool{"isRead": read}, nil)
}

// ---------------------------------------------------------------- CALENDAR

// DateTimeZone is Graph's local time with its zone.
type DateTimeZone struct {
	DateTime string `json:"dateTime"`
	TimeZone string `json:"timeZone"`
}

// Event is the part of a calendar event the agent needs.
type Event struct {
	ID       string       `json:"id"`
	Subject  string       `json:"subject"`
	Start    DateTimeZone `json:"start"`
	End      DateTimeZone `json:"end"`
	Location *struct {
		DisplayName string `json:"displayName"`
	} `json:"location,omitempty"`
	Organizer *recipient `json:"organizer,omitempty"`
	Attendees []struct {
		EmailAddress Address `json:"emailAddress"`
		Status       struct {
			Response string `json:"response"`
		} `json:"status"`
	} `json:"attendees,omitempty"`
	IsOnline       bool `json:"isOnlineMeeting,omitempty"`
	ResponseStatus *struct {
		Response string `json:"response"`
	} `json:"responseStatus,omitempty"`
	BodyPreview string `json:"bodyPreview,omitempty"`
	WebLink     string `json:"webLink,omitempty"`
	IsCancelled bool   `json:"isCancelled,omitempty"`
}

// Events lists the calendar view between two instants (recurrences expanded).
func (c *Client) Events(ctx context.Context, from, to, tz string) ([]Event, error) {
	q := url.Values{
		"startDateTime": {from},
		"endDateTime":   {to},
		"$orderby":      {"start/dateTime"},
		"$top":          {"100"},
		"$select":       {"id,subject,start,end,location,organizer,attendees,isOnlineMeeting,responseStatus,bodyPreview,webLink,isCancelled"},
	}
	var out struct {
		Value []Event `json:"value"`
	}
	err := c.doTZ(ctx, c.user()+"/calendarView?"+q.Encode(), tz, &out)
	return out.Value, err
}

// doTZ is a GET with the time zone the agent reads times in.
func (c *Client) doTZ(ctx context.Context, u, tz string, out any) error {
	req, err := http.NewRequestWithContext(ctx, http.MethodGet, u, nil)
	if err != nil {
		return err
	}
	req.Header.Set("Authorization", "Bearer "+c.Token)
	prefer := `IdType="ImmutableId", outlook.body-content-type="text"`
	if tz != "" {
		prefer += `, outlook.timezone="` + strings.ReplaceAll(tz, `"`, "") + `"`
	}
	req.Header.Set("Prefer", prefer)
	resp, err := c.HTTP.Do(req)
	if err != nil {
		return err
	}
	defer resp.Body.Close()
	data, err := io.ReadAll(io.LimitReader(resp.Body, 8<<20))
	if err != nil {
		return err
	}
	if resp.StatusCode < 200 || resp.StatusCode >= 300 {
		return graphError(http.MethodGet, u, resp.StatusCode, data)
	}
	return json.Unmarshal(data, out)
}

// NewEvent is an event the agent creates.
type NewEvent struct {
	Subject   string
	Start     string
	End       string
	TimeZone  string
	Attendees []string
	Body      string
	Location  string
	Online    bool
}

func (c *Client) CreateEvent(ctx context.Context, e NewEvent) (Event, error) {
	tz := e.TimeZone
	if tz == "" {
		tz = "UTC"
	}
	body := map[string]any{
		"subject": e.Subject,
		"start":   DateTimeZone{DateTime: e.Start, TimeZone: tz},
		"end":     DateTimeZone{DateTime: e.End, TimeZone: tz},
	}
	if e.Body != "" {
		body["body"] = map[string]string{"contentType": "Text", "content": e.Body}
	}
	if e.Location != "" {
		body["location"] = map[string]string{"displayName": e.Location}
	}
	if len(e.Attendees) > 0 {
		att := make([]map[string]any, 0, len(e.Attendees))
		for _, a := range e.Attendees {
			if a = strings.TrimSpace(a); a != "" {
				att = append(att, map[string]any{"emailAddress": Address{Address: a}, "type": "required"})
			}
		}
		body["attendees"] = att
	}
	if e.Online {
		body["isOnlineMeeting"] = true
		body["onlineMeetingProvider"] = "teamsForBusiness"
	}
	var out Event
	err := c.do(ctx, http.MethodPost, c.user()+"/events", body, &out)
	return out, err
}

// Respond accepts, declines or tentatively accepts an invitation.
func (c *Client) Respond(ctx context.Context, id, response, comment string) error {
	verb := map[string]string{"accept": "/accept", "decline": "/decline", "tentative": "/tentativelyAccept"}[response]
	if verb == "" {
		return fmt.Errorf("response must be accept, decline or tentative, not %q", response)
	}
	return c.do(ctx, http.MethodPost, c.user()+"/events/"+url.PathEscape(id)+verb, map[string]any{"comment": comment, "sendResponse": true}, nil)
}

func (c *Client) CancelEvent(ctx context.Context, id, comment string) error {
	return c.do(ctx, http.MethodPost, c.user()+"/events/"+url.PathEscape(id)+"/cancel", map[string]string{"comment": comment}, nil)
}

// Schedule is the free/busy picture of one person.
type Schedule struct {
	ScheduleID       string `json:"scheduleId"`
	AvailabilityView string `json:"availabilityView"`
	Items            []struct {
		Status string       `json:"status"`
		Start  DateTimeZone `json:"start"`
		End    DateTimeZone `json:"end"`
	} `json:"scheduleItems"`
	Error *struct {
		Message string `json:"message"`
	} `json:"error,omitempty"`
}

// FreeBusy asks for the busy times of colleagues (and the agent itself).
func (c *Client) FreeBusy(ctx context.Context, emails []string, from, to, tz string) ([]Schedule, error) {
	if tz == "" {
		tz = "UTC"
	}
	body := map[string]any{
		"schedules":                emails,
		"startTime":                DateTimeZone{DateTime: from, TimeZone: tz},
		"endTime":                  DateTimeZone{DateTime: to, TimeZone: tz},
		"availabilityViewInterval": 30,
	}
	var out struct {
		Value []Schedule `json:"value"`
	}
	err := c.do(ctx, http.MethodPost, c.user()+"/calendar/getSchedule", body, &out)
	return out.Value, err
}

// ---------------------------------------------------------------- FILES

// Site is a SharePoint site.
type Site struct {
	ID          string `json:"id"`
	Name        string `json:"name"`
	DisplayName string `json:"displayName"`
	WebURL      string `json:"webUrl"`
}

func (c *Client) Sites(ctx context.Context, search string) ([]Site, error) {
	s := strings.TrimSpace(search)
	if s == "" {
		s = "*"
	}
	var out struct {
		Value []Site `json:"value"`
	}
	err := c.do(ctx, http.MethodGet, c.Graph+"/v1.0/sites?search="+url.QueryEscape(s)+"&$top=50", nil, &out)
	return out.Value, err
}

// Drive is a document library (or a OneDrive).
type Drive struct {
	ID        string `json:"id"`
	Name      string `json:"name"`
	DriveType string `json:"driveType"`
	WebURL    string `json:"webUrl"`
}

func (c *Client) Drives(ctx context.Context, siteID string) ([]Drive, error) {
	var out struct {
		Value []Drive `json:"value"`
	}
	err := c.do(ctx, http.MethodGet, c.Graph+"/v1.0/sites/"+url.PathEscape(siteID)+"/drives", nil, &out)
	return out.Value, err
}

// drivePath is the root of a drive: "me" (or empty) is the agent's own
// OneDrive, anything else a drive id from list_drives.
func (c *Client) drivePath(driveID string) string {
	d := strings.TrimSpace(driveID)
	if d == "" || strings.EqualFold(d, "me") {
		return c.user() + "/drive"
	}
	return c.Graph + "/v1.0/drives/" + url.PathEscape(d)
}

// itemPath addresses an item by its path below the drive root.
func (c *Client) itemPath(driveID, p string) string {
	p = strings.Trim(p, "/")
	if p == "" {
		return c.drivePath(driveID) + "/root"
	}
	segs := strings.Split(p, "/")
	for i, s := range segs {
		segs[i] = url.PathEscape(s)
	}
	return c.drivePath(driveID) + "/root:/" + strings.Join(segs, "/") + ":"
}

// Item is a file or folder.
type Item struct {
	ID           string `json:"id"`
	Name         string `json:"name"`
	Size         int64  `json:"size"`
	WebURL       string `json:"webUrl"`
	LastModified string `json:"lastModifiedDateTime"`
	Folder       *struct {
		ChildCount int `json:"childCount"`
	} `json:"folder,omitempty"`
	File *struct {
		MimeType string `json:"mimeType"`
	} `json:"file,omitempty"`
	ParentReference *struct {
		DriveID string `json:"driveId"`
		Path    string `json:"path"`
	} `json:"parentReference,omitempty"`
}

func (c *Client) Children(ctx context.Context, driveID, p string) ([]Item, bool, error) {
	var out struct {
		Value []Item `json:"value"`
		Next  string `json:"@odata.nextLink"`
	}
	base := c.itemPath(driveID, p)
	err := c.do(ctx, http.MethodGet, base+"/children?$top=200&$select=id,name,size,webUrl,lastModifiedDateTime,folder,file", nil, &out)
	return out.Value, out.Next != "", err
}

func (c *Client) SearchDrive(ctx context.Context, driveID, query string, limit int) ([]Item, error) {
	if limit <= 0 || limit > 50 {
		limit = 25
	}
	q := strings.ReplaceAll(strings.TrimSpace(query), "'", "''")
	var out struct {
		Value []Item `json:"value"`
	}
	err := c.do(ctx, http.MethodGet, c.drivePath(driveID)+"/root/search(q='"+url.PathEscape(q)+"')?$top="+fmt.Sprint(limit)+"&$select=id,name,size,webUrl,lastModifiedDateTime,folder,file,parentReference", nil, &out)
	return out.Value, err
}

func (c *Client) Download(ctx context.Context, driveID, p string) (io.ReadCloser, string, error) {
	return c.stream(ctx, c.itemPath(driveID, p)+"/content")
}

func (c *Client) Upload(ctx context.Context, driveID, p string, data []byte) (Item, error) {
	var it Item
	err := c.put(ctx, c.itemPath(driveID, p)+"/content", data, &it)
	return it, err
}

// Mkdir creates a folder, and the missing ones above it.
func (c *Client) Mkdir(ctx context.Context, driveID, p string) (Item, error) {
	p = strings.Trim(p, "/")
	if p == "" {
		return Item{}, fmt.Errorf("path missing")
	}
	var last Item
	parent := ""
	for _, seg := range strings.Split(p, "/") {
		body := map[string]any{"name": seg, "folder": map[string]any{}, "@microsoft.graph.conflictBehavior": "fail"}
		err := c.do(ctx, http.MethodPost, c.itemPath(driveID, parent)+"/children", body, &last)
		if err != nil && !strings.Contains(err.Error(), "HTTP 409") {
			return Item{}, err
		}
		if parent == "" {
			parent = seg
		} else {
			parent += "/" + seg
		}
	}
	return last, nil
}
