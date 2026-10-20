package m365

import (
	"context"
	"encoding/json"
	"fmt"
	"io"
	"mime"
	"net/url"
	"os"
	"path/filepath"
	"strings"
	"unicode/utf8"

	"github.com/benjaminLedel/covey-plugin-sdk/target"
)

// System binds a Microsoft 365 user in as a target system: the agent's own
// mailbox, calendar and OneDrive, and the SharePoint sites the app may read.
// There is no webhook entry — Graph change notifications need a publicly
// validated subscription that has to be renewed every few days; the intake
// runs by HEARTBEAT.md polling (nur-wenn: m365), as with the e-mail plugin.
type System struct{}

func init() {
	target.Register(target.Descriptor{
		Env:         []string{"COVEY_M365_UPLOAD_MAX_MB", "COVEY_M365_ATTACHMENT_MAX_MB"},
		Name:        "m365",
		Label:       "Microsoft 365",
		Description: "The agent's own Microsoft 365 user through Microsoft Graph: its mailbox (list, read, reply, send, drafts, move, attachments into the sandbox), its calendar (events, invitations, free/busy of colleagues), its OneDrive and the SharePoint sites the app may reach (find sites and libraries, search, download into the sandbox, upload). Auth through an Entra ID app registration (client credentials), secrets m365_url (the agent's address) + m365_token (tenant:client:secret). Intake through HEARTBEAT.md (nur-wenn: m365 = unread mail in the inbox).",
		Kind:        "builtin",
		Category:    target.CategoryComms,
		Scopes:      []string{"read", "write"},
		System:      System{},
		SetupDoc: `1. Give the agent a Microsoft 365 user of its own (a licensed user with a
   mailbox, e.g. agent@contoso.com). Its OneDrive comes with the licence.
2. In Entra ID (App registrations) register an app for the agent and add the
   APPLICATION permissions of Microsoft Graph, then grant admin consent:
     Mail.ReadWrite, Mail.Send, Calendars.ReadWrite   — mailbox and calendar
     Files.ReadWrite.All                               — OneDrive and libraries
     Sites.ReadWrite.All (or Sites.Selected, see 4)    — SharePoint sites
   Create a client secret (Certificates & secrets).
3. Fence the app into the agent's mailbox. Application permissions hold
   tenant-wide; without this step the app could read every mailbox. In
   Exchange Online PowerShell (RBAC for Applications):
     New-ManagementScope -Name "covey-agent" -RecipientRestrictionFilter "PrimarySmtpAddress -eq 'agent@contoso.com'"
     New-ServicePrincipal -AppId <client-id> -ObjectId <enterprise-app-object-id> -DisplayName "covey agent"
     New-ManagementRoleAssignment -App <client-id> -Role "Application Mail.ReadWrite" -CustomResourceScope "covey-agent"
     (the same for "Application Mail.Send" and "Application Calendars.ReadWrite")
   and remove Mail.*/Calendars.* from the Graph permissions again — the
   Exchange role assignment carries them, scoped. Test-ServicePrincipalAuthorization
   shows what the app can reach.
4. Least privilege for SharePoint: Sites.Selected instead of Sites.ReadWrite.All,
   then grant the app write access to each site it is to work in through Graph
   (POST /sites/{site-id}/permissions). list_sites then shows only those.
5. Store under Secrets and assign to the agent:
     m365_url   = agent@contoso.com
     m365_token = tenant-id:client-id:client-secret
6. Enable in the agent's ACCESS.md:
     - system: m365 scope: read,write
7. Egress: graph.microsoft.com, login.microsoftonline.com and *.sharepoint.com
   (file downloads are redirected there).
8. Optional intake by heartbeat — in the agent's HEARTBEAT.md:
     - alle: 10m nur-wenn: m365 titel: Look after the mailbox aufgabe: Read
       the unread mail (list_mail unread:true) and work on it by the playbook.`,
	})
}

func (System) Name() string { return "m365" }

// ActionSubject: every action is its own guard-rail subject, so sending mail
// or deleting nothing-yet can be ruled on apart from reading.
func (System) ActionSubject(action string, _ json.RawMessage) string {
	return "m365:" + action
}

// params is the union of what any action takes; the agent sends one flat object.
type params struct {
	ID           string   `json:"id"`
	AttachmentID string   `json:"attachment_id"`
	Folder       string   `json:"folder"`
	Unread       bool     `json:"unread"`
	Read         *bool    `json:"read"`
	Search       string   `json:"search"`
	Query        string   `json:"query"`
	Limit        int      `json:"limit"`
	To           any      `json:"to"`
	Cc           []string `json:"cc"`
	Subject      string   `json:"subject"`
	Body         string   `json:"body"`
	Attachments  []string `json:"attachments"`
	ReplyAll     bool     `json:"reply_all"`
	From         string   `json:"from"`
	Until        string   `json:"until"`
	Start        string   `json:"start"`
	End          string   `json:"end"`
	TimeZone     string   `json:"timezone"`
	Attendees    []string `json:"attendees"`
	Emails       []string `json:"emails"`
	Location     string   `json:"location"`
	Online       bool     `json:"online"`
	Response     string   `json:"response"`
	Comment      string   `json:"comment"`
	SiteID       string   `json:"site_id"`
	DriveID      string   `json:"drive_id"`
	Path         string   `json:"path"`
	Content      string   `json:"content"`
	Save         string   `json:"save"`
}

// toList reads "to" as one address or a list of them.
func (p params) toList() []string {
	switch v := p.To.(type) {
	case string:
		return splitAddrs(v)
	case []any:
		var out []string
		for _, x := range v {
			if s, ok := x.(string); ok {
				out = append(out, splitAddrs(s)...)
			}
		}
		return out
	}
	return nil
}

func splitAddrs(s string) []string {
	var out []string
	for _, a := range strings.FieldsFunc(s, func(r rune) bool { return r == ',' || r == ';' }) {
		if a = strings.TrimSpace(a); a != "" {
			out = append(out, a)
		}
	}
	return out
}

type action func(ctx context.Context, c *Client, in params) (any, error)

var actions = map[string]action{
	// --- mail
	"list_mail_folders": func(ctx context.Context, c *Client, in params) (any, error) {
		return c.MailFolders(ctx)
	},
	"list_mail": func(ctx context.Context, c *Client, in params) (any, error) {
		msgs, err := c.ListMessages(ctx, in.Folder, in.Unread, in.Search, in.Limit)
		if err != nil {
			return nil, err
		}
		rows := make([]map[string]any, 0, len(msgs))
		for _, m := range msgs {
			rows = append(rows, mailRow(m))
		}
		return map[string]any{"folder": orDefault(in.Folder, "inbox"), "messages": rows}, nil
	},
	"get_mail": func(ctx context.Context, c *Client, in params) (any, error) {
		if in.ID == "" {
			return nil, fmt.Errorf("id missing — list_mail names the messages")
		}
		m, err := c.GetMessage(ctx, in.ID)
		if err != nil {
			return nil, err
		}
		out := mailRow(m)
		out["to"], out["cc"] = addrs(m.ToRecipients), addrs(m.CcRecipients)
		if m.Body != nil {
			out["body"] = m.Body.Content
		}
		out["web_link"] = m.WebLink
		if m.HasAttachments {
			atts, err := c.Attachments(ctx, in.ID)
			if err != nil {
				return nil, err
			}
			out["attachments"] = atts
		}
		return out, nil
	},
	"get_mail_attachment": func(ctx context.Context, c *Client, in params) (any, error) {
		if in.ID == "" || in.AttachmentID == "" {
			return nil, fmt.Errorf("id and attachment_id needed — get_mail names the attachments")
		}
		atts, err := c.Attachments(ctx, in.ID)
		if err != nil {
			return nil, err
		}
		var found *AttachmentInfo
		for i := range atts {
			if atts[i].ID == in.AttachmentID {
				found = &atts[i]
			}
		}
		if found == nil {
			return nil, fmt.Errorf("no attachment %q on this message", in.AttachmentID)
		}
		if strings.Contains(found.Type, "itemAttachment") || strings.Contains(found.Type, "referenceAttachment") {
			return nil, fmt.Errorf("attachment %q is an embedded item or a cloud link, not a file — open the link in the mail body instead", found.Name)
		}
		workdir := target.Workdir(ctx)
		if workdir == "" {
			return nil, fmt.Errorf("no sandbox working directory")
		}
		body, ctype, err := c.AttachmentContent(ctx, in.ID, in.AttachmentID)
		if err != nil {
			return nil, err
		}
		defer body.Close()
		if found.ContentType != "" {
			ctype = found.ContentType
		}
		saved, err := target.StoreStream(workdir, "mail-attachments", found.Name, body, attachmentMaxBytes(), ctype)
		if err != nil {
			return nil, err
		}
		return map[string]any{"path": saved.Path, "name": saved.FileName, "content_type": saved.ContentType, "bytes": saved.Bytes, "hint": target.Hint(saved.Path, saved.ContentType)}, nil
	},
	"send_mail": func(ctx context.Context, c *Client, in params) (any, error) {
		o, err := outgoing(ctx, in)
		if err != nil {
			return nil, err
		}
		if len(o.To) == 0 {
			return nil, fmt.Errorf("to missing")
		}
		if err := c.Send(ctx, o); err != nil {
			return nil, err
		}
		return map[string]any{"sent": true, "to": o.To, "subject": o.Subject, "attachments": len(o.Attachments)}, nil
	},
	"create_draft": func(ctx context.Context, c *Client, in params) (any, error) {
		o, err := outgoing(ctx, in)
		if err != nil {
			return nil, err
		}
		m, err := c.CreateDraft(ctx, o)
		if err != nil {
			return nil, err
		}
		return map[string]any{"draft_id": m.ID, "web_link": m.WebLink, "hint": "The draft lies in Drafts; a person opens it with web_link, checks and sends it."}, nil
	},
	"reply_mail": func(ctx context.Context, c *Client, in params) (any, error) {
		if in.ID == "" || strings.TrimSpace(in.Body) == "" {
			return nil, fmt.Errorf("id and body needed")
		}
		atts, err := loadAttachments(ctx, in.Attachments)
		if err != nil {
			return nil, err
		}
		if err := c.Reply(ctx, in.ID, in.Body, in.ReplyAll, atts); err != nil {
			return nil, err
		}
		return map[string]any{"replied": true, "reply_all": in.ReplyAll, "attachments": len(atts)}, nil
	},
	"move_mail": func(ctx context.Context, c *Client, in params) (any, error) {
		if in.ID == "" || in.Folder == "" {
			return nil, fmt.Errorf("id and folder needed")
		}
		m, err := c.Move(ctx, in.ID, in.Folder)
		if err != nil {
			return nil, err
		}
		return map[string]any{"moved": true, "id": m.ID, "folder": in.Folder}, nil
	},
	"mark_read": func(ctx context.Context, c *Client, in params) (any, error) {
		if in.ID == "" {
			return nil, fmt.Errorf("id missing")
		}
		read := true
		if in.Read != nil {
			read = *in.Read
		}
		if err := c.MarkRead(ctx, in.ID, read); err != nil {
			return nil, err
		}
		return map[string]any{"id": in.ID, "read": read}, nil
	},
	// --- calendar
	"list_events": func(ctx context.Context, c *Client, in params) (any, error) {
		if in.From == "" || in.Until == "" {
			return nil, fmt.Errorf("from and until needed (ISO 8601, e.g. 2026-10-21T00:00:00)")
		}
		return c.Events(ctx, in.From, in.Until, in.TimeZone)
	},
	"create_event": func(ctx context.Context, c *Client, in params) (any, error) {
		if in.Subject == "" || in.Start == "" || in.End == "" {
			return nil, fmt.Errorf("subject, start and end needed")
		}
		ev, err := c.CreateEvent(ctx, NewEvent{Subject: in.Subject, Start: in.Start, End: in.End, TimeZone: in.TimeZone,
			Attendees: in.Attendees, Body: in.Body, Location: in.Location, Online: in.Online})
		if err != nil {
			return nil, err
		}
		return map[string]any{"id": ev.ID, "web_link": ev.WebLink, "invited": len(in.Attendees)}, nil
	},
	"respond_event": func(ctx context.Context, c *Client, in params) (any, error) {
		if in.ID == "" {
			return nil, fmt.Errorf("id missing")
		}
		if err := c.Respond(ctx, in.ID, in.Response, in.Comment); err != nil {
			return nil, err
		}
		return map[string]any{"id": in.ID, "response": in.Response}, nil
	},
	"cancel_event": func(ctx context.Context, c *Client, in params) (any, error) {
		if in.ID == "" {
			return nil, fmt.Errorf("id missing")
		}
		if err := c.CancelEvent(ctx, in.ID, in.Comment); err != nil {
			return nil, err
		}
		return map[string]any{"id": in.ID, "cancelled": true}, nil
	},
	"free_busy": func(ctx context.Context, c *Client, in params) (any, error) {
		if len(in.Emails) == 0 || in.From == "" || in.Until == "" {
			return nil, fmt.Errorf("emails, from and until needed")
		}
		return c.FreeBusy(ctx, in.Emails, in.From, in.Until, in.TimeZone)
	},
	// --- files
	"list_sites": func(ctx context.Context, c *Client, in params) (any, error) {
		return c.Sites(ctx, in.Search)
	},
	"list_drives": func(ctx context.Context, c *Client, in params) (any, error) {
		if in.SiteID == "" {
			return nil, fmt.Errorf("site_id missing — list_sites names the sites")
		}
		return c.Drives(ctx, in.SiteID)
	},
	"list_files": func(ctx context.Context, c *Client, in params) (any, error) {
		p, err := cleanRemotePath(in.Path)
		if err != nil {
			return nil, err
		}
		items, more, err := c.Children(ctx, in.DriveID, p)
		if err != nil {
			return nil, err
		}
		out := map[string]any{"drive_id": orDefault(in.DriveID, "me"), "path": p, "items": itemRows(items)}
		if more {
			out["truncated"] = true
			out["hint"] = "More than 200 entries — go one folder deeper with path, or use search_files."
		}
		return out, nil
	},
	"search_files": func(ctx context.Context, c *Client, in params) (any, error) {
		if strings.TrimSpace(in.Query) == "" {
			return nil, fmt.Errorf("query missing")
		}
		items, err := c.SearchDrive(ctx, in.DriveID, in.Query, in.Limit)
		if err != nil {
			return nil, err
		}
		return map[string]any{"drive_id": orDefault(in.DriveID, "me"), "items": itemRows(items)}, nil
	},
	"read_file": func(ctx context.Context, c *Client, in params) (any, error) {
		p, err := cleanRemotePath(in.Path)
		if err != nil || p == "" {
			return nil, fmt.Errorf("path missing")
		}
		body, _, err := c.Download(ctx, in.DriveID, p)
		if err != nil {
			return nil, err
		}
		defer body.Close()
		data, err := io.ReadAll(io.LimitReader(body, readMaxBytes+1))
		if err != nil {
			return nil, err
		}
		if len(data) > readMaxBytes || !utf8.Valid(data) {
			return nil, fmt.Errorf("%q is binary or over 1 MB — fetch it with download_file and open it in the sandbox", p)
		}
		return map[string]any{"path": p, "size": len(data), "content": string(data)}, nil
	},
	"download_file": func(ctx context.Context, c *Client, in params) (any, error) {
		p, err := cleanRemotePath(in.Path)
		if err != nil || p == "" {
			return nil, fmt.Errorf("path missing")
		}
		dest := in.Save
		if dest == "" {
			dest = filepath.Join("m365", filepath.FromSlash(p))
		}
		local, err := localPath(ctx, dest)
		if err != nil {
			return nil, err
		}
		body, _, err := c.Download(ctx, in.DriveID, p)
		if err != nil {
			return nil, err
		}
		defer body.Close()
		if err := os.MkdirAll(filepath.Dir(local), 0o755); err != nil {
			return nil, err
		}
		f, err := os.Create(local)
		if err != nil {
			return nil, err
		}
		n, err := io.Copy(f, io.LimitReader(body, uploadMaxBytes()+1))
		if cerr := f.Close(); err == nil {
			err = cerr
		}
		if err != nil {
			return nil, err
		}
		return map[string]any{"path": local, "size": n,
			"hint": "The file is in your sandbox now. Office files: python3 with python-docx / openpyxl / python-pptx reads and edits them; put the result back with upload_file."}, nil
	},
	"upload_file": func(ctx context.Context, c *Client, in params) (any, error) {
		local, err := localPath(ctx, in.From)
		if err != nil {
			return nil, err
		}
		st, err := os.Stat(local)
		if err != nil {
			return nil, fmt.Errorf("local file: %w", err)
		}
		if st.IsDir() {
			return nil, fmt.Errorf("%q is a directory — upload_file transfers single files", in.From)
		}
		if st.Size() > uploadMaxBytes() {
			return nil, fmt.Errorf("%q is %d bytes, over the limit of %d", in.From, st.Size(), uploadMaxBytes())
		}
		p, err := cleanRemotePath(in.Path)
		if err != nil {
			return nil, err
		}
		if p == "" {
			p = filepath.Base(local)
		}
		data, err := os.ReadFile(local)
		if err != nil {
			return nil, err
		}
		it, err := c.Upload(ctx, in.DriveID, p, data)
		if err != nil {
			return nil, err
		}
		return map[string]any{"uploaded": p, "id": it.ID, "size": it.Size, "web_url": it.WebURL}, nil
	},
	"write_file": func(ctx context.Context, c *Client, in params) (any, error) {
		p, err := cleanRemotePath(in.Path)
		if err != nil || p == "" {
			return nil, fmt.Errorf("path missing")
		}
		it, err := c.Upload(ctx, in.DriveID, p, []byte(in.Content))
		if err != nil {
			return nil, err
		}
		return map[string]any{"written": p, "id": it.ID, "size": it.Size, "web_url": it.WebURL}, nil
	},
	"create_folder": func(ctx context.Context, c *Client, in params) (any, error) {
		p, err := cleanRemotePath(in.Path)
		if err != nil {
			return nil, err
		}
		if _, err := c.Mkdir(ctx, in.DriveID, p); err != nil {
			return nil, err
		}
		return map[string]any{"folder": p}, nil
	},
}

// Execute runs one action.
func (System) Execute(ctx context.Context, name string, raw json.RawMessage, cred target.Credential) (any, error) {
	fn, ok := actions[strings.TrimSpace(name)]
	if !ok {
		return nil, fmt.Errorf("unknown action %q", strings.TrimSpace(name))
	}
	c, err := clientFor(ctx, cred)
	if err != nil {
		return nil, err
	}
	var in params
	if len(raw) > 0 {
		if err := json.Unmarshal(raw, &in); err != nil {
			return nil, fmt.Errorf("params: %w", err)
		}
	}
	return fn(ctx, c, in)
}

func clientFor(ctx context.Context, cred target.Credential) (*Client, error) {
	cfg, err := ParseConfig(cred.BaseURL, cred.Token)
	if err != nil {
		return nil, err
	}
	tok, err := cfg.AccessToken(ctx)
	if err != nil {
		return nil, err
	}
	return NewClient(cfg, tok), nil
}

// Probe (target.Prober) is the connection test: it reads the mailbox's inbox,
// the one thing every agent of this plugin needs. A 403 here is almost always
// the Exchange scope, and the error says so.
func (System) Probe(ctx context.Context, cred target.Credential) (string, error) {
	c, err := clientFor(ctx, cred)
	if err != nil {
		return "", err
	}
	folders, err := c.MailFolders(ctx)
	if err != nil {
		return "", err
	}
	for _, f := range folders {
		if strings.EqualFold(f.DisplayName, "inbox") || strings.EqualFold(f.DisplayName, "posteingang") {
			return fmt.Sprintf("mailbox %s reachable: %d unread of %d in the inbox", c.Mailbox, f.Unread, f.Total), nil
		}
	}
	return fmt.Sprintf("mailbox %s reachable: %d folders", c.Mailbox, len(folders)), nil
}

// HasWork / HasWorkSigned (target.SignedWorkChecker): unread mail in the
// inbox. One entry per message, key "mail:<immutable id>", version its
// lastModifiedDateTime — so a workflow on nur-wenn: m365 runs one case per
// mail, and an agent that read a mail and left it unread is not woken again
// by the same state.
func (System) HasWork(ctx context.Context, cred target.Credential) (bool, error) {
	has, _, err := System{}.HasWorkSigned(ctx, cred, "")
	return has, err
}

func (System) HasWorkSigned(ctx context.Context, cred target.Credential, kind string) (bool, string, error) {
	folder := strings.TrimSpace(kind)
	c, err := clientFor(ctx, cred)
	if err != nil {
		return false, "", err
	}
	msgs, err := c.ListMessages(ctx, folder, true, "", 50)
	if err != nil {
		return false, "", err
	}
	entries := make([]target.WorkEntry, 0, len(msgs))
	for _, m := range msgs {
		v := m.LastModified
		if v == "" {
			v = m.ReceivedDateTime
		}
		// Immutable ids are base64url and carry neither "," nor "@"; a
		// version never does. Both are what WorkSignature requires.
		entries = append(entries, target.WorkEntry{Key: "mail:" + m.ID, Version: v})
	}
	if len(entries) == 0 {
		return false, "", nil
	}
	return true, target.WorkSignature(entries), nil
}

// SignsWorkItems (target.ItemSigner): one entry per unread mail.
func (System) SignsWorkItems() bool { return true }

// WritesWorkSignature (target.SignatureWriter): the actions that can take a mail
// out of the unread set or change it. A new one that does must be added here,
// or the agent's own move counts as foreign activity and wakes it once more.
func (System) WritesWorkSignature(subject string) bool {
	switch strings.TrimPrefix(subject, "m365:") {
	case "mark_read", "move_mail", "reply_mail":
		return true
	}
	return false
}

// ---------------------------------------------------------------- helpers

const readMaxBytes = 1 << 20

func uploadMaxBytes() int64 {
	return target.MaxBytesFromEnv("COVEY_M365_UPLOAD_MAX_MB", 250, 250)
}

func attachmentMaxBytes() int64 {
	return target.MaxBytesFromEnv("COVEY_M365_ATTACHMENT_MAX_MB", 25, 150)
}

// outgoingAttachmentMax: a file attachment travels inline in the request, and
// Graph refuses a request over 4 MB. Larger files go to OneDrive and the mail
// carries the link.
const outgoingAttachmentMax = 3 << 20

func outgoing(ctx context.Context, in params) (Outgoing, error) {
	atts, err := loadAttachments(ctx, in.Attachments)
	if err != nil {
		return Outgoing{}, err
	}
	return Outgoing{To: in.toList(), Cc: in.Cc, Subject: in.Subject, Body: in.Body, Attachments: atts}, nil
}

func loadAttachments(ctx context.Context, paths []string) ([]fileAttachment, error) {
	var out []fileAttachment
	var total int64
	for _, p := range paths {
		local, err := localPath(ctx, p)
		if err != nil {
			return nil, err
		}
		data, err := os.ReadFile(local)
		if err != nil {
			return nil, fmt.Errorf("attachment %q: %w", p, err)
		}
		total += int64(len(data))
		if total > outgoingAttachmentMax {
			return nil, fmt.Errorf("attachments over 3 MB together — upload the file with upload_file and put its web_url into the mail instead")
		}
		ctype := mime.TypeByExtension(strings.ToLower(filepath.Ext(local)))
		out = append(out, fileAttachment{Type: "#microsoft.graph.fileAttachment", Name: filepath.Base(local), ContentType: ctype, ContentBytes: data})
	}
	return out, nil
}

// localPath resolves a sandbox path against the working directory, without a
// way out through ".." or an absolute path elsewhere.
func localPath(ctx context.Context, p string) (string, error) {
	workdir := target.Workdir(ctx)
	if workdir == "" {
		return "", fmt.Errorf("no sandbox working directory")
	}
	p = strings.TrimSpace(p)
	if p == "" {
		return "", fmt.Errorf("local path missing")
	}
	if !filepath.IsAbs(p) {
		p = filepath.Join(workdir, p)
	}
	resolved := filepath.Clean(p)
	if resolved != workdir && !strings.HasPrefix(resolved, workdir+string(filepath.Separator)) {
		return "", fmt.Errorf("path %q lies outside the sandbox working directory", p)
	}
	return resolved, nil
}

// cleanRemotePath normalises a path inside a drive and refuses a way up.
func cleanRemotePath(p string) (string, error) {
	p = strings.Trim(strings.ReplaceAll(strings.TrimSpace(p), "\\", "/"), "/")
	if p == "" {
		return "", nil
	}
	for _, seg := range strings.Split(p, "/") {
		if seg == ".." || seg == "." || seg == "" {
			return "", fmt.Errorf("path %q: no empty, . or .. segments", p)
		}
	}
	if u, err := url.PathUnescape(p); err == nil && u != p && strings.Contains(u, "..") {
		return "", fmt.Errorf("path %q: no escaped way up", p)
	}
	return p, nil
}

func mailRow(m Message) map[string]any {
	row := map[string]any{
		"id": m.ID, "subject": m.Subject, "received": m.ReceivedDateTime,
		"read": m.IsRead, "attachments": m.HasAttachments, "preview": m.BodyPreview,
		"conversation_id": m.ConversationID,
	}
	if m.From != nil {
		row["from"] = m.From.EmailAddress
	}
	return row
}

func addrs(rs []recipient) []Address {
	out := make([]Address, 0, len(rs))
	for _, r := range rs {
		out = append(out, r.EmailAddress)
	}
	return out
}

func itemRows(items []Item) []map[string]any {
	out := make([]map[string]any, 0, len(items))
	for _, it := range items {
		row := map[string]any{"name": it.Name, "id": it.ID, "size": it.Size, "modified": it.LastModified, "web_url": it.WebURL}
		if it.Folder != nil {
			row["folder"] = true
		}
		if it.ParentReference != nil {
			row["drive_id"] = it.ParentReference.DriveID
			if _, rest, ok := strings.Cut(it.ParentReference.Path, "root:"); ok {
				row["path"] = strings.TrimPrefix(rest+"/"+it.Name, "/")
			}
		}
		out = append(out, row)
	}
	return out
}

func orDefault(s, d string) string {
	if strings.TrimSpace(s) == "" {
		return d
	}
	return s
}

// PromptDoc is what the agent reads about this system.
func (System) PromptDoc() string {
	return `Microsoft 365 — you are the user configured as your mailbox: its mail, its calendar, its OneDrive, and the
   SharePoint sites the app may reach. Times are ISO 8601 (2026-10-21T09:00:00) with a timezone name
   (Europe/Berlin); ids are Graph ids, copy them as they come.
   MAIL: list_mail_folders {} · list_mail {"folder":"inbox (default)|sentitems|drafts|archive|<name>","unread":true,"search":"words","limit":20}
   · get_mail {"id"} (body as text, recipients, attachment list) · get_mail_attachment {"id","attachment_id"} puts the file into
   your sandbox and returns its path · reply_mail {"id","body","reply_all":false,"attachments":["sandbox/path"]}
   · send_mail {"to":["a@b"],"cc":[],"subject","body","attachments":[]} · create_draft {same as send_mail} leaves it in Drafts for a
   person to check and send · move_mail {"id","folder"} · mark_read {"id","read":true}.
   Attachments to send: at most 3 MB together — a larger file goes to OneDrive with upload_file, and the mail carries its web_url.
   CALENDAR: list_events {"from","until","timezone"} · create_event {"subject","start","end","timezone","attendees":[],"body","location","online":true}
   (online = Teams meeting; attendees get an invitation) · respond_event {"id","response":"accept|decline|tentative","comment"}
   · cancel_event {"id","comment"} (as organiser) · free_busy {"emails":[],"from","until","timezone"}.
   FILES: drive_id "me" (or none) is your own OneDrive; for SharePoint, list_sites {"search":"words"} → list_drives {"site_id"} →
   use the drive's id. list_files {"drive_id","path"} · search_files {"drive_id","query"} · read_file {"drive_id","path"} (text up to 1 MB)
   · download_file {"drive_id","path","save":"sandbox/path (optional)"} · upload_file {"drive_id","from":"sandbox/path","path":"remote/path"}
   (replaces what is there) · write_file {"drive_id","path","content"} (text) · create_folder {"drive_id","path"}.
   OFFICE FILES (docx, xlsx, pptx): download_file → edit in the sandbox with python3 (python-docx, openpyxl, python-pptx) →
   upload_file onto the same path. Never write_file an Office file — it is a zip, not text.
   Sending mail and inviting people reach real people: send only what your assignment asks for; when in doubt, create_draft.
   Delete nothing — there is no delete action, on purpose.`
}
