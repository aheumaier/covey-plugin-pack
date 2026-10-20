package m365

import (
	"context"
	"encoding/base64"
	"encoding/json"
	"io"
	"net/http"
	"net/http/httptest"
	"net/url"
	"os"
	"path/filepath"
	"strings"
	"sync"
	"testing"

	"github.com/benjaminLedel/covey-plugin-sdk/target"
)

// fakeGraph is a Microsoft Graph double: the token endpoint, one mailbox, one
// OneDrive. It records what it was asked so the tests can check the requests,
// not only the answers.
type fakeGraph struct {
	t        *testing.T
	srv      *httptest.Server
	mu       sync.Mutex
	requests []string
	bodies   map[string]string
	tokens   int
	sent     []map[string]any
	read     map[string]bool
	files    map[string][]byte
}

func newFakeGraph(t *testing.T) *fakeGraph {
	f := &fakeGraph{t: t, bodies: map[string]string{}, read: map[string]bool{}, files: map[string][]byte{"Berichte/q3.txt": []byte("Quartal drei")}}
	f.srv = httptest.NewServer(http.HandlerFunc(f.serve))
	t.Cleanup(f.srv.Close)
	return f
}

func (f *fakeGraph) cred() target.Credential {
	return target.Credential{
		BaseURL: "agent@contoso.example graph=" + f.srv.URL + " login=" + f.srv.URL,
		Token:   "tenant-1:client-1:secret-1",
	}
}

func (f *fakeGraph) serve(w http.ResponseWriter, r *http.Request) {
	f.mu.Lock()
	defer f.mu.Unlock()
	body, _ := io.ReadAll(r.Body)
	key := r.Method + " " + r.URL.Path
	f.requests = append(f.requests, key+"?"+r.URL.RawQuery)
	f.bodies[key] = string(body)
	w.Header().Set("Content-Type", "application/json")

	if r.URL.Path == "/tenant-1/oauth2/v2.0/token" {
		f.tokens++
		form, _ := url.ParseQuery(string(body))
		if form.Get("client_secret") != "secret-1" || form.Get("grant_type") != "client_credentials" {
			w.WriteHeader(401)
			return
		}
		io.WriteString(w, `{"access_token":"tok-abc","expires_in":3600}`)
		return
	}
	if r.Header.Get("Authorization") != "Bearer tok-abc" {
		w.WriteHeader(401)
		io.WriteString(w, `{"error":{"code":"InvalidAuthenticationToken","message":"no"}}`)
		return
	}
	user := "/v1.0/users/agent@contoso.example"
	switch {
	case r.Method == "GET" && r.URL.Path == user+"/mailFolders":
		io.WriteString(w, `{"value":[{"id":"AAMk-inbox","displayName":"Inbox","unreadItemCount":2,"totalItemCount":9},{"id":"AAMk-proj","displayName":"Projekte","unreadItemCount":0,"totalItemCount":3}]}`)
	case r.Method == "GET" && r.URL.Path == user+"/mailFolders/inbox/messages":
		msgs := []map[string]any{
			{"id": "AAMkAD1", "subject": "Protokoll", "from": map[string]any{"emailAddress": map[string]string{"name": "Ada", "address": "ada@contoso.example"}}, "receivedDateTime": "2026-10-20T08:00:00Z", "lastModifiedDateTime": "2026-10-20T08:00:05Z", "isRead": f.read["AAMkAD1"], "hasAttachments": true},
			{"id": "AAMkAD2", "subject": "Termin", "receivedDateTime": "2026-10-20T07:00:00Z", "lastModifiedDateTime": "2026-10-20T07:00:01Z", "isRead": f.read["AAMkAD2"]},
		}
		if strings.Contains(r.URL.Query().Get("$filter"), "isRead eq false") {
			var unread []map[string]any
			for _, m := range msgs {
				if !m["isRead"].(bool) {
					unread = append(unread, m)
				}
			}
			msgs = unread
		}
		json.NewEncoder(w).Encode(map[string]any{"value": msgs})
	case r.Method == "GET" && r.URL.Path == user+"/messages/AAMkAD1":
		io.WriteString(w, `{"id":"AAMkAD1","subject":"Protokoll","body":{"contentType":"text","content":"Anbei das Protokoll."},"toRecipients":[{"emailAddress":{"address":"agent@contoso.example"}}],"hasAttachments":true}`)
	case r.Method == "GET" && r.URL.Path == user+"/messages/AAMkAD1/attachments":
		io.WriteString(w, `{"value":[{"@odata.type":"#microsoft.graph.fileAttachment","id":"att1","name":"protokoll.docx","contentType":"application/vnd.openxmlformats-officedocument.wordprocessingml.document","size":5}]}`)
	case r.Method == "GET" && r.URL.Path == user+"/messages/AAMkAD1/attachments/att1/$value":
		w.Header().Set("Content-Type", "application/octet-stream")
		io.WriteString(w, "PK\x03\x04x")
	case r.Method == "POST" && r.URL.Path == user+"/sendMail":
		var m map[string]any
		json.Unmarshal(body, &m)
		f.sent = append(f.sent, m)
		w.WriteHeader(202)
	case r.Method == "PATCH" && strings.HasPrefix(r.URL.Path, user+"/messages/"):
		f.read[strings.TrimPrefix(r.URL.Path, user+"/messages/")] = true
		io.WriteString(w, `{}`)
	case r.Method == "GET" && strings.HasPrefix(r.URL.Path, user+"/drive/root:/") && strings.HasSuffix(r.URL.Path, ":/content"):
		p := strings.TrimSuffix(strings.TrimPrefix(r.URL.Path, user+"/drive/root:/"), ":/content")
		data, ok := f.files[p]
		if !ok {
			w.WriteHeader(404)
			io.WriteString(w, `{"error":{"code":"itemNotFound","message":"gone"}}`)
			return
		}
		w.Header().Set("Content-Type", "text/plain")
		w.Write(data)
	case r.Method == "PUT" && strings.HasPrefix(r.URL.Path, user+"/drive/root:/"):
		p := strings.TrimSuffix(strings.TrimPrefix(r.URL.Path, user+"/drive/root:/"), ":/content")
		f.files[p] = body
		json.NewEncoder(w).Encode(map[string]any{"id": "item-" + p, "name": filepath.Base(p), "size": len(body), "webUrl": "https://contoso.example/" + p})
	case r.Method == "GET" && r.URL.Path == user+"/mailFolders/inbox":
		w.WriteHeader(404)
	default:
		w.WriteHeader(404)
		io.WriteString(w, `{"error":{"code":"notInFake","message":"`+key+`"}}`)
	}
}

func exec(t *testing.T, f *fakeGraph, ctx context.Context, action string, p any) (map[string]any, error) {
	t.Helper()
	raw, _ := json.Marshal(p)
	out, err := (System{}).Execute(ctx, action, raw, f.cred())
	if err != nil {
		return nil, err
	}
	data, _ := json.Marshal(out)
	var m map[string]any
	json.Unmarshal(data, &m)
	return m, nil
}

func resetTokens() {
	tokenMu.Lock()
	tokenCache = map[string]cachedToken{}
	tokenMu.Unlock()
}

func TestParseConfig(t *testing.T) {
	cfg, err := ParseConfig("agent@contoso.example", "t:c:s")
	if err != nil || cfg.Mailbox != "agent@contoso.example" || cfg.Tenant != "t" || cfg.ClientSecret != "s" {
		t.Fatalf("%+v %v", cfg, err)
	}
	for _, bad := range []string{"", "https://contoso.sharepoint.com/x", "agent@contoso.example extra", "agent@contoso.example graph=http://graph.example"} {
		if _, err := ParseConfig(bad, "t:c:s"); err == nil {
			t.Errorf("m365_url %q accepted", bad)
		}
	}
	if _, err := ParseConfig("agent@contoso.example", "t::s"); err == nil {
		t.Error("an empty client id was accepted")
	}
}

// TestMailRoundTrip: the inbox lists, a mail reads with its attachment, the
// attachment lands in the sandbox, and one token serves all of it.
func TestMailRoundTrip(t *testing.T) {
	resetTokens()
	f := newFakeGraph(t)
	dir := t.TempDir()
	ctx := target.WithWorkdir(context.Background(), dir)

	list, err := exec(t, f, ctx, "list_mail", map[string]any{"unread": true})
	if err != nil {
		t.Fatal(err)
	}
	if n := len(list["messages"].([]any)); n != 2 {
		t.Fatalf("unread: %d messages", n)
	}
	// The unread filter must lead with the orderby property, or Exchange
	// refuses it as inefficient.
	var listReq string
	for _, r := range f.requests {
		if strings.Contains(r, "/mailFolders/inbox/messages") {
			listReq = r
		}
	}
	if !strings.Contains(listReq, "receivedDateTime+ge") || !strings.Contains(listReq, "%24orderby=receivedDateTime+desc") {
		t.Errorf("list request: %s", listReq)
	}

	mail, err := exec(t, f, ctx, "get_mail", map[string]any{"id": "AAMkAD1"})
	if err != nil {
		t.Fatal(err)
	}
	if mail["body"] != "Anbei das Protokoll." || len(mail["attachments"].([]any)) != 1 {
		t.Fatalf("get_mail: %v", mail)
	}

	att, err := exec(t, f, ctx, "get_mail_attachment", map[string]any{"id": "AAMkAD1", "attachment_id": "att1"})
	if err != nil {
		t.Fatal(err)
	}
	p := att["path"].(string)
	if !strings.HasPrefix(p, dir) {
		t.Fatalf("attachment outside the sandbox: %s", p)
	}
	if data, _ := os.ReadFile(p); string(data) != "PK\x03\x04x" {
		t.Errorf("attachment content: %q", data)
	}
	if f.tokens != 1 {
		t.Errorf("%d token requests for four actions — the cache does not hold", f.tokens)
	}
}

func TestSendMailCarriesSandboxAttachment(t *testing.T) {
	resetTokens()
	f := newFakeGraph(t)
	dir := t.TempDir()
	os.WriteFile(filepath.Join(dir, "bericht.pdf"), []byte("%PDF-1.7"), 0o644)
	ctx := target.WithWorkdir(context.Background(), dir)

	if _, err := exec(t, f, ctx, "send_mail", map[string]any{"to": "ada@contoso.example; bob@contoso.example", "subject": "Bericht", "body": "Anbei.", "attachments": []string{"bericht.pdf"}}); err != nil {
		t.Fatal(err)
	}
	if len(f.sent) != 1 {
		t.Fatalf("sent %d", len(f.sent))
	}
	msg := f.sent[0]["message"].(map[string]any)
	if n := len(msg["toRecipients"].([]any)); n != 2 {
		t.Errorf("to: %d recipients", n)
	}
	att := msg["attachments"].([]any)[0].(map[string]any)
	got, _ := base64.StdEncoding.DecodeString(att["contentBytes"].(string))
	if att["name"] != "bericht.pdf" || string(got) != "%PDF-1.7" {
		t.Errorf("attachment: %v", att)
	}
	if f.sent[0]["saveToSentItems"] != true {
		t.Error("no copy in Sent Items")
	}

	// A path out of the sandbox is refused before anything is sent.
	if _, err := exec(t, f, ctx, "send_mail", map[string]any{"to": "ada@contoso.example", "subject": "x", "attachments": []string{"../../etc/passwd"}}); err == nil {
		t.Error("an attachment outside the sandbox was sent")
	}
}

// TestGateSignsUnreadMailItemByItem: nur-wenn: m365 reports one item per unread
// mail, and marking one read takes it out.
func TestGateSignsUnreadMailItemByItem(t *testing.T) {
	resetTokens()
	f := newFakeGraph(t)
	has, sig, err := (System{}).HasWorkSigned(context.Background(), f.cred(), "")
	if err != nil || !has {
		t.Fatalf("has=%v err=%v", has, err)
	}
	items, ok := target.WorkEntries(sig)
	if !ok || len(items) != 2 || items[0].Key != "mail:AAMkAD1" {
		t.Fatalf("signature %q → %v", sig, items)
	}
	if !target.SignsItems(System{}) {
		t.Error("the plugin does not say it signs item by item")
	}
	if !(System{}).WritesWorkSignature("m365:mark_read") || (System{}).WritesWorkSignature("m365:list_mail") {
		t.Error("WritesWorkSignature: mark_read must count, a read must not")
	}

	if _, err := exec(t, f, context.Background(), "mark_read", map[string]any{"id": "AAMkAD1"}); err != nil {
		t.Fatal(err)
	}
	_, sig2, _ := (System{}).HasWorkSigned(context.Background(), f.cred(), "")
	if items, _ := target.WorkEntries(sig2); len(items) != 1 || items[0].Key != "mail:AAMkAD2" {
		t.Errorf("after mark_read: %v", items)
	}
}

func TestFilesDownloadEditUpload(t *testing.T) {
	resetTokens()
	f := newFakeGraph(t)
	dir := t.TempDir()
	ctx := target.WithWorkdir(context.Background(), dir)

	got, err := exec(t, f, ctx, "download_file", map[string]any{"path": "Berichte/q3.txt"})
	if err != nil {
		t.Fatal(err)
	}
	local := got["path"].(string)
	if data, _ := os.ReadFile(local); string(data) != "Quartal drei" {
		t.Fatalf("downloaded: %q", data)
	}
	os.WriteFile(local, []byte("Quartal drei, ergänzt"), 0o644)
	rel, _ := filepath.Rel(dir, local)
	if _, err := exec(t, f, ctx, "upload_file", map[string]any{"from": rel, "path": "Berichte/q3.txt"}); err != nil {
		t.Fatal(err)
	}
	if string(f.files["Berichte/q3.txt"]) != "Quartal drei, ergänzt" {
		t.Errorf("uploaded: %q", f.files["Berichte/q3.txt"])
	}

	if _, err := exec(t, f, ctx, "download_file", map[string]any{"path": "../geheim"}); err == nil {
		t.Error("a path up out of the drive was accepted")
	}
	_, err = exec(t, f, ctx, "read_file", map[string]any{"path": "fehlt.txt"})
	if err == nil || !strings.Contains(err.Error(), "404") {
		t.Errorf("missing file: %v", err)
	}
}

func TestForbiddenNamesTheOperator(t *testing.T) {
	err := graphError("GET", "https://graph.example/v1.0/users/a@b/messages", 403, []byte(`{"error":{"code":"ErrorAccessDenied","message":"Access is denied."}}`))
	if !strings.Contains(err.Error(), "Exchange scope") || !strings.Contains(err.Error(), "ErrorAccessDenied") {
		t.Errorf("403 message: %v", err)
	}
}

func TestProbe(t *testing.T) {
	resetTokens()
	f := newFakeGraph(t)
	msg, err := (System{}).Probe(context.Background(), f.cred())
	if err != nil || !strings.Contains(msg, "2 unread of 9") {
		t.Errorf("probe: %q %v", msg, err)
	}
}

func TestUnknownActionAndRegistration(t *testing.T) {
	if _, ok := target.Get("m365"); !ok {
		t.Fatal("m365 is not registered")
	}
	if _, err := (System{}).Execute(context.Background(), "delete_mail", nil, target.Credential{}); err == nil {
		t.Error("an action that does not exist ran")
	}
	for _, a := range []string{"list_mail", "send_mail", "create_event", "upload_file", "search_files"} {
		if !strings.Contains((System{}).PromptDoc(), a) {
			t.Errorf("PromptDoc does not name %s", a)
		}
		if _, ok := actions[a]; !ok {
			t.Errorf("action %s missing", a)
		}
	}
}
