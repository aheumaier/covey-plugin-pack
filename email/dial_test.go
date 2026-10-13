package email

import (
	"bufio"
	"context"
	"encoding/base64"
	"io"
	"net"
	"net/http"
	"strings"
	"sync"
	"testing"
)

// connectProxy is the egress proxy in miniature: CONNECT host:port, a
// credential to check, an allowlist of exactly one target.
type connectProxy struct {
	ln      net.Listener
	allow   string
	mu      sync.Mutex
	auth    []string
	tunnels []string
}

func newConnectProxy(t *testing.T, allow string) *connectProxy {
	t.Helper()
	ln, err := net.Listen("tcp", "127.0.0.1:0")
	if err != nil {
		t.Fatal(err)
	}
	p := &connectProxy{ln: ln, allow: allow}
	t.Cleanup(func() { ln.Close() })
	go func() {
		for {
			c, err := ln.Accept()
			if err != nil {
				return
			}
			go p.handle(c)
		}
	}()
	return p
}

func (p *connectProxy) handle(c net.Conn) {
	defer c.Close()
	req, err := http.ReadRequest(bufio.NewReader(c))
	if err != nil || req.Method != http.MethodConnect {
		return
	}
	p.mu.Lock()
	p.auth = append(p.auth, req.Header.Get("Proxy-Authorization"))
	p.tunnels = append(p.tunnels, req.Host)
	p.mu.Unlock()
	if req.Host != p.allow {
		io.WriteString(c, "HTTP/1.1 403 Forbidden\r\n\r\n")
		return
	}
	up, err := net.Dial("tcp", req.Host)
	if err != nil {
		io.WriteString(c, "HTTP/1.1 502 Bad Gateway\r\n\r\n")
		return
	}
	defer up.Close()
	io.WriteString(c, "HTTP/1.1 200 Connection Established\r\n\r\n")
	go io.Copy(up, c)
	io.Copy(c, up)
}

func clearProxyEnv(t *testing.T) {
	for _, n := range []string{"HTTPS_PROXY", "https_proxy", "ALL_PROXY", "all_proxy", "NO_PROXY", "no_proxy"} {
		t.Setenv(n, "")
	}
}

func TestIMAPGoesThroughTheEgressProxy(t *testing.T) {
	clearProxyEnv(t)
	imapAddr := newMemIMAP(t, testUser, "pw")
	smtp := newFakeSMTP(t)
	cred := testCred(t, imapAddr, smtp.ln.Addr().String())
	appendMail(t, imapAddr, testUser, "pw", "From: kunde@example.com\r\nSubject: a\r\n\r\nx\r\n", false)

	proxy := newConnectProxy(t, imapAddr)
	t.Setenv("HTTPS_PROXY", "http://agent-1:s3cret@"+proxy.ln.Addr().String())

	// IMAP is on the allowlist: the greeting arrives through the tunnel.
	unread := exec(t, cred, "list_unread", `{}`).([]MessageSummary)
	if len(unread) != 1 {
		t.Fatalf("list_unread through the proxy: %+v", unread)
	}
	proxy.mu.Lock()
	auth, tunnels := proxy.auth[0], proxy.tunnels[0]
	proxy.mu.Unlock()
	if tunnels != imapAddr || auth != "Basic "+base64.StdEncoding.EncodeToString([]byte("agent-1:s3cret")) {
		t.Fatalf("CONNECT %q with %q", tunnels, auth)
	}

	// SMTP is not: the proxy's refusal is the error, naming the allowlist.
	_, err := (System{}).Execute(context.Background(), "send",
		[]byte(`{"to":["kunde@example.com"],"subject":"x","body":"y"}`), cred)
	if err == nil || !strings.Contains(err.Error(), "403") || !strings.Contains(err.Error(), "allowlist") {
		t.Fatalf("a refused tunnel has to say so, got %v", err)
	}
	if from, _, _ := smtp.snapshot(); from != "" {
		t.Fatal("nothing may reach the SMTP server around the proxy")
	}

	// NO_PROXY exempts the host: straight to the server again.
	t.Setenv("NO_PROXY", "localhost,127.0.0.1")
	exec(t, cred, "send", `{"to":["kunde@example.com"],"subject":"x","body":"y"}`)
	if from, _, _ := smtp.snapshot(); from != testUser {
		t.Fatalf("NO_PROXY host must be dialled directly, envelope from %q", from)
	}
}

func TestNoProxyMatching(t *testing.T) {
	for _, c := range []struct {
		host, list string
		want       bool
	}{
		{"mail.example.org", "", false},
		{"mail.example.org", "*", true},
		{"mail.example.org", "example.org", true},
		{"mail.example.org", ".example.org", true},
		{"mail.example.org", "*.example.org", true},
		{"example.org", ".example.org", true},
		{"badexample.org", "example.org", false},
		{"127.0.0.1", "localhost, 127.0.0.1:25", true},
	} {
		if got := noProxy(c.host, c.list); got != c.want {
			t.Errorf("noProxy(%q, %q) = %v", c.host, c.list, got)
		}
	}
}
