package m365

import (
	"context"
	"encoding/json"
	"fmt"
	"io"
	"net/http"
	"net/mail"
	"net/url"
	"strings"
	"sync"
	"time"

	"github.com/benjaminLedel/covey-plugin-sdk/target"
)

// Configuration out of the brokered secret pair. The broker knows two secrets
// per system, so m365_url carries the agent's mailbox plus optional endpoint
// overrides (separated by spaces):
//
//	m365_url   = agent@contoso.com [graph=https://graph.microsoft.com]
//	                               [login=https://login.microsoftonline.com]
//	m365_token = tenant-id:client-id:client-secret
//
// The mailbox is the one user the agent IS: its mail, its calendar, its
// OneDrive. The app registration behind m365_token holds application
// permissions, which Graph grants tenant-wide — what keeps the agent in its own
// mailbox is the Exchange scope the setup puts on the app (RBAC for
// Applications), not this plugin. The plugin only ever addresses the
// configured mailbox, so a scope the operator forgot is a wider door, not a
// different agent.
//
// A token WITHOUT the two colons is used as a ready-made bearer token
// (tests, demos).

// Config is the parsed connection configuration.
type Config struct {
	Mailbox   string // the agent's own user (UPN / primary address)
	GraphBase string // without /v1.0
	LoginBase string

	Tenant       string
	ClientID     string
	ClientSecret string

	StaticToken string
}

// ParseConfig breaks the brokered credential down.
func ParseConfig(baseURL, token string) (Config, error) {
	cfg := Config{
		GraphBase: "https://graph.microsoft.com",
		LoginBase: "https://login.microsoftonline.com",
	}
	for _, part := range strings.Fields(baseURL) {
		switch {
		case strings.HasPrefix(part, "graph="):
			cfg.GraphBase = strings.TrimRight(strings.TrimPrefix(part, "graph="), "/")
		case strings.HasPrefix(part, "login="):
			cfg.LoginBase = strings.TrimRight(strings.TrimPrefix(part, "login="), "/")
		case cfg.Mailbox == "":
			cfg.Mailbox = part
		default:
			return Config{}, fmt.Errorf("m365_url: unexpected component %q (expected: mailbox [graph=…] [login=…])", part)
		}
	}
	if cfg.Mailbox == "" {
		return Config{}, fmt.Errorf("m365_url: mailbox missing — store the agent's own address, e.g. agent@contoso.com")
	}
	if a, err := mail.ParseAddress(cfg.Mailbox); err != nil || a.Address != cfg.Mailbox {
		return Config{}, fmt.Errorf("m365_url: %q is not a mail address — the plugin works in exactly one mailbox, the agent's own", cfg.Mailbox)
	}
	for _, base := range []string{cfg.GraphBase, cfg.LoginBase} {
		if u, err := url.Parse(base); err != nil || u.Host == "" || (u.Scheme != "https" && !isLoopback(u.Hostname())) {
			return Config{}, fmt.Errorf("m365_url: endpoint %q must be https", base)
		}
	}

	token = strings.TrimSpace(token)
	if token == "" {
		return Config{}, fmt.Errorf("m365_token missing")
	}
	if parts := strings.SplitN(token, ":", 3); len(parts) == 3 {
		if parts[0] == "" || parts[1] == "" || parts[2] == "" {
			return Config{}, fmt.Errorf("m365_token must be %q", "tenant-id:client-id:client-secret")
		}
		cfg.Tenant, cfg.ClientID, cfg.ClientSecret = parts[0], parts[1], parts[2]
	} else {
		cfg.StaticToken = token
	}
	return cfg, nil
}

func isLoopback(host string) bool {
	return host == "127.0.0.1" || host == "localhost" || host == "::1"
}

// Token cache for the client-credentials flow: Entra ID tokens live about an
// hour, and one per action would be a roundtrip for nothing and an early
// throttle. It lives in the daemon's memory only, like the credential itself.
var (
	tokenMu    sync.Mutex
	tokenCache = map[string]cachedToken{}
)

type cachedToken struct {
	token   string
	expires time.Time
}

// AccessToken returns the static token, the cached one, or a fresh one.
func (cfg Config) AccessToken(ctx context.Context) (string, error) {
	if cfg.StaticToken != "" {
		return cfg.StaticToken, nil
	}
	key := cfg.LoginBase + "|" + cfg.Tenant + "|" + cfg.ClientID
	tokenMu.Lock()
	cached, ok := tokenCache[key]
	tokenMu.Unlock()
	if ok && time.Now().Before(cached.expires) {
		return cached.token, nil
	}

	form := url.Values{
		"grant_type":    {"client_credentials"},
		"client_id":     {cfg.ClientID},
		"client_secret": {cfg.ClientSecret},
		"scope":         {cfg.GraphBase + "/.default"},
	}
	endpoint := cfg.LoginBase + "/" + url.PathEscape(cfg.Tenant) + "/oauth2/v2.0/token"
	req, err := http.NewRequestWithContext(ctx, http.MethodPost, endpoint, strings.NewReader(form.Encode()))
	if err != nil {
		return "", err
	}
	req.Header.Set("Content-Type", "application/x-www-form-urlencoded")
	resp, err := target.Client("m365", 15*time.Second).Do(req)
	if err != nil {
		return "", fmt.Errorf("entra id token: %w", err)
	}
	defer resp.Body.Close()
	data, err := io.ReadAll(io.LimitReader(resp.Body, 1<<20))
	if err != nil {
		return "", err
	}
	if resp.StatusCode < 200 || resp.StatusCode >= 300 {
		return "", fmt.Errorf("entra id token: HTTP %d: %.300s", resp.StatusCode, data)
	}
	var out struct {
		AccessToken string `json:"access_token"`
		ExpiresIn   int    `json:"expires_in"`
	}
	if err := json.Unmarshal(data, &out); err != nil || out.AccessToken == "" {
		return "", fmt.Errorf("entra id token: unexpected response")
	}
	ttl := time.Duration(out.ExpiresIn) * time.Second
	if ttl > 2*time.Minute {
		ttl -= 2 * time.Minute
	}
	tokenMu.Lock()
	tokenCache[key] = cachedToken{token: out.AccessToken, expires: time.Now().Add(ttl)}
	tokenMu.Unlock()
	return out.AccessToken, nil
}
