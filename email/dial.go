package email

import (
	"bufio"
	"encoding/base64"
	"fmt"
	"net"
	"net/http"
	"net/url"
	"os"
	"strings"
	"time"
)

// IMAP and SMTP are not HTTP, so nothing finds the egress proxy for them on its
// own. In an isolated sandbox that proxy is the only way out: it tunnels any
// host:port on the agent's allowlist with HTTP CONNECT. dialTCP takes that
// tunnel whenever HTTPS_PROXY names a proxy and NO_PROXY does not exempt the
// host, and dials directly otherwise — on a host without a proxy nothing
// changes.

const dialTimeout = 30 * time.Second

func dialTCP(addr string) (net.Conn, error) {
	proxy, err := proxyFor(addr)
	if err != nil {
		return nil, err
	}
	if proxy == nil {
		return net.DialTimeout("tcp", addr, dialTimeout)
	}
	return dialConnect(proxy, addr)
}

// proxyFor reads HTTPS_PROXY (then ALL_PROXY) and NO_PROXY, upper or lower case.
func proxyFor(addr string) (*url.URL, error) {
	raw := firstEnv("HTTPS_PROXY", "https_proxy", "ALL_PROXY", "all_proxy")
	if raw == "" {
		return nil, nil
	}
	host, _, err := net.SplitHostPort(addr)
	if err != nil {
		host = addr
	}
	if noProxy(host, firstEnv("NO_PROXY", "no_proxy")) {
		return nil, nil
	}
	if !strings.Contains(raw, "://") {
		raw = "http://" + raw
	}
	u, err := url.Parse(raw)
	if err != nil || u.Host == "" {
		return nil, fmt.Errorf("HTTPS_PROXY %q is not a proxy URL", raw)
	}
	if u.Scheme != "http" {
		return nil, fmt.Errorf("HTTPS_PROXY %q: only an http:// proxy can tunnel IMAP/SMTP", raw)
	}
	return u, nil
}

func firstEnv(names ...string) string {
	for _, n := range names {
		if v := strings.TrimSpace(os.Getenv(n)); v != "" {
			return v
		}
	}
	return ""
}

// noProxy matches host against a NO_PROXY list: "*", an exact host, or a
// domain that also covers its subdomains (with or without a leading dot).
func noProxy(host, list string) bool {
	host = strings.ToLower(strings.TrimSuffix(host, "."))
	for _, e := range strings.Split(list, ",") {
		e = strings.ToLower(strings.TrimSpace(e))
		if h, _, err := net.SplitHostPort(e); err == nil {
			e = h
		}
		if e == "*" {
			return true
		}
		e = strings.TrimPrefix(e, "*")
		switch {
		case e == "":
		case host == strings.TrimPrefix(e, "."):
			return true
		case strings.HasSuffix(host, "."+strings.TrimPrefix(e, ".")):
			return true
		}
	}
	return false
}

// dialConnect opens a CONNECT tunnel to addr through the proxy.
func dialConnect(proxy *url.URL, addr string) (net.Conn, error) {
	proxyAddr := proxy.Host
	if proxy.Port() == "" {
		proxyAddr = net.JoinHostPort(proxy.Hostname(), "80")
	}
	conn, err := net.DialTimeout("tcp", proxyAddr, dialTimeout)
	if err != nil {
		return nil, fmt.Errorf("egress proxy %s: %w", proxyAddr, err)
	}
	req := "CONNECT " + addr + " HTTP/1.1\r\nHost: " + addr + "\r\n"
	if proxy.User != nil {
		pw, _ := proxy.User.Password()
		cred := base64.StdEncoding.EncodeToString([]byte(proxy.User.Username() + ":" + pw))
		req += "Proxy-Authorization: Basic " + cred + "\r\n"
	}
	req += "\r\n"
	_ = conn.SetDeadline(time.Now().Add(dialTimeout))
	if _, err := conn.Write([]byte(req)); err != nil {
		conn.Close()
		return nil, fmt.Errorf("egress proxy %s: %w", proxyAddr, err)
	}
	br := bufio.NewReader(conn)
	resp, err := http.ReadResponse(br, &http.Request{Method: http.MethodConnect})
	if err != nil {
		conn.Close()
		return nil, fmt.Errorf("egress proxy %s: %w", proxyAddr, err)
	}
	resp.Body.Close()
	if resp.StatusCode != http.StatusOK {
		conn.Close()
		return nil, fmt.Errorf("egress proxy refused %s: %s (is the host on the agent's egress allowlist?)", addr, resp.Status)
	}
	_ = conn.SetDeadline(time.Time{})
	// The server may speak first (an IMAP greeting): whatever the reader
	// already holds belongs to the tunnel, not to the proxy's answer.
	return &bufferedConn{Conn: conn, r: br}, nil
}

type bufferedConn struct {
	net.Conn
	r *bufio.Reader
}

func (c *bufferedConn) Read(p []byte) (int, error) { return c.r.Read(p) }
