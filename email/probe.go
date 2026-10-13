package email

import (
	"context"
	"fmt"
	"strings"

	"github.com/emersion/go-imap/v2"
	"github.com/emersion/go-imap/v2/imapclient"

	"github.com/benjaminLedel/covey-plugin-sdk/target"
)

// Probe answers what an installation wants to know right after storing the
// mail credential, before the agent's first run finds out the hard way: do
// both endpoints take the password, and as which sender will the agent write?
//
// IMAP: log in and ask INBOX for its unread count. SMTP: connect, secure and
// log in, then QUIT — no MAIL FROM, nothing is delivered. The answer names the
// sender address and the fixed copy from cc= as well, because those are the
// two things a typo in email_url gets wrong without any error at all.
func (System) Probe(ctx context.Context, cred target.Credential) (string, error) {
	cfg, err := ParseConfig(cred)
	if err != nil {
		return "", err
	}
	var unseen uint32
	err = withIMAP(cfg, func(c *imapclient.Client) error {
		st, err := c.Status("INBOX", &imap.StatusOptions{NumUnseen: true}).Wait()
		if err != nil {
			return fmt.Errorf("imap INBOX: %w", err)
		}
		if st.NumUnseen != nil {
			unseen = *st.NumUnseen
		}
		return nil
	})
	if err != nil {
		return "", err
	}
	c, err := dialSMTP(cfg)
	if err != nil {
		return "", err
	}
	_ = c.Quit()
	c.Close()

	id := fmt.Sprintf("%s (INBOX: %d unread", cfg.From, unseen)
	if len(cfg.Cc) > 0 {
		id += "; copy on every mail: " + strings.Join(cfg.Cc, ", ")
	}
	return id + ")", nil
}

var _ target.Prober = System{}
