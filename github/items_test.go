package github

import (
	"testing"

	"github.com/benjaminLedel/covey-plugin-sdk/target"
)

// The signature reads back item by item, as SignsWorkItems promises.
func TestWorkSignatureReadsBackAsItems(t *testing.T) {
	sig := workSig([]string{
		threadSig("pr", "acme/app", 3, []Comment{{ID: 77}}, "0123456789abcdef"),
		threadSig("issue", "acme/app", 9, nil, ""),
	})
	items, ok := target.WorkEntries(sig)
	if !target.SignsItems(System{}) || !ok || len(items) != 2 || items[1].Key != "pr:acme/app#3" || items[1].Version != "77+0123456789ab" {
		t.Fatalf("%q read as %v, %v", sig, items, ok)
	}
}
