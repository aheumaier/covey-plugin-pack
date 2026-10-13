package jira

import (
	"testing"

	"github.com/benjaminLedel/covey-plugin-sdk/target"
)

func TestTheGateSaysItSignsItems(t *testing.T) {
	if !target.SignsItems(System{}) {
		t.Fatal("the gate does not say that it signs items")
	}
}
