package msgraph

import (
	"os"
	"testing"
)

var (
	// Microsoft Graph tenant ID
	TenantID = os.Getenv("MSGraphTenantID")
	// Microsoft Graph Application ID
	ApplicationID = os.Getenv("MSGraphApplicationID")
	// Microsoft Graph Client Secret
	ClientSecret = os.Getenv("MSGraphClientSecret")
)

func TestEnv(t *testing.T) {
	// t.Fatal(TenantID, ApplicationID, ClientSecret)
}

func TestSendEmail(t *testing.T) {
	mail := MakeMail()

	mail.From("")
	mail.AddRecipient("")
	mail.Body("HTML", "<strong>Hello world</strong>")
	mail.Subject("Hello World")
	mail.AddFileAttachment("test.csv", "text/plain", "test,1\n2,3\n")

	err := graphClient.SendEmail(mail)
	if err != nil {
		t.Fatalf(err.Error())
	}

}
