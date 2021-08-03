package msgraph

import (
	b64 "encoding/base64"
	"fmt"
)

type Mail struct {
	Message Message `json:"message"`
}

type Message struct {
	Subject       string       `json:"subject"`
	Body          MsgBody      `json:"body"`
	ToRecipients  []Recipient  `json:"toRecipients"`
	CcRecipients  []Recipient  `json:"ccRecipients"`
	BccRecipients []Recipient  `json:"bccRecipients"`
	From          Recipient    `json:"from"`
	Attachments   []Attachment `json:"attachments"`
	// SaveToSentItems bool         `json:"saveToSentItems"`
}

type MsgBody struct {
	ContentType string `json:"contentType"`
	Content     string `json:"content"`
}

type Recipient struct {
	EmailAddress EmailAddress `json:"emailAddress"`
}

type Attachment struct {
	DataType     string `json:"@odata.type"`
	Name         string `json:"name"`
	ContentType  string `json:"contentType"`
	ContentBytes string `json:"contentBytes"` //base64 encoded string
}

func NewMail() *Mail {
	return &Mail{
		Message{
			Body: MsgBody{
				ContentType: "Text",
			},
			CcRecipients:  []Recipient{},
			BccRecipients: []Recipient{},
			From:          Recipient{},
			Attachments:   []Attachment{},
		},
	}
}

func MakeMail() Mail {
	return Mail{
		Message{
			Body: MsgBody{
				ContentType: "Text",
			},
			ToRecipients:  []Recipient{},
			CcRecipients:  []Recipient{},
			BccRecipients: []Recipient{},
			From:          Recipient{},
			Attachments:   []Attachment{},
		},
	}
}

func (m *Mail) Subject(subject string) {
	m.Message.Subject = subject
}

// Add recipient takes 1 or 2 inputs
// Hint: First argument: email (required), Second argument: Name (optional)
func (m *Mail) AddRecipient(recipient ...string) error {
	switch len(recipient) {
	case 1:
		m.Message.ToRecipients = append(m.Message.ToRecipients, Recipient{EmailAddress: EmailAddress{Address: recipient[0]}})
	case 2:
		m.Message.ToRecipients = append(m.Message.ToRecipients, Recipient{EmailAddress: EmailAddress{Address: recipient[0], Name: recipient[1]}})
	default:
		return fmt.Errorf("add recipient only accepts 1-2 arguments")
	}

	return nil
}

// Takes 1 or 2 arguments in order Address, Name
func (m *Mail) CCRecipient(recipient ...string) error {
	switch len(recipient) {
	case 1:
		m.Message.CcRecipients = append(m.Message.CcRecipients, Recipient{EmailAddress: EmailAddress{Address: recipient[0]}})
	case 2:
		m.Message.CcRecipients = append(m.Message.CcRecipients, Recipient{EmailAddress: EmailAddress{Address: recipient[0], Name: recipient[1]}})
	default:
		return fmt.Errorf("CCRecipient() only accepts 1-2 arguments")
	}

	return nil
}

// Takes 1 or 2 arguments in order Address, Name
func (m *Mail) BCCRecipient(recipient ...string) error {
	switch len(recipient) {
	case 1:
		m.Message.BccRecipients = append(m.Message.BccRecipients, Recipient{EmailAddress: EmailAddress{Address: recipient[0]}})
	case 2:
		m.Message.BccRecipients = append(m.Message.BccRecipients, Recipient{EmailAddress: EmailAddress{Address: recipient[0], Name: recipient[1]}})
	default:
		return fmt.Errorf("BCCRecipient() only accepts 1-2 arguments")
	}

	return nil
}

// Takes 1 or 2 arguments in order Address, Name
func (m *Mail) From(from ...string) error {
	switch len(from) {
	case 1:
		m.Message.From = Recipient{EmailAddress{Address: from[0]}}
	case 2:
		m.Message.From = Recipient{EmailAddress{Address: from[0], Name: from[1]}}
	default:
		return fmt.Errorf("From() only accepts 1-2 arguments")
	}

	return nil
}

func (m *Mail) Body(contentType, content string) {
	m.Message.Body.ContentType = contentType
	m.Message.Body.Content = content
}

func (m *Mail) AddFileAttachment(attachmentName, contentType, content string) {
	attachment := Attachment{
		DataType:     "#microsoft.graph.fileAttachment",
		Name:         attachmentName,
		ContentType:  contentType,
		ContentBytes: b64.StdEncoding.EncodeToString([]byte(content)),
	}

	m.Message.Attachments = append(m.Message.Attachments, attachment)
}
