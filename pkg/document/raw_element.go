// Package document provides raw element preservation for complex XML content
package document

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"strings"
)

// RawElement preserves unknown XML elements that we can't fully parse.
// This allows complex content like mc:AlternateContent, shapes, SmartArt,
// and other advanced Word features to survive the parse-modify-serialize cycle.
type RawElement struct {
	// The full raw XML of the element, including the start and end tags
	RawXML string
	// The element name for debugging/logging
	ElementName string
}

// ElementType returns the element type identifier
func (r *RawElement) ElementType() string {
	return "raw"
}

// MarshalXML writes the raw XML directly to the encoder
func (r *RawElement) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	// Flush any pending tokens
	if err := e.Flush(); err != nil {
		return err
	}

	// We need to write raw XML, but xml.Encoder doesn't support that directly.
	// We'll use a special marker that we'll replace after encoding.
	// For now, we'll use CharData which will be escaped, then we'll need
	// to handle this differently.

	// Actually, we need to write to the underlying writer directly.
	// This is a limitation of Go's xml package.
	// We'll use xml.CharData and mark it for post-processing.

	// Better approach: use a custom token
	return e.EncodeToken(xml.CharData([]byte(r.RawXML)))
}

// captureRawElement reads an entire XML element and its children as raw XML string.
// This is used to preserve complex elements that we don't fully parse.
func captureRawElement(decoder *xml.Decoder, startElement xml.StartElement) (*RawElement, error) {
	var buf bytes.Buffer

	// Write the start element
	if err := writeStartElement(&buf, startElement); err != nil {
		return nil, fmt.Errorf("failed to write start element: %w", err)
	}

	// Read and write all content until we hit the matching end element
	depth := 1
	for depth > 0 {
		token, err := decoder.Token()
		if err != nil {
			if err == io.EOF {
				return nil, fmt.Errorf("unexpected EOF while capturing element %s", startElement.Name.Local)
			}
			return nil, fmt.Errorf("error reading token: %w", err)
		}

		switch t := token.(type) {
		case xml.StartElement:
			depth++
			if err := writeStartElement(&buf, t); err != nil {
				return nil, err
			}
		case xml.EndElement:
			depth--
			if err := writeEndElement(&buf, t); err != nil {
				return nil, err
			}
		case xml.CharData:
			// Escape and write character data
			xml.EscapeText(&buf, t)
		case xml.Comment:
			buf.WriteString("<!--")
			buf.Write(t)
			buf.WriteString("-->")
		case xml.ProcInst:
			buf.WriteString("<?")
			buf.WriteString(t.Target)
			if len(t.Inst) > 0 {
				buf.WriteByte(' ')
				buf.Write(t.Inst)
			}
			buf.WriteString("?>")
		case xml.Directive:
			buf.WriteString("<!")
			buf.Write(t)
			buf.WriteString(">")
		}
	}

	return &RawElement{
		RawXML:      buf.String(),
		ElementName: formatElementName(startElement.Name),
	}, nil
}

// writeStartElement writes an XML start element to the buffer
func writeStartElement(buf *bytes.Buffer, elem xml.StartElement) error {
	buf.WriteByte('<')

	// Write element name with namespace prefix if present
	if elem.Name.Space != "" {
		// Find the prefix for this namespace from attributes
		prefix := findNamespacePrefix(elem)
		if prefix != "" {
			buf.WriteString(prefix)
			buf.WriteByte(':')
		}
	}
	buf.WriteString(elem.Name.Local)

	// Write attributes
	for _, attr := range elem.Attr {
		buf.WriteByte(' ')
		if attr.Name.Space != "" {
			// Handle namespace prefixes in attributes
			if attr.Name.Space == "http://www.w3.org/2000/xmlns/" || attr.Name.Local == "xmlns" {
				if attr.Name.Local == "xmlns" {
					buf.WriteString("xmlns")
				} else {
					buf.WriteString("xmlns:")
					buf.WriteString(attr.Name.Local)
				}
			} else {
				// Try to find prefix
				prefix := getAttributePrefix(attr.Name)
				if prefix != "" {
					buf.WriteString(prefix)
					buf.WriteByte(':')
				}
				buf.WriteString(attr.Name.Local)
			}
		} else {
			buf.WriteString(attr.Name.Local)
		}
		buf.WriteString(`="`)
		xml.EscapeText(buf, []byte(attr.Value))
		buf.WriteByte('"')
	}

	buf.WriteByte('>')
	return nil
}

// writeEndElement writes an XML end element to the buffer
func writeEndElement(buf *bytes.Buffer, elem xml.EndElement) error {
	buf.WriteString("</")
	if elem.Name.Space != "" {
		prefix := getNamespacePrefixFromSpace(elem.Name.Space)
		if prefix != "" {
			buf.WriteString(prefix)
			buf.WriteByte(':')
		}
	}
	buf.WriteString(elem.Name.Local)
	buf.WriteByte('>')
	return nil
}

// findNamespacePrefix finds the namespace prefix from element attributes
func findNamespacePrefix(elem xml.StartElement) string {
	for _, attr := range elem.Attr {
		if attr.Name.Space == "http://www.w3.org/2000/xmlns/" && attr.Value == elem.Name.Space {
			return attr.Name.Local
		}
		if attr.Name.Local == "xmlns" && attr.Value == elem.Name.Space {
			return ""
		}
	}
	return getNamespacePrefixFromSpace(elem.Name.Space)
}

// getAttributePrefix gets the prefix for an attribute namespace
func getAttributePrefix(name xml.Name) string {
	return getNamespacePrefixFromSpace(name.Space)
}

// getNamespacePrefixFromSpace returns common namespace prefixes
func getNamespacePrefixFromSpace(space string) string {
	prefixes := map[string]string{
		"http://schemas.openxmlformats.org/wordprocessingml/2006/main":                      "w",
		"http://schemas.openxmlformats.org/markup-compatibility/2006":                       "mc",
		"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing":            "wp",
		"http://schemas.openxmlformats.org/drawingml/2006/main":                             "a",
		"http://schemas.openxmlformats.org/drawingml/2006/picture":                          "pic",
		"http://schemas.openxmlformats.org/officeDocument/2006/relationships":               "r",
		"http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing":               "wp14",
		"http://schemas.microsoft.com/office/word/2010/wordml":                              "w14",
		"http://schemas.microsoft.com/office/word/2010/wordprocessingShape":                 "wps",
		"http://schemas.microsoft.com/office/word/2010/wordprocessingGroup":                 "wpg",
		"http://schemas.microsoft.com/office/drawing/2014/chartex":                          "cx",
		"urn:schemas-microsoft-com:vml":                                                     "v",
		"urn:schemas-microsoft-com:office:office":                                           "o",
		"urn:schemas-microsoft-com:office:word":                                             "w10",
		"http://schemas.openxmlformats.org/officeDocument/2006/math":                        "m",
	}
	if prefix, ok := prefixes[space]; ok {
		return prefix
	}
	return ""
}

// formatElementName formats an xml.Name for logging
func formatElementName(name xml.Name) string {
	if name.Space != "" {
		prefix := getNamespacePrefixFromSpace(name.Space)
		if prefix != "" {
			return prefix + ":" + name.Local
		}
		return "{" + name.Space + "}" + name.Local
	}
	return name.Local
}

// unescapeRawXMLMarkers finds raw XML markers in the serialized output and unescapes their content.
// The markers are: __RAW_XML_START__...content...__RAW_XML_END__
// The content inside gets XML-escaped during marshaling, so we need to selectively unescape it.
func unescapeRawXMLMarkers(data []byte) []byte {
	const startMarker = "__RAW_XML_START__"
	const endMarker = "__RAW_XML_END__"

	result := make([]byte, 0, len(data))
	remaining := data

	for {
		// Find the start marker
		startIdx := bytes.Index(remaining, []byte(startMarker))
		if startIdx == -1 {
			// No more markers, append the rest
			result = append(result, remaining...)
			break
		}

		// Append everything before the marker
		result = append(result, remaining[:startIdx]...)

		// Find the end marker
		afterStart := remaining[startIdx+len(startMarker):]
		endIdx := bytes.Index(afterStart, []byte(endMarker))
		if endIdx == -1 {
			// Malformed - no end marker, append the rest as-is
			result = append(result, remaining[startIdx:]...)
			break
		}

		// Extract the escaped content between markers
		escapedContent := afterStart[:endIdx]

		// Unescape the XML entities in this content
		unescaped := unescapeXMLEntities(escapedContent)

		// Append the unescaped content (without the markers)
		result = append(result, unescaped...)

		// Move past the end marker
		remaining = afterStart[endIdx+len(endMarker):]
	}

	return result
}

// unescapeXMLEntities unescapes XML entities in the given byte slice
func unescapeXMLEntities(data []byte) []byte {
	s := string(data)
	// Handle named entities
	s = strings.ReplaceAll(s, "&lt;", "<")
	s = strings.ReplaceAll(s, "&gt;", ">")
	s = strings.ReplaceAll(s, "&amp;", "&")
	s = strings.ReplaceAll(s, "&quot;", "\"")
	s = strings.ReplaceAll(s, "&apos;", "'")
	// Handle numeric entities for common characters
	s = strings.ReplaceAll(s, "&#34;", "\"")  // Double quote
	s = strings.ReplaceAll(s, "&#39;", "'")   // Single quote
	s = strings.ReplaceAll(s, "&#60;", "<")   // Less than
	s = strings.ReplaceAll(s, "&#62;", ">")   // Greater than
	s = strings.ReplaceAll(s, "&#38;", "&")   // Ampersand
	// Handle hex numeric entities
	s = strings.ReplaceAll(s, "&#x22;", "\"") // Double quote (hex)
	s = strings.ReplaceAll(s, "&#x27;", "'")  // Single quote (hex)
	s = strings.ReplaceAll(s, "&#x3C;", "<")  // Less than (hex)
	s = strings.ReplaceAll(s, "&#x3E;", ">")  // Greater than (hex)
	s = strings.ReplaceAll(s, "&#x26;", "&")  // Ampersand (hex)
	return []byte(s)
}
