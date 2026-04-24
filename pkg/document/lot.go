// Package document provides Word document List of Tables/Figures generation
package document

import (
	"encoding/xml"
	"fmt"
	"strings"
)

// LOTConfig configuration for List of Tables/Figures
type LOTConfig struct {
	Title          string // Title for the list (e.g., "List of Tables", "Lista de Tablas")
	SeqIdentifier  string // SEQ field identifier (e.g., "Table", "Figure")
	ShowPageNum    bool   // Whether to show page numbers
	RightAlign     bool   // Whether to right-align page numbers
	UseHyperlink   bool   // Whether to use hyperlinks
	DotLeader      bool   // Whether to use dot leaders
	InsertPosition int    // Insert position (-1 for auto)
	FontFamily     string // Font family (empty for default)
	FontSize       int    // Font size in points (0 for default)
	TitleFontSize  int    // Title font size in points (0 for default)
}

// LOTEntry represents an entry in the List of Tables/Figures
type LOTEntry struct {
	Caption    string // Full caption text (e.g., "Table 1: Sales Data")
	Number     int    // Sequence number
	PageNum    int    // Page number
	BookmarkID string // Bookmark ID for hyperlink
}

// SEQField represents a SEQ (sequence) field in Word
type SEQField struct {
	XMLName    xml.Name `xml:"w:fldSimple"`
	Instr      string   `xml:"w:instr,attr"`
	Identifier string   // The sequence identifier (e.g., "Table", "Figure")
	Number     int      // The current sequence number
}

// DefaultLOTConfig returns default configuration for List of Tables
func DefaultLOTConfig() *LOTConfig {
	return &LOTConfig{
		Title:          "List of Tables",
		SeqIdentifier:  "Table",
		ShowPageNum:    true,
		RightAlign:     true,
		UseHyperlink:   true,
		DotLeader:      true,
		InsertPosition: -1,
		FontFamily:     "",
		FontSize:       0,
		TitleFontSize:  0,
	}
}

// DefaultLOFConfig returns default configuration for List of Figures
func DefaultLOFConfig() *LOTConfig {
	return &LOTConfig{
		Title:          "List of Figures",
		SeqIdentifier:  "Figure",
		ShowPageNum:    true,
		RightAlign:     true,
		UseHyperlink:   true,
		DotLeader:      true,
		InsertPosition: -1,
		FontFamily:     "",
		FontSize:       0,
		TitleFontSize:  0,
	}
}

// AddTableCaption adds a caption paragraph for a table with SEQ field
// Returns the caption paragraph and the sequence number assigned
func (d *Document) AddTableCaption(captionText string, seqIdentifier string) (*Paragraph, int) {
	if seqIdentifier == "" {
		seqIdentifier = "Table"
	}

	// Get the next sequence number for this identifier
	seqNum := d.getNextSeqNumber(seqIdentifier)

	// Generate a unique bookmark ID for this caption
	bookmarkID := fmt.Sprintf("_Ref%s%d", seqIdentifier, generateUniqueID(fmt.Sprintf("%s%d%s", seqIdentifier, seqNum, captionText)))

	// Create the caption paragraph
	para := &Paragraph{
		Properties: &ParagraphProperties{
			ParagraphStyle: &ParagraphStyle{Val: "Caption"},
			Spacing: &Spacing{
				Before: "120",
				After:  "120",
			},
		},
		Runs: []Run{},
	}

	// Add bookmark start
	bookmarkStart := &BookmarkStart{
		ID:   fmt.Sprintf("%d", seqNum),
		Name: bookmarkID,
	}

	// Store the bookmark for later reference
	d.trackCaptionBookmark(seqIdentifier, seqNum, bookmarkID, captionText)

	// Add the label text (e.g., "Table ")
	para.Runs = append(para.Runs, Run{
		Text: Text{Content: seqIdentifier + " "},
	})

	// Add SEQ field begin
	para.Runs = append(para.Runs, Run{
		FieldChar: &FieldChar{
			FieldCharType: "begin",
		},
	})

	// Add SEQ field instruction
	para.Runs = append(para.Runs, Run{
		InstrText: &InstrText{
			Space:   "preserve",
			Content: fmt.Sprintf(" SEQ %s \\* ARABIC ", seqIdentifier),
		},
	})

	// Add SEQ field separator
	para.Runs = append(para.Runs, Run{
		FieldChar: &FieldChar{
			FieldCharType: "separate",
		},
	})

	// Add the sequence number as text (will be updated by Word)
	para.Runs = append(para.Runs, Run{
		Text: Text{Content: fmt.Sprintf("%d", seqNum)},
	})

	// Add SEQ field end
	para.Runs = append(para.Runs, Run{
		FieldChar: &FieldChar{
			FieldCharType: "end",
		},
	})

	// Add the caption text (e.g., ": Sales Data")
	if captionText != "" {
		para.Runs = append(para.Runs, Run{
			Text: Text{Content: ": " + captionText},
		})
	}

	// Add bookmark end
	bookmarkEnd := &BookmarkEnd{
		ID: fmt.Sprintf("%d", seqNum),
	}

	// Add elements to document
	d.Body.Elements = append(d.Body.Elements, bookmarkStart, para, bookmarkEnd)

	return para, seqNum
}

// AddFigureCaption adds a caption paragraph for a figure with SEQ field
func (d *Document) AddFigureCaption(captionText string) (*Paragraph, int) {
	return d.AddTableCaption(captionText, "Figure")
}

// CaptionBookmark stores information about a caption for LOT/LOF generation
type CaptionBookmark struct {
	SeqIdentifier string
	SeqNumber     int
	BookmarkID    string
	CaptionText   string
}

// trackCaptionBookmark stores caption bookmark information for later LOT/LOF generation
func (d *Document) trackCaptionBookmark(seqIdentifier string, seqNum int, bookmarkID string, captionText string) {
	if d.captionBookmarks == nil {
		d.captionBookmarks = make(map[string][]CaptionBookmark)
	}

	bookmark := CaptionBookmark{
		SeqIdentifier: seqIdentifier,
		SeqNumber:     seqNum,
		BookmarkID:    bookmarkID,
		CaptionText:   captionText,
	}

	d.captionBookmarks[seqIdentifier] = append(d.captionBookmarks[seqIdentifier], bookmark)
}

// getNextSeqNumber returns the next sequence number for a given identifier
func (d *Document) getNextSeqNumber(seqIdentifier string) int {
	if d.seqCounters == nil {
		d.seqCounters = make(map[string]int)
	}

	d.seqCounters[seqIdentifier]++
	return d.seqCounters[seqIdentifier]
}

// GetSeqCount returns the current count for a sequence identifier
func (d *Document) GetSeqCount(seqIdentifier string) int {
	if d.seqCounters == nil {
		return 0
	}
	return d.seqCounters[seqIdentifier]
}

// GenerateLOT generates a List of Tables at the specified position
func (d *Document) GenerateLOT(config *LOTConfig) error {
	if config == nil {
		config = DefaultLOTConfig()
	}

	return d.generateListOfCaptions(config)
}

// GenerateLOF generates a List of Figures at the specified position
func (d *Document) GenerateLOF(config *LOTConfig) error {
	if config == nil {
		config = DefaultLOFConfig()
	}

	return d.generateListOfCaptions(config)
}

// generateListOfCaptions generates a list (LOT or LOF) based on SEQ fields
func (d *Document) generateListOfCaptions(config *LOTConfig) error {
	// Collect all captions with the specified SEQ identifier
	entries := d.collectCaptions(config.SeqIdentifier)

	if len(entries) == 0 {
		// No entries found, but we can still create the TOC field
		// Word will populate it when the document is opened
	}

	// Create the LOT/LOF elements
	lotElements := d.createLOTElements(config, entries)

	// Determine insert position
	insertIndex := config.InsertPosition
	if insertIndex < 0 {
		// Auto: insert after TOC if exists, otherwise at beginning
		_, tocIndex := d.findTOCSDT()
		if tocIndex >= 0 {
			// Find the page break after TOC and insert after it
			insertIndex = tocIndex + 2 // After TOC SDT and page break
			if insertIndex > len(d.Body.Elements) {
				insertIndex = len(d.Body.Elements)
			}
		} else {
			insertIndex = 0
		}
	}

	// Insert the LOT/LOF elements
	if insertIndex >= len(d.Body.Elements) {
		d.Body.Elements = append(d.Body.Elements, lotElements...)
	} else {
		newElements := make([]interface{}, 0, len(d.Body.Elements)+len(lotElements))
		newElements = append(newElements, d.Body.Elements[:insertIndex]...)
		newElements = append(newElements, lotElements...)
		newElements = append(newElements, d.Body.Elements[insertIndex:]...)
		d.Body.Elements = newElements
	}

	return nil
}

// collectCaptions collects all captions with the specified SEQ identifier from tracked bookmarks
func (d *Document) collectCaptions(seqIdentifier string) []LOTEntry {
	var entries []LOTEntry

	// First, check tracked bookmarks (for programmatically added captions)
	if d.captionBookmarks != nil {
		bookmarks, ok := d.captionBookmarks[seqIdentifier]
		if ok {
			for _, bm := range bookmarks {
				entry := LOTEntry{
					Caption:    fmt.Sprintf("%s %d: %s", bm.SeqIdentifier, bm.SeqNumber, bm.CaptionText),
					Number:     bm.SeqNumber,
					PageNum:    1, // Placeholder - Word will update this
					BookmarkID: bm.BookmarkID,
				}
				entries = append(entries, entry)
			}
		}
	}

	// Also scan document for existing SEQ fields (from templates)
	scannedEntries := d.scanForSEQFields(seqIdentifier)
	entries = append(entries, scannedEntries...)

	return entries
}

// scanForSEQFields scans the document for existing SEQ fields and extracts caption information
func (d *Document) scanForSEQFields(seqIdentifier string) []LOTEntry {
	var entries []LOTEntry
	seqNum := 0

	if d.Body == nil || d.Body.Elements == nil {
		return entries
	}

	for _, element := range d.Body.Elements {
		para, ok := element.(*Paragraph)
		if !ok {
			continue
		}

		// Check if this paragraph contains a SEQ field for the specified identifier
		hasSEQ := false
		var captionText strings.Builder

		for _, run := range para.Runs {
			// Check for SEQ field instruction
			if run.InstrText != nil {
				instrContent := strings.TrimSpace(run.InstrText.Content)
				if strings.Contains(instrContent, "SEQ") && strings.Contains(instrContent, seqIdentifier) {
					hasSEQ = true
					seqNum++
				}
			}

			// Collect text content
			if run.Text.Content != "" {
				captionText.WriteString(run.Text.Content)
			}
		}

		if hasSEQ {
			// Extract caption text (remove the "Table X: " prefix if present)
			fullCaption := strings.TrimSpace(captionText.String())
			
			// Generate a bookmark ID for this caption
			bookmarkID := fmt.Sprintf("_RefScanned%s%d", seqIdentifier, generateUniqueID(fullCaption))

			entry := LOTEntry{
				Caption:    fullCaption,
				Number:     seqNum,
				PageNum:    1, // Placeholder - Word will update this
				BookmarkID: bookmarkID,
			}
			entries = append(entries, entry)
		}
	}

	return entries
}

// createLOTElements creates the document elements for a List of Tables/Figures
func (d *Document) createLOTElements(config *LOTConfig, entries []LOTEntry) []interface{} {
	var elements []interface{}

	// Determine font settings
	fontFamily := config.FontFamily
	if fontFamily == "" {
		fontFamily = "Calibri"
	}

	fontSize := config.FontSize
	if fontSize <= 0 {
		fontSize = 11
	}
	fontSizeVal := fmt.Sprintf("%d", fontSize*2) // Word uses half-points

	titleFontSize := config.TitleFontSize
	if titleFontSize <= 0 {
		titleFontSize = 14
	}
	titleFontSizeVal := fmt.Sprintf("%d", titleFontSize*2)

	// Create SDT container for the list
	lotSDT := &SDT{
		Properties: &SDTProperties{
			RunPr: &RunProperties{
				FontFamily: &FontFamily{ASCII: fontFamily, HAnsi: fontFamily, EastAsia: fontFamily, CS: fontFamily},
				FontSize:   &FontSize{Val: fontSizeVal},
			},
			ID:    &SDTID{Val: fmt.Sprintf("%d", generateUniqueID(config.Title))},
			Color: &SDTColor{Val: "DBDBDB"},
			DocPartObj: &DocPartObj{
				DocPartGallery: &DocPartGallery{Val: "Table of Figures"},
				DocPartUnique:  &DocPartUnique{},
			},
		},
		EndPr: &SDTEndPr{
			RunPr: &RunProperties{
				FontFamily: &FontFamily{ASCII: fontFamily, HAnsi: fontFamily, EastAsia: fontFamily, CS: fontFamily},
				Bold:       &Bold{},
				FontSize:   &FontSize{Val: titleFontSizeVal},
			},
		},
		Content: &SDTContent{
			Elements: []interface{}{},
		},
	}

	// Add title paragraph
	titlePara := &Paragraph{
		Properties: &ParagraphProperties{
			Spacing: &Spacing{
				Before: "0",
				After:  "200",
				Line:   "276",
			},
			Justification: &Justification{Val: "left"},
		},
		Runs: []Run{
			{
				Text: Text{Content: config.Title},
				Properties: &RunProperties{
					FontFamily: &FontFamily{ASCII: fontFamily, HAnsi: fontFamily, EastAsia: fontFamily, CS: fontFamily},
					FontSize:   &FontSize{Val: titleFontSizeVal},
					Bold:       &Bold{},
				},
			},
		},
	}

	lotSDT.Content.Elements = append(lotSDT.Content.Elements, titlePara)

	// Create the TOC field paragraph with \c switch for captions
	tocFieldPara := &Paragraph{
		Properties: &ParagraphProperties{
			Tabs: &Tabs{
				Tabs: []TabDef{
					{
						Val:    "right",
						Leader: "dot",
						Pos:    "8640",
					},
				},
			},
		},
		Runs: []Run{},
	}

	// Add TOC field begin
	tocFieldPara.Runs = append(tocFieldPara.Runs, Run{
		Properties: &RunProperties{
			FontFamily: &FontFamily{ASCII: fontFamily, HAnsi: fontFamily, EastAsia: fontFamily, CS: fontFamily},
			FontSize:   &FontSize{Val: fontSizeVal},
		},
		FieldChar: &FieldChar{
			FieldCharType: "begin",
		},
	})

	// Add TOC instruction with \c switch for SEQ identifier
	// The \c switch tells Word to build a list from SEQ fields with the specified identifier
	instrContent := fmt.Sprintf(` TOC \c "%s" `, config.SeqIdentifier)
	if config.UseHyperlink {
		instrContent += `\h `
	}

	tocFieldPara.Runs = append(tocFieldPara.Runs, Run{
		Properties: &RunProperties{
			FontFamily: &FontFamily{ASCII: fontFamily, HAnsi: fontFamily, EastAsia: fontFamily, CS: fontFamily},
			FontSize:   &FontSize{Val: fontSizeVal},
		},
		InstrText: &InstrText{
			Space:   "preserve",
			Content: instrContent,
		},
	})

	// Add TOC field separator
	tocFieldPara.Runs = append(tocFieldPara.Runs, Run{
		Properties: &RunProperties{
			FontFamily: &FontFamily{ASCII: fontFamily, HAnsi: fontFamily, EastAsia: fontFamily, CS: fontFamily},
			FontSize:   &FontSize{Val: fontSizeVal},
		},
		FieldChar: &FieldChar{
			FieldCharType: "separate",
		},
	})

	lotSDT.Content.Elements = append(lotSDT.Content.Elements, tocFieldPara)

	// Add entries (these are placeholders - Word will regenerate them)
	for _, entry := range entries {
		entryPara := d.createLOTEntry(entry, config, fontFamily, fontSizeVal)
		lotSDT.Content.Elements = append(lotSDT.Content.Elements, entryPara)
	}

	// Add TOC field end paragraph
	endPara := &Paragraph{
		Runs: []Run{
			{
				Properties: &RunProperties{
					FontFamily: &FontFamily{ASCII: fontFamily, HAnsi: fontFamily, EastAsia: fontFamily, CS: fontFamily},
					FontSize:   &FontSize{Val: fontSizeVal},
				},
				FieldChar: &FieldChar{
					FieldCharType: "end",
				},
			},
		},
	}

	lotSDT.Content.Elements = append(lotSDT.Content.Elements, endPara)
	elements = append(elements, lotSDT)

	return elements
}

// createLOTEntry creates a single entry paragraph for the LOT/LOF
func (d *Document) createLOTEntry(entry LOTEntry, config *LOTConfig, fontFamily string, fontSizeVal string) *Paragraph {
	para := &Paragraph{
		Properties: &ParagraphProperties{
			ParagraphStyle: &ParagraphStyle{Val: "TableofFigures"},
			Tabs: &Tabs{
				Tabs: []TabDef{
					{
						Val:    "right",
						Leader: "dot",
						Pos:    "8640",
					},
				},
			},
			Spacing: &Spacing{
				Before: "60",
				After:  "60",
				Line:   "276",
			},
		},
		Runs: []Run{},
	}

	runProps := &RunProperties{
		FontFamily: &FontFamily{ASCII: fontFamily, HAnsi: fontFamily, EastAsia: fontFamily, CS: fontFamily},
		FontSize:   &FontSize{Val: fontSizeVal},
	}

	if config.UseHyperlink {
		// Add hyperlink field begin
		para.Runs = append(para.Runs, Run{
			Properties: runProps,
			FieldChar: &FieldChar{
				FieldCharType: "begin",
			},
		})

		// Add hyperlink instruction
		para.Runs = append(para.Runs, Run{
			Properties: runProps,
			InstrText: &InstrText{
				Space:   "preserve",
				Content: fmt.Sprintf(" HYPERLINK \\l %s ", entry.BookmarkID),
			},
		})

		// Add hyperlink separator
		para.Runs = append(para.Runs, Run{
			Properties: runProps,
			FieldChar: &FieldChar{
				FieldCharType: "separate",
			},
		})
	}

	// Add caption text
	para.Runs = append(para.Runs, Run{
		Properties: runProps,
		Text:       Text{Content: entry.Caption},
	})

	// Add tab for page number
	para.Runs = append(para.Runs, Run{
		Properties: runProps,
		Text:       Text{Content: "\t"},
	})

	// Add PAGEREF field for page number
	para.Runs = append(para.Runs, Run{
		Properties: runProps,
		FieldChar: &FieldChar{
			FieldCharType: "begin",
		},
	})

	para.Runs = append(para.Runs, Run{
		Properties: runProps,
		InstrText: &InstrText{
			Space:   "preserve",
			Content: fmt.Sprintf(" PAGEREF %s \\h ", entry.BookmarkID),
		},
	})

	para.Runs = append(para.Runs, Run{
		Properties: runProps,
		FieldChar: &FieldChar{
			FieldCharType: "separate",
		},
	})

	// Page number placeholder
	para.Runs = append(para.Runs, Run{
		Properties: runProps,
		Text:       Text{Content: fmt.Sprintf("%d", entry.PageNum)},
	})

	// PAGEREF field end
	para.Runs = append(para.Runs, Run{
		Properties: runProps,
		FieldChar: &FieldChar{
			FieldCharType: "end",
		},
	})

	if config.UseHyperlink {
		// Hyperlink field end
		para.Runs = append(para.Runs, Run{
			Properties: runProps,
			FieldChar: &FieldChar{
				FieldCharType: "end",
			},
		})
	}

	return para
}

// InsertTableCaptionBefore inserts a table caption before a specific element index
func (d *Document) InsertTableCaptionBefore(captionText string, seqIdentifier string, beforeIndex int) (*Paragraph, int) {
	if seqIdentifier == "" {
		seqIdentifier = "Table"
	}

	// Get the next sequence number
	seqNum := d.getNextSeqNumber(seqIdentifier)

	// Generate bookmark ID
	bookmarkID := fmt.Sprintf("_Ref%s%d", seqIdentifier, generateUniqueID(fmt.Sprintf("%s%d%s", seqIdentifier, seqNum, captionText)))

	// Create caption paragraph
	para := &Paragraph{
		Properties: &ParagraphProperties{
			ParagraphStyle: &ParagraphStyle{Val: "Caption"},
			Spacing: &Spacing{
				Before: "120",
				After:  "120",
			},
		},
		Runs: []Run{},
	}

	// Add label text
	para.Runs = append(para.Runs, Run{
		Text: Text{Content: seqIdentifier + " "},
	})

	// Add SEQ field
	para.Runs = append(para.Runs, Run{
		FieldChar: &FieldChar{FieldCharType: "begin"},
	})
	para.Runs = append(para.Runs, Run{
		InstrText: &InstrText{
			Space:   "preserve",
			Content: fmt.Sprintf(" SEQ %s \\* ARABIC ", seqIdentifier),
		},
	})
	para.Runs = append(para.Runs, Run{
		FieldChar: &FieldChar{FieldCharType: "separate"},
	})
	para.Runs = append(para.Runs, Run{
		Text: Text{Content: fmt.Sprintf("%d", seqNum)},
	})
	para.Runs = append(para.Runs, Run{
		FieldChar: &FieldChar{FieldCharType: "end"},
	})

	// Add caption text
	if captionText != "" {
		para.Runs = append(para.Runs, Run{
			Text: Text{Content: ": " + captionText},
		})
	}

	// Track the caption
	d.trackCaptionBookmark(seqIdentifier, seqNum, bookmarkID, captionText)

	// Create bookmark elements
	bookmarkStart := &BookmarkStart{
		ID:   fmt.Sprintf("%d", seqNum+1000), // Offset to avoid conflicts
		Name: bookmarkID,
	}
	bookmarkEnd := &BookmarkEnd{
		ID: fmt.Sprintf("%d", seqNum+1000),
	}

	// Insert at the specified position
	if beforeIndex < 0 {
		beforeIndex = 0
	}
	if beforeIndex > len(d.Body.Elements) {
		beforeIndex = len(d.Body.Elements)
	}

	// Insert bookmark start, caption, bookmark end
	newElements := make([]interface{}, 0, len(d.Body.Elements)+3)
	newElements = append(newElements, d.Body.Elements[:beforeIndex]...)
	newElements = append(newElements, bookmarkStart, para, bookmarkEnd)
	newElements = append(newElements, d.Body.Elements[beforeIndex:]...)
	d.Body.Elements = newElements

	return para, seqNum
}

// GetLOTTitleForLanguage returns the default LOT title for a language
func GetLOTTitleForLanguage(lang string) string {
	switch strings.ToLower(lang) {
	case "es", "spanish":
		return "Lista de Tablas"
	case "en", "english":
		return "List of Tables"
	case "ca", "catalan":
		return "Llista de Taules"
	case "fr", "french":
		return "Liste des Tableaux"
	case "de", "german":
		return "Tabellenverzeichnis"
	case "pt", "portuguese":
		return "Lista de Tabelas"
	default:
		return "List of Tables"
	}
}

// GetLOFTitleForLanguage returns the default LOF title for a language
func GetLOFTitleForLanguage(lang string) string {
	switch strings.ToLower(lang) {
	case "es", "spanish":
		return "Lista de Figuras"
	case "en", "english":
		return "List of Figures"
	case "ca", "catalan":
		return "Llista de Figures"
	case "fr", "french":
		return "Liste des Figures"
	case "de", "german":
		return "Abbildungsverzeichnis"
	case "pt", "portuguese":
		return "Lista de Figuras"
	default:
		return "List of Figures"
	}
}
