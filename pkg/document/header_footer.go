// Package document 提供Word文档的页眉页脚操作功能
package document

import (
	"encoding/xml"
	"fmt"
	"strconv"
	"strings"
)

// HeaderFooterType 页眉页脚类型
type HeaderFooterType string

const (
	// HeaderFooterTypeDefault 默认页眉页脚
	HeaderFooterTypeDefault HeaderFooterType = "default"
	// HeaderFooterTypeFirst 首页页眉页脚
	HeaderFooterTypeFirst HeaderFooterType = "first"
	// HeaderFooterTypeEven 偶数页页眉页脚
	HeaderFooterTypeEven HeaderFooterType = "even"
)

// Header 页眉结构
type Header struct {
	XMLName     xml.Name     `xml:"w:hdr"`
	XmlnsWPC    string       `xml:"xmlns:wpc,attr"`
	XmlnsMC     string       `xml:"xmlns:mc,attr"`
	XmlnsO      string       `xml:"xmlns:o,attr"`
	XmlnsR      string       `xml:"xmlns:r,attr"`
	XmlnsM      string       `xml:"xmlns:m,attr"`
	XmlnsV      string       `xml:"xmlns:v,attr"`
	XmlnsWP14   string       `xml:"xmlns:wp14,attr"`
	XmlnsWP     string       `xml:"xmlns:wp,attr"`
	XmlnsW10    string       `xml:"xmlns:w10,attr"`
	XmlnsW      string       `xml:"xmlns:w,attr"`
	XmlnsW14    string       `xml:"xmlns:w14,attr"`
	XmlnsW15    string       `xml:"xmlns:w15,attr"`
	XmlnsWPG    string       `xml:"xmlns:wpg,attr"`
	XmlnsWPI    string       `xml:"xmlns:wpi,attr"`
	XmlnsWNE    string       `xml:"xmlns:wne,attr"`
	XmlnsWPS    string       `xml:"xmlns:wps,attr"`
	XmlnsWPSCD  string       `xml:"xmlns:wpsCustomData,attr"`
	MCIgnorable string       `xml:"mc:Ignorable,attr"`
	Paragraphs  []*Paragraph `xml:"w:p"`
}

// Footer 页脚结构
type Footer struct {
	XMLName     xml.Name     `xml:"w:ftr"`
	XmlnsWPC    string       `xml:"xmlns:wpc,attr"`
	XmlnsMC     string       `xml:"xmlns:mc,attr"`
	XmlnsO      string       `xml:"xmlns:o,attr"`
	XmlnsR      string       `xml:"xmlns:r,attr"`
	XmlnsM      string       `xml:"xmlns:m,attr"`
	XmlnsV      string       `xml:"xmlns:v,attr"`
	XmlnsWP14   string       `xml:"xmlns:wp14,attr"`
	XmlnsWP     string       `xml:"xmlns:wp,attr"`
	XmlnsW10    string       `xml:"xmlns:w10,attr"`
	XmlnsW      string       `xml:"xmlns:w,attr"`
	XmlnsW14    string       `xml:"xmlns:w14,attr"`
	XmlnsW15    string       `xml:"xmlns:w15,attr"`
	XmlnsWPG    string       `xml:"xmlns:wpg,attr"`
	XmlnsWPI    string       `xml:"xmlns:wpi,attr"`
	XmlnsWNE    string       `xml:"xmlns:wne,attr"`
	XmlnsWPS    string       `xml:"xmlns:wps,attr"`
	XmlnsWPSCD  string       `xml:"xmlns:wpsCustomData,attr"`
	MCIgnorable string       `xml:"mc:Ignorable,attr"`
	Paragraphs  []*Paragraph `xml:"w:p"`
}

// HeaderFooterReference 页眉页脚引用
type HeaderFooterReference struct {
	XMLName xml.Name `xml:"w:headerReference"`
	Type    string   `xml:"w:type,attr"`
	ID      string   `xml:"r:id,attr"`
}

// FooterReference 页脚引用
type FooterReference struct {
	XMLName xml.Name `xml:"w:footerReference"`
	Type    string   `xml:"w:type,attr"`
	ID      string   `xml:"r:id,attr"`
}

// TitlePage 首页不同设置
type TitlePage struct {
	XMLName xml.Name `xml:"w:titlePg"`
}

// PageNumber 页码字段
type PageNumber struct {
	XMLName xml.Name `xml:"w:fldSimple"`
	Instr   string   `xml:"w:instr,attr"`
	Text    *Text    `xml:"w:t,omitempty"`
}

// createStandardHeader 创建标准页眉结构
func createStandardHeader() *Header {
	return &Header{
		XmlnsWPC:    "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
		XmlnsMC:     "http://schemas.openxmlformats.org/markup-compatibility/2006",
		XmlnsO:      "urn:schemas-microsoft-com:office:office",
		XmlnsR:      "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
		XmlnsM:      "http://schemas.openxmlformats.org/officeDocument/2006/math",
		XmlnsV:      "urn:schemas-microsoft-com:vml",
		XmlnsWP14:   "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
		XmlnsWP:     "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
		XmlnsW10:    "urn:schemas-microsoft-com:office:word",
		XmlnsW:      "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
		XmlnsW14:    "http://schemas.microsoft.com/office/word/2010/wordml",
		XmlnsW15:    "http://schemas.microsoft.com/office/word/2012/wordml",
		XmlnsWPG:    "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
		XmlnsWPI:    "http://schemas.microsoft.com/office/word/2010/wordprocessingInk",
		XmlnsWNE:    "http://schemas.microsoft.com/office/word/2006/wordml",
		XmlnsWPS:    "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
		XmlnsWPSCD:  "http://www.wps.cn/officeDocument/2013/wpsCustomData",
		MCIgnorable: "w14 w15 wp14",
		Paragraphs:  make([]*Paragraph, 0),
	}
}

// createStandardFooter 创建标准页脚结构
func createStandardFooter() *Footer {
	return &Footer{
		XmlnsWPC:    "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
		XmlnsMC:     "http://schemas.openxmlformats.org/markup-compatibility/2006",
		XmlnsO:      "urn:schemas-microsoft-com:office:office",
		XmlnsR:      "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
		XmlnsM:      "http://schemas.openxmlformats.org/officeDocument/2006/math",
		XmlnsV:      "urn:schemas-microsoft-com:vml",
		XmlnsWP14:   "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
		XmlnsWP:     "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
		XmlnsW10:    "urn:schemas-microsoft-com:office:word",
		XmlnsW:      "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
		XmlnsW14:    "http://schemas.microsoft.com/office/word/2010/wordml",
		XmlnsW15:    "http://schemas.microsoft.com/office/word/2012/wordml",
		XmlnsWPG:    "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
		XmlnsWPI:    "http://schemas.microsoft.com/office/word/2010/wordprocessingInk",
		XmlnsWNE:    "http://schemas.microsoft.com/office/word/2006/wordml",
		XmlnsWPS:    "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
		XmlnsWPSCD:  "http://www.wps.cn/officeDocument/2013/wpsCustomData",
		MCIgnorable: "w14 w15 wp14",
		Paragraphs:  make([]*Paragraph, 0),
	}
}

// createPageNumberRuns 创建页码域代码的Run集合
func createPageNumberRuns() []Run {
	return []Run{
		{
			FieldChar: &FieldChar{
				FieldCharType: "begin",
			},
		},
		{
			InstrText: &InstrText{
				Space:   "preserve",
				Content: " PAGE  \\* MERGEFORMAT ",
			},
		},
		{
			FieldChar: &FieldChar{
				FieldCharType: "separate",
			},
		},
		{
			Text: Text{
				Content: "1",
			},
		},
		{
			FieldChar: &FieldChar{
				FieldCharType: "end",
			},
		},
	}
}

// createStyledPageNumberRuns creates page number runs with Open Sans font and black color
func createStyledPageNumberRuns() []Run {
	// Run properties for Open Sans font, black color, 10pt size
	runProps := &RunProperties{
		FontFamily: &FontFamily{
			ASCII:    "Open Sans",
			HAnsi:    "Open Sans",
			EastAsia: "Open Sans",
			CS:       "Open Sans",
		},
		Color: &Color{
			Val: "000000",
		},
		FontSize: &FontSize{
			Val: "20", // 10pt = 20 half-points
		},
		FontSizeCs: &FontSizeCs{
			Val: "20",
		},
	}

	return []Run{
		{
			Properties: runProps,
			FieldChar: &FieldChar{
				FieldCharType: "begin",
			},
		},
		{
			Properties: runProps,
			InstrText: &InstrText{
				Space:   "preserve",
				Content: " PAGE ",
			},
		},
		{
			Properties: runProps,
			FieldChar: &FieldChar{
				FieldCharType: "separate",
			},
		},
		{
			Properties: runProps,
			Text: Text{
				Content: "1",
			},
		},
		{
			Properties: runProps,
			FieldChar: &FieldChar{
				FieldCharType: "end",
			},
		},
	}
}

// getFileNameForType 获取页眉页脚文件名
func getFileNameForType(typePrefix string, headerType HeaderFooterType) string {
	switch headerType {
	case HeaderFooterTypeDefault:
		return fmt.Sprintf("%s1.xml", typePrefix)
	case HeaderFooterTypeFirst:
		return fmt.Sprintf("%sfirst.xml", typePrefix)
	case HeaderFooterTypeEven:
		return fmt.Sprintf("%seven.xml", typePrefix)
	default:
		return fmt.Sprintf("%s1.xml", typePrefix)
	}
}

// AddHeader 添加页眉
func (d *Document) AddHeader(headerType HeaderFooterType, text string) error {
	header := createStandardHeader()

	// 创建页眉段落
	paragraph := &Paragraph{}
	if text != "" {
		run := Run{
			Text: Text{
				Content: text,
				Space:   "preserve",
			},
		}
		paragraph.Runs = append(paragraph.Runs, run)
	}
	header.Paragraphs = append(header.Paragraphs, paragraph)

	// 生成关系ID
	headerID := fmt.Sprintf("rId%d", len(d.documentRelationships.Relationships)+2) // +2因为rId1保留给styles

	// 序列化页眉
	headerXML, err := xml.MarshalIndent(header, "", "  ")
	if err != nil {
		return fmt.Errorf("序列化页眉失败: %v", err)
	}

	// 添加XML声明
	fullXML := append([]byte(xml.Header), headerXML...)

	// 获取文件名
	fileName := getFileNameForType("header", headerType)
	headerPartName := fmt.Sprintf("word/%s", fileName)

	// 存储页眉内容
	d.parts[headerPartName] = fullXML

	// 添加关系到文档关系
	relationship := Relationship{
		ID:     headerID,
		Type:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header",
		Target: fileName,
	}
	d.documentRelationships.Relationships = append(d.documentRelationships.Relationships, relationship)

	// 添加内容类型
	d.addContentType(headerPartName, "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml")

	// 更新节属性
	d.addHeaderReference(headerType, headerID)

	return nil
}

// AddFooter 添加页脚
func (d *Document) AddFooter(footerType HeaderFooterType, text string) error {
	footer := createStandardFooter()

	// 创建页脚段落
	paragraph := &Paragraph{}
	if text != "" {
		run := Run{
			Text: Text{
				Content: text,
				Space:   "preserve",
			},
		}
		paragraph.Runs = append(paragraph.Runs, run)
	}
	footer.Paragraphs = append(footer.Paragraphs, paragraph)

	// 生成关系ID
	footerID := fmt.Sprintf("rId%d", len(d.documentRelationships.Relationships)+2) // +2因为rId1保留给styles

	// 序列化页脚
	footerXML, err := xml.MarshalIndent(footer, "", "  ")
	if err != nil {
		return fmt.Errorf("序列化页脚失败: %v", err)
	}

	// 添加XML声明
	fullXML := append([]byte(xml.Header), footerXML...)

	// 获取文件名
	fileName := getFileNameForType("footer", footerType)
	footerPartName := fmt.Sprintf("word/%s", fileName)

	// 存储页脚内容
	d.parts[footerPartName] = fullXML

	// 添加关系到文档关系
	relationship := Relationship{
		ID:     footerID,
		Type:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer",
		Target: fileName,
	}
	d.documentRelationships.Relationships = append(d.documentRelationships.Relationships, relationship)

	// 添加内容类型
	d.addContentType(footerPartName, "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml")

	// 更新节属性
	d.addFooterReference(footerType, footerID)

	return nil
}

// AddHeaderWithPageNumber 添加带页码的页眉
func (d *Document) AddHeaderWithPageNumber(headerType HeaderFooterType, text string, showPageNum bool) error {
	header := createStandardHeader()

	// 创建页眉段落
	paragraph := &Paragraph{}

	if text != "" {
		run := Run{
			Text: Text{
				Content: text,
				Space:   "preserve",
			},
		}
		paragraph.Runs = append(paragraph.Runs, run)
	}

	if showPageNum {
		// 添加"第"字
		pageNumRun := Run{
			Text: Text{
				Content: " 第 ",
				Space:   "preserve",
			},
		}
		paragraph.Runs = append(paragraph.Runs, pageNumRun)

		// 添加页码域代码
		pageNumberRuns := createPageNumberRuns()
		paragraph.Runs = append(paragraph.Runs, pageNumberRuns...)

		// 添加"页"字
		pageNumRun2 := Run{
			Text: Text{
				Content: " 页",
				Space:   "preserve",
			},
		}
		paragraph.Runs = append(paragraph.Runs, pageNumRun2)
	}

	header.Paragraphs = append(header.Paragraphs, paragraph)

	// 生成关系ID
	headerID := fmt.Sprintf("rId%d", len(d.documentRelationships.Relationships)+2) // +2因为rId1保留给styles

	// 序列化页眉
	headerXML, err := xml.MarshalIndent(header, "", "  ")
	if err != nil {
		return fmt.Errorf("序列化页眉失败: %v", err)
	}

	// 添加XML声明
	fullXML := append([]byte(xml.Header), headerXML...)

	// 获取文件名
	fileName := getFileNameForType("header", headerType)
	headerPartName := fmt.Sprintf("word/%s", fileName)

	// 存储页眉内容
	d.parts[headerPartName] = fullXML

	// 添加关系到文档关系
	relationship := Relationship{
		ID:     headerID,
		Type:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header",
		Target: fileName,
	}
	d.documentRelationships.Relationships = append(d.documentRelationships.Relationships, relationship)

	// 添加内容类型
	d.addContentType(headerPartName, "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml")

	// 更新节属性
	d.addHeaderReference(headerType, headerID)

	return nil
}

// MergePageNumberIntoFooter adds a page number to an existing footer, or creates a new footer if none exists.
// This preserves any existing footer content from templates by using raw XML manipulation.
func (d *Document) MergePageNumberIntoFooter(footerType HeaderFooterType) error {
	// First, try to find the existing footer from section properties
	var footerID string
	var footerTarget string
	
	// Look up the footer reference from section properties
	sectPr := d.getSectionPropertiesForHeaderFooter()
	if sectPr != nil && sectPr.FooterReferences != nil {
		for _, ref := range sectPr.FooterReferences {
			if ref.Type == string(footerType) {
				footerID = ref.ID
				break
			}
		}
	}
	
	// If we found a footer reference, look up the actual file target
	if footerID != "" && d.documentRelationships != nil {
		for _, rel := range d.documentRelationships.Relationships {
			if rel.ID == footerID {
				footerTarget = rel.Target
				break
			}
		}
	}
	
	// If we found an existing footer, modify it
	if footerTarget != "" {
		footerPartName := "word/" + footerTarget
		if existingFooterXML, exists := d.parts[footerPartName]; exists {
			fmt.Printf("[MergePageNumberIntoFooter] Found existing footer: %s (rId: %s)\n", footerTarget, footerID)
			
			// Use raw XML manipulation to preserve complex content
			// Find the closing </w:ftr> tag and insert the page number paragraph before it
			pageNumXML := createPageNumberParagraphXML()
			
			xmlStr := string(existingFooterXML)
			closingTag := "</w:ftr>"
			closingIdx := strings.LastIndex(xmlStr, closingTag)
			if closingIdx == -1 {
				// Try alternative closing tag format
				closingTag = "</ftr>"
				closingIdx = strings.LastIndex(xmlStr, closingTag)
			}
			
			if closingIdx != -1 {
				// Insert page number paragraph before the closing tag
				newXML := xmlStr[:closingIdx] + pageNumXML + xmlStr[closingIdx:]
				d.parts[footerPartName] = []byte(newXML)
				fmt.Printf("[MergePageNumberIntoFooter] Added page number to existing footer\n")
				return nil
			}
			
			fmt.Printf("[MergePageNumberIntoFooter] Warning: Could not find closing tag in %s, using fallback method\n", footerTarget)
		}
	}
	
	// Fallback: use the default filename approach
	fileName := getFileNameForType("footer", footerType)
	footerPartName := fmt.Sprintf("word/%s", fileName)
	
	// Check if footer already exists in parts with default name
	if existingFooterXML, exists := d.parts[footerPartName]; exists {
		// Find the existing relationship ID for this footer
		for _, rel := range d.documentRelationships.Relationships {
			if rel.Target == fileName {
				footerID = rel.ID
				break
			}
		}
		
		// Use raw XML manipulation to preserve complex content
		pageNumXML := createPageNumberParagraphXML()
		
		xmlStr := string(existingFooterXML)
		closingTag := "</w:ftr>"
		closingIdx := strings.LastIndex(xmlStr, closingTag)
		if closingIdx == -1 {
			closingTag = "</ftr>"
			closingIdx = strings.LastIndex(xmlStr, closingTag)
		}
		
		if closingIdx != -1 {
			newXML := xmlStr[:closingIdx] + pageNumXML + xmlStr[closingIdx:]
			d.parts[footerPartName] = []byte(newXML)
			return nil
		}
		
		fmt.Printf("[MergePageNumberIntoFooter] Warning: Could not find closing tag, using fallback method\n")
	}
	
	// No existing footer - create a new one
	fmt.Printf("[MergePageNumberIntoFooter] Creating new footer: %s\n", fileName)
	footer := createStandardFooter()
	
	// Create page number paragraph with centered alignment
	pageNumParagraph := &Paragraph{
		Properties: &ParagraphProperties{
			Justification: &Justification{Val: "center"},
		},
	}
	
	// Add styled page number runs
	pageNumberRuns := createStyledPageNumberRuns()
	pageNumParagraph.Runs = append(pageNumParagraph.Runs, pageNumberRuns...)
	
	// Append page number paragraph to footer
	footer.Paragraphs = append(footer.Paragraphs, pageNumParagraph)
	
	// Generate new relationship ID if needed
	if footerID == "" {
		footerID = fmt.Sprintf("rId%d", len(d.documentRelationships.Relationships)+2)
		
		// Add relationship
		relationship := Relationship{
			ID:     footerID,
			Type:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer",
			Target: fileName,
		}
		d.documentRelationships.Relationships = append(d.documentRelationships.Relationships, relationship)
		
		// Add content type
		d.addContentType(footerPartName, "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml")
		
		// Add footer reference to section properties
		d.addFooterReference(footerType, footerID)
	}
	
	// Serialize and store footer
	footerXML, err := xml.MarshalIndent(footer, "", "  ")
	if err != nil {
		return fmt.Errorf("failed to serialize footer: %v", err)
	}
	
	fullXML := append([]byte(xml.Header), footerXML...)
	d.parts[footerPartName] = fullXML
	
	return nil
}

// createPageNumberParagraphXML creates the XML for a centered page number paragraph
func createPageNumberParagraphXML() string {
	return `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:pPr>
    <w:jc w:val="center"/>
  </w:pPr>
  <w:r>
    <w:rPr>
      <w:rFonts w:ascii="Aptos" w:hAnsi="Aptos"/>
      <w:color w:val="000000"/>
    </w:rPr>
    <w:fldChar w:fldCharType="begin"/>
  </w:r>
  <w:r>
    <w:rPr>
      <w:rFonts w:ascii="Aptos" w:hAnsi="Aptos"/>
      <w:color w:val="000000"/>
    </w:rPr>
    <w:instrText xml:space="preserve"> PAGE </w:instrText>
  </w:r>
  <w:r>
    <w:rPr>
      <w:rFonts w:ascii="Aptos" w:hAnsi="Aptos"/>
      <w:color w:val="000000"/>
    </w:rPr>
    <w:fldChar w:fldCharType="separate"/>
  </w:r>
  <w:r>
    <w:rPr>
      <w:rFonts w:ascii="Aptos" w:hAnsi="Aptos"/>
      <w:color w:val="000000"/>
    </w:rPr>
    <w:t>1</w:t>
  </w:r>
  <w:r>
    <w:rPr>
      <w:rFonts w:ascii="Aptos" w:hAnsi="Aptos"/>
      <w:color w:val="000000"/>
    </w:rPr>
    <w:fldChar w:fldCharType="end"/>
  </w:r>
</w:p>`
}

// MergePageNumberIntoHeader adds a page number to an existing header, or creates a new header if none exists.
// This preserves any existing header content from templates.
func (d *Document) MergePageNumberIntoHeader(headerType HeaderFooterType) error {
	fileName := getFileNameForType("header", headerType)
	headerPartName := fmt.Sprintf("word/%s", fileName)
	
	var header *Header
	var headerID string
	
	// Check if header already exists in parts
	if existingHeaderXML, exists := d.parts[headerPartName]; exists {
		// Parse existing header
		header = &Header{}
		if err := xml.Unmarshal(existingHeaderXML, header); err != nil {
			// If parsing fails, create a new header
			header = createStandardHeader()
		}
		
		// Find the existing relationship ID for this header
		for _, rel := range d.documentRelationships.Relationships {
			if rel.Target == fileName {
				headerID = rel.ID
				break
			}
		}
	} else {
		// No existing header, create a new one
		header = createStandardHeader()
	}
	
	// Create page number paragraph with right alignment (typical for headers)
	pageNumParagraph := &Paragraph{
		Properties: &ParagraphProperties{
			Justification: &Justification{Val: "right"},
		},
	}
	
	// Add styled page number runs
	pageNumberRuns := createStyledPageNumberRuns()
	pageNumParagraph.Runs = append(pageNumParagraph.Runs, pageNumberRuns...)
	
	// Append page number paragraph to existing header content
	header.Paragraphs = append(header.Paragraphs, pageNumParagraph)
	
	// Generate new relationship ID if needed
	if headerID == "" {
		headerID = fmt.Sprintf("rId%d", len(d.documentRelationships.Relationships)+2)
		
		// Add relationship
		relationship := Relationship{
			ID:     headerID,
			Type:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header",
			Target: fileName,
		}
		d.documentRelationships.Relationships = append(d.documentRelationships.Relationships, relationship)
		
		// Add content type
		d.addContentType(headerPartName, "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml")
		
		// Add header reference to section properties
		d.addHeaderReference(headerType, headerID)
	}
	
	// Serialize and store header
	headerXML, err := xml.MarshalIndent(header, "", "  ")
	if err != nil {
		return fmt.Errorf("failed to serialize header: %v", err)
	}
	
	fullXML := append([]byte(xml.Header), headerXML...)
	d.parts[headerPartName] = fullXML
	
	return nil
}

// AddFooterWithPageNumber 添加带页码的页脚
func (d *Document) AddFooterWithPageNumber(footerType HeaderFooterType, text string, showPageNum bool) error {
	footer := createStandardFooter()

	// 创建页脚段落 - centered alignment
	paragraph := &Paragraph{
		Properties: &ParagraphProperties{
			Justification: &Justification{Val: "center"},
		},
	}

	if text != "" {
		run := Run{
			Text: Text{
				Content: text,
				Space:   "preserve",
			},
		}
		paragraph.Runs = append(paragraph.Runs, run)
	}

	if showPageNum {
		// Create page number runs with Open Sans font and black color
		pageNumberRuns := createStyledPageNumberRuns()
		paragraph.Runs = append(paragraph.Runs, pageNumberRuns...)
	}

	footer.Paragraphs = append(footer.Paragraphs, paragraph)

	// 生成关系ID
	footerID := fmt.Sprintf("rId%d", len(d.documentRelationships.Relationships)+2) // +2因为rId1保留给styles

	// 序列化页脚
	footerXML, err := xml.MarshalIndent(footer, "", "  ")
	if err != nil {
		return fmt.Errorf("序列化页脚失败: %v", err)
	}

	// 添加XML声明
	fullXML := append([]byte(xml.Header), footerXML...)

	// 获取文件名
	fileName := getFileNameForType("footer", footerType)
	footerPartName := fmt.Sprintf("word/%s", fileName)

	// 存储页脚内容
	d.parts[footerPartName] = fullXML

	// 添加关系到文档关系
	relationship := Relationship{
		ID:     footerID,
		Type:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer",
		Target: fileName,
	}
	d.documentRelationships.Relationships = append(d.documentRelationships.Relationships, relationship)

	// 添加内容类型
	d.addContentType(footerPartName, "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml")

	// 更新节属性
	d.addFooterReference(footerType, footerID)

	return nil
}

// HeaderFooterConfig 页眉页脚配置
type HeaderFooterConfig struct {
	Text      string        // 文本内容
	Format    *TextFormat   // 文本格式配置
	Alignment AlignmentType // 对齐方式
}

// createFormattedParagraph 创建格式化的段落
func createFormattedParagraph(text string, format *TextFormat, alignment AlignmentType) *Paragraph {
	paragraph := &Paragraph{}

	// 设置段落对齐方式
	if alignment != "" {
		paragraph.Properties = &ParagraphProperties{
			Justification: &Justification{Val: string(alignment)},
		}
	}

	// 如果有文本内容，创建带格式的Run
	if text != "" {
		run := Run{
			Text: Text{
				Content: text,
				Space:   "preserve",
			},
		}

		// 应用文本格式
		if format != nil {
			runProps := &RunProperties{}

			// 设置字体
			fontName := ""
			if format.FontFamily != "" {
				fontName = format.FontFamily
			} else if format.FontName != "" {
				fontName = format.FontName
			}
			if fontName != "" {
				runProps.FontFamily = &FontFamily{
					ASCII:    fontName,
					HAnsi:    fontName,
					EastAsia: fontName,
					CS:       fontName,
				}
			}

			// 设置粗体
			if format.Bold {
				runProps.Bold = &Bold{}
			}

			// 设置斜体
			if format.Italic {
				runProps.Italic = &Italic{}
			}

			// 设置字体颜色
			if format.FontColor != "" {
				// 确保颜色格式正确（移除#前缀）
				color := strings.TrimPrefix(format.FontColor, "#")
				runProps.Color = &Color{Val: color}
			}

			// 设置字体大小
			if format.FontSize > 0 {
				// Word中字体大小是半磅为单位，所以需要乘以2
				runProps.FontSize = &FontSize{Val: strconv.Itoa(format.FontSize * 2)}
			}

			// 设置下划线
			if format.Underline {
				runProps.Underline = &Underline{Val: "single"}
			}

			// 设置删除线
			if format.Strike {
				runProps.Strike = &Strike{}
			}

			// 设置高亮
			if format.Highlight != "" {
				runProps.Highlight = &Highlight{Val: format.Highlight}
			}

			run.Properties = runProps
		}

		paragraph.Runs = append(paragraph.Runs, run)
	}

	return paragraph
}

// AddFormattedHeader 添加格式化页眉
//
// 该方法允许添加带有自定义文本格式和对齐方式的页眉。
//
// 参数:
//   - headerType: 页眉类型 (HeaderFooterTypeDefault, HeaderFooterTypeFirst, HeaderFooterTypeEven)
//   - config: 页眉配置，包含文本内容、格式和对齐方式
//
// 示例:
//
//	doc.AddFormattedHeader(document.HeaderFooterTypeDefault, &document.HeaderFooterConfig{
//		Text: "公司报告",
//		Format: &document.TextFormat{
//			FontSize:   10,
//			FontColor:  "8e8e8e",
//			FontFamily: "Arial",
//		},
//		Alignment: document.AlignCenter,
//	})
func (d *Document) AddFormattedHeader(headerType HeaderFooterType, config *HeaderFooterConfig) error {
	header := createStandardHeader()

	// 创建格式化页眉段落
	if config == nil {
		config = &HeaderFooterConfig{}
	}
	paragraph := createFormattedParagraph(config.Text, config.Format, config.Alignment)
	header.Paragraphs = append(header.Paragraphs, paragraph)

	// 生成关系ID
	headerID := fmt.Sprintf("rId%d", len(d.documentRelationships.Relationships)+2) // +2因为rId1保留给styles

	// 序列化页眉
	headerXML, err := xml.MarshalIndent(header, "", "  ")
	if err != nil {
		return fmt.Errorf("序列化页眉失败: %v", err)
	}

	// 添加XML声明
	fullXML := append([]byte(xml.Header), headerXML...)

	// 获取文件名
	fileName := getFileNameForType("header", headerType)
	headerPartName := fmt.Sprintf("word/%s", fileName)

	// 存储页眉内容
	d.parts[headerPartName] = fullXML

	// 添加关系到文档关系
	relationship := Relationship{
		ID:     headerID,
		Type:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header",
		Target: fileName,
	}
	d.documentRelationships.Relationships = append(d.documentRelationships.Relationships, relationship)

	// 添加内容类型
	d.addContentType(headerPartName, "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml")

	// 更新节属性
	d.addHeaderReference(headerType, headerID)

	return nil
}

// AddFormattedFooter 添加格式化页脚
//
// 该方法允许添加带有自定义文本格式和对齐方式的页脚。
//
// 参数:
//   - footerType: 页脚类型 (HeaderFooterTypeDefault, HeaderFooterTypeFirst, HeaderFooterTypeEven)
//   - config: 页脚配置，包含文本内容、格式和对齐方式
//
// 示例:
//
//	doc.AddFormattedFooter(document.HeaderFooterTypeDefault, &document.HeaderFooterConfig{
//		Text: "第 1 页",
//		Format: &document.TextFormat{
//			FontSize:   9,
//			FontColor:  "666666",
//			FontFamily: "宋体",
//		},
//		Alignment: document.AlignCenter,
//	})
func (d *Document) AddFormattedFooter(footerType HeaderFooterType, config *HeaderFooterConfig) error {
	footer := createStandardFooter()

	// 创建格式化页脚段落
	if config == nil {
		config = &HeaderFooterConfig{}
	}
	paragraph := createFormattedParagraph(config.Text, config.Format, config.Alignment)
	footer.Paragraphs = append(footer.Paragraphs, paragraph)

	// 生成关系ID
	footerID := fmt.Sprintf("rId%d", len(d.documentRelationships.Relationships)+2) // +2因为rId1保留给styles

	// 序列化页脚
	footerXML, err := xml.MarshalIndent(footer, "", "  ")
	if err != nil {
		return fmt.Errorf("序列化页脚失败: %v", err)
	}

	// 添加XML声明
	fullXML := append([]byte(xml.Header), footerXML...)

	// 获取文件名
	fileName := getFileNameForType("footer", footerType)
	footerPartName := fmt.Sprintf("word/%s", fileName)

	// 存储页脚内容
	d.parts[footerPartName] = fullXML

	// 添加关系到文档关系
	relationship := Relationship{
		ID:     footerID,
		Type:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer",
		Target: fileName,
	}
	d.documentRelationships.Relationships = append(d.documentRelationships.Relationships, relationship)

	// 添加内容类型
	d.addContentType(footerPartName, "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml")

	// 更新节属性
	d.addFooterReference(footerType, footerID)

	return nil
}

// SetDifferentFirstPage 设置首页不同
func (d *Document) SetDifferentFirstPage(different bool) {
	sectPr := d.getSectionPropertiesForHeaderFooter()
	if different {
		sectPr.TitlePage = &TitlePage{}
	} else {
		sectPr.TitlePage = nil
	}
}

// addHeaderReference 添加页眉引用到节属性
func (d *Document) addHeaderReference(headerType HeaderFooterType, headerID string) {
	sectPr := d.getSectionPropertiesForHeaderFooter()

	// 确保设置关系命名空间
	if sectPr.XmlnsR == "" {
		sectPr.XmlnsR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
	}

	headerRef := &HeaderFooterReference{
		Type: string(headerType),
		ID:   headerID,
	}

	sectPr.HeaderReferences = append(sectPr.HeaderReferences, headerRef)
}

// addFooterReference 添加页脚引用到节属性
func (d *Document) addFooterReference(footerType HeaderFooterType, footerID string) {
	sectPr := d.getSectionPropertiesForHeaderFooter()

	// 确保设置关系命名空间
	if sectPr.XmlnsR == "" {
		sectPr.XmlnsR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
	}

	footerRef := &FooterReference{
		Type: string(footerType),
		ID:   footerID,
	}

	// Check if a footer reference of this type already exists and update it
	found := false
	for i, ref := range sectPr.FooterReferences {
		if ref.Type == string(footerType) {
			sectPr.FooterReferences[i] = footerRef
			found = true
			break
		}
	}
	if !found {
		sectPr.FooterReferences = append(sectPr.FooterReferences, footerRef)
	}
	
	fmt.Printf("[addFooterReference] Added footer reference: type=%s, id=%s, total refs=%d\n", 
		footerType, footerID, len(sectPr.FooterReferences))
}

// getSectionPropertiesForHeaderFooter 获取或创建带页眉页脚支持的节属性
func (d *Document) getSectionPropertiesForHeaderFooter() *SectionProperties {
	// 查找文档中最后一个节属性（MarshalXML uses the last one）
	var lastSectPr *SectionProperties
	var lastIndex int = -1
	
	for i, element := range d.Body.Elements {
		if sectPr, ok := element.(*SectionProperties); ok {
			lastSectPr = sectPr
			lastIndex = i
		}
	}
	
	if lastSectPr != nil {
		// 确保设置了关系命名空间
		if lastSectPr.XmlnsR == "" {
			lastSectPr.XmlnsR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
		}
		fmt.Printf("[getSectionPropertiesForHeaderFooter] Found existing sectPr at index %d\n", lastIndex)
		return lastSectPr
	}

	// 如果不存在，创建新的节属性
	fmt.Printf("[getSectionPropertiesForHeaderFooter] Creating new sectPr (total elements: %d)\n", len(d.Body.Elements))
	sectPr := &SectionProperties{
		XMLName: xml.Name{Local: "w:sectPr"},
		XmlnsR:  "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
		PageNumType: &PageNumType{
			Fmt: "decimal",
		},
		Columns: &Columns{
			Space: "720",
			Num:   "1",
		},
	}
	d.Body.Elements = append(d.Body.Elements, sectPr)
	return sectPr
}

// addContentType 添加内容类型
func (d *Document) addContentType(partName, contentType string) {
	// 检查是否已存在
	for _, override := range d.contentTypes.Overrides {
		if override.PartName == "/"+partName {
			return
		}
	}

	// 添加新的内容类型覆盖
	override := Override{
		PartName:    "/" + partName,
		ContentType: contentType,
	}
	d.contentTypes.Overrides = append(d.contentTypes.Overrides, override)
}

// CreateCustomFooter creates a footer with text on the left and page number on the right.
// This is useful for report footers like "INFORME DE EMISIONES 2024" with page numbers.
func (d *Document) CreateCustomFooter(footerType HeaderFooterType, leftText string, fontName string, fontSize int) error {
	fileName := getFileNameForType("footer", footerType)
	footerPartName := fmt.Sprintf("word/%s", fileName)
	
	// Create footer XML with left text and right-aligned page number using a table for layout
	footerXML := createCustomFooterXML(leftText, fontName, fontSize)
	
	d.parts[footerPartName] = []byte(footerXML)
	
	// Generate relationship ID
	footerID := fmt.Sprintf("rId%d", len(d.documentRelationships.Relationships)+2)
	
	// Check if relationship already exists
	for _, rel := range d.documentRelationships.Relationships {
		if rel.Target == fileName {
			footerID = rel.ID
			break
		}
	}
	
	// Add relationship if it doesn't exist
	exists := false
	for _, rel := range d.documentRelationships.Relationships {
		if rel.Target == fileName {
			exists = true
			break
		}
	}
	if !exists {
		relationship := Relationship{
			ID:     footerID,
			Type:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer",
			Target: fileName,
		}
		d.documentRelationships.Relationships = append(d.documentRelationships.Relationships, relationship)
	}
	
	// Add content type
	d.addContentType(footerPartName, "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml")
	
	// Add footer reference to section properties
	d.addFooterReference(footerType, footerID)
	
	fmt.Printf("[CreateCustomFooter] Created footer %s with text '%s'\n", fileName, leftText)
	return nil
}

// createCustomFooterXML creates the XML for a footer with left text and right page number
func createCustomFooterXML(leftText string, fontName string, fontSize int) string {
	if fontName == "" {
		fontName = "Aptos"
	}
	if fontSize == 0 {
		fontSize = 10
	}
	fontSizeHalfPt := fontSize * 2 // Word uses half-points
	
	return fmt.Sprintf(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
       xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
       xmlns:o="urn:schemas-microsoft-com:office:office"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
       xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
       xmlns:v="urn:schemas-microsoft-com:vml"
       xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
       xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
       xmlns:w10="urn:schemas-microsoft-com:office:word"
       xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
       xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
       xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
       xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
       xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk"
       xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"
       xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
       mc:Ignorable="w14 w15 wp14">
  <w:p>
    <w:pPr>
      <w:pStyle w:val="Footer"/>
      <w:tabs>
        <w:tab w:val="right" w:pos="9072"/>
      </w:tabs>
    </w:pPr>
    <w:r>
      <w:rPr>
        <w:rFonts w:ascii="%s" w:hAnsi="%s"/>
        <w:sz w:val="%d"/>
        <w:szCs w:val="%d"/>
        <w:color w:val="000000"/>
      </w:rPr>
      <w:t>%s</w:t>
    </w:r>
    <w:r>
      <w:rPr>
        <w:rFonts w:ascii="%s" w:hAnsi="%s"/>
        <w:sz w:val="%d"/>
        <w:szCs w:val="%d"/>
        <w:color w:val="000000"/>
      </w:rPr>
      <w:tab/>
    </w:r>
    <w:r>
      <w:rPr>
        <w:rFonts w:ascii="%s" w:hAnsi="%s"/>
        <w:sz w:val="%d"/>
        <w:szCs w:val="%d"/>
        <w:color w:val="000000"/>
      </w:rPr>
      <w:fldChar w:fldCharType="begin"/>
    </w:r>
    <w:r>
      <w:rPr>
        <w:rFonts w:ascii="%s" w:hAnsi="%s"/>
        <w:sz w:val="%d"/>
        <w:szCs w:val="%d"/>
        <w:color w:val="000000"/>
      </w:rPr>
      <w:instrText xml:space="preserve"> PAGE </w:instrText>
    </w:r>
    <w:r>
      <w:rPr>
        <w:rFonts w:ascii="%s" w:hAnsi="%s"/>
        <w:sz w:val="%d"/>
        <w:szCs w:val="%d"/>
        <w:color w:val="000000"/>
      </w:rPr>
      <w:fldChar w:fldCharType="separate"/>
    </w:r>
    <w:r>
      <w:rPr>
        <w:rFonts w:ascii="%s" w:hAnsi="%s"/>
        <w:sz w:val="%d"/>
        <w:szCs w:val="%d"/>
        <w:color w:val="000000"/>
      </w:rPr>
      <w:t>1</w:t>
    </w:r>
    <w:r>
      <w:rPr>
        <w:rFonts w:ascii="%s" w:hAnsi="%s"/>
        <w:sz w:val="%d"/>
        <w:szCs w:val="%d"/>
        <w:color w:val="000000"/>
      </w:rPr>
      <w:fldChar w:fldCharType="end"/>
    </w:r>
  </w:p>
</w:ftr>`,
		fontName, fontName, fontSizeHalfPt, fontSizeHalfPt, leftText,
		fontName, fontName, fontSizeHalfPt, fontSizeHalfPt,
		fontName, fontName, fontSizeHalfPt, fontSizeHalfPt,
		fontName, fontName, fontSizeHalfPt, fontSizeHalfPt,
		fontName, fontName, fontSizeHalfPt, fontSizeHalfPt,
		fontName, fontName, fontSizeHalfPt, fontSizeHalfPt,
		fontName, fontName, fontSizeHalfPt, fontSizeHalfPt)
}

// CreateHeaderWithImage creates a header with an image (typically a logo).
// The imageData should be the raw bytes of the image file.
// widthMM and heightMM specify the image dimensions in millimeters.
func (d *Document) CreateHeaderWithImage(headerType HeaderFooterType, imageData []byte, widthMM, heightMM float64) error {
	fileName := getFileNameForType("header", headerType)
	headerPartName := fmt.Sprintf("word/%s", fileName)
	
	// Add the image to the document
	imageNum := d.getNextImageNumber()
	imageName := fmt.Sprintf("media/image%d.png", imageNum)
	imagePartName := "word/" + imageName
	
	// Store the image data
	d.parts[imagePartName] = imageData
	
	// Add image content type
	d.addContentType(imagePartName, "image/png")
	
	// Create header relationship file for the image
	headerRelsPath := fmt.Sprintf("word/_rels/%s.rels", fileName)
	imageRId := "rId1"
	headerRelsXML := fmt.Sprintf(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="%s"/>
</Relationships>`, imageRId, imageName)
	d.parts[headerRelsPath] = []byte(headerRelsXML)
	
	// Convert mm to EMUs (English Metric Units) - 1 inch = 914400 EMUs, 1 inch = 25.4 mm
	widthEMU := int64(widthMM * 914400 / 25.4)
	heightEMU := int64(heightMM * 914400 / 25.4)
	
	// Create header XML with the image
	headerXML := createHeaderWithImageXML(imageRId, widthEMU, heightEMU, imageName)
	d.parts[headerPartName] = []byte(headerXML)
	
	// Generate relationship ID for the header
	headerID := fmt.Sprintf("rId%d", len(d.documentRelationships.Relationships)+2)
	
	// Check if relationship already exists
	for _, rel := range d.documentRelationships.Relationships {
		if rel.Target == fileName {
			headerID = rel.ID
			break
		}
	}
	
	// Add relationship if it doesn't exist
	exists := false
	for _, rel := range d.documentRelationships.Relationships {
		if rel.Target == fileName {
			exists = true
			break
		}
	}
	if !exists {
		relationship := Relationship{
			ID:     headerID,
			Type:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header",
			Target: fileName,
		}
		d.documentRelationships.Relationships = append(d.documentRelationships.Relationships, relationship)
	}
	
	// Add content type for header
	d.addContentType(headerPartName, "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml")
	
	// Add header reference to section properties
	d.addHeaderReference(headerType, headerID)
	
	fmt.Printf("[CreateHeaderWithImage] Created header %s with image (%.1fx%.1f mm)\n", fileName, widthMM, heightMM)
	return nil
}

// getNextImageNumber returns the next available image number
func (d *Document) getNextImageNumber() int {
	maxNum := 0
	for partName := range d.parts {
		if strings.HasPrefix(partName, "word/media/image") {
			var num int
			if _, err := fmt.Sscanf(partName, "word/media/image%d", &num); err == nil {
				if num > maxNum {
					maxNum = num
				}
			}
		}
	}
	return maxNum + 1
}

// createHeaderWithImageXML creates the XML for a header with an embedded image
func createHeaderWithImageXML(imageRId string, widthEMU, heightEMU int64, imageName string) string {
	return fmt.Sprintf(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
       xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
       xmlns:o="urn:schemas-microsoft-com:office:office"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
       xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
       xmlns:v="urn:schemas-microsoft-com:vml"
       xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
       xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
       xmlns:w10="urn:schemas-microsoft-com:office:word"
       xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
       xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
       xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
       xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
       xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk"
       xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"
       xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
       mc:Ignorable="w14 w15 wp14">
  <w:p>
    <w:pPr>
      <w:pStyle w:val="Header"/>
    </w:pPr>
    <w:r>
      <w:drawing>
        <wp:inline distT="0" distB="0" distL="0" distR="0">
          <wp:extent cx="%d" cy="%d"/>
          <wp:effectExtent l="0" t="0" r="0" b="0"/>
          <wp:docPr id="1" name="Logo"/>
          <wp:cNvGraphicFramePr>
            <a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>
          </wp:cNvGraphicFramePr>
          <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
              <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                <pic:nvPicPr>
                  <pic:cNvPr id="1" name="%s"/>
                  <pic:cNvPicPr/>
                </pic:nvPicPr>
                <pic:blipFill>
                  <a:blip r:embed="%s"/>
                  <a:stretch>
                    <a:fillRect/>
                  </a:stretch>
                </pic:blipFill>
                <pic:spPr>
                  <a:xfrm>
                    <a:off x="0" y="0"/>
                    <a:ext cx="%d" cy="%d"/>
                  </a:xfrm>
                  <a:prstGeom prst="rect">
                    <a:avLst/>
                  </a:prstGeom>
                </pic:spPr>
              </pic:pic>
            </a:graphicData>
          </a:graphic>
        </wp:inline>
      </w:drawing>
    </w:r>
  </w:p>
</w:hdr>`, widthEMU, heightEMU, imageName, imageRId, widthEMU, heightEMU)
}

// CreateEmptyHeaderFooter creates an empty header or footer (useful for first page with no header/footer)
func (d *Document) CreateEmptyHeaderFooter(isHeader bool, hfType HeaderFooterType) error {
	var fileName, partName, contentType, relType string
	
	if isHeader {
		fileName = getFileNameForType("header", hfType)
		partName = fmt.Sprintf("word/%s", fileName)
		contentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"
		relType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
	} else {
		fileName = getFileNameForType("footer", hfType)
		partName = fmt.Sprintf("word/%s", fileName)
		contentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"
		relType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"
	}
	
	// Create empty header/footer XML
	var emptyXML string
	if isHeader {
		emptyXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:pPr>
      <w:pStyle w:val="Header"/>
    </w:pPr>
  </w:p>
</w:hdr>`
	} else {
		emptyXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:pPr>
      <w:pStyle w:val="Footer"/>
    </w:pPr>
  </w:p>
</w:ftr>`
	}
	
	d.parts[partName] = []byte(emptyXML)
	
	// Generate relationship ID
	hfID := fmt.Sprintf("rId%d", len(d.documentRelationships.Relationships)+2)
	
	// Check if relationship already exists
	for _, rel := range d.documentRelationships.Relationships {
		if rel.Target == fileName {
			hfID = rel.ID
			break
		}
	}
	
	// Add relationship if it doesn't exist
	exists := false
	for _, rel := range d.documentRelationships.Relationships {
		if rel.Target == fileName {
			exists = true
			break
		}
	}
	if !exists {
		relationship := Relationship{
			ID:     hfID,
			Type:   relType,
			Target: fileName,
		}
		d.documentRelationships.Relationships = append(d.documentRelationships.Relationships, relationship)
	}
	
	// Add content type
	d.addContentType(partName, contentType)
	
	// Add reference to section properties
	if isHeader {
		d.addHeaderReference(hfType, hfID)
	} else {
		d.addFooterReference(hfType, hfID)
	}
	
	return nil
}
