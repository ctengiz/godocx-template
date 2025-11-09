package godocx

import (
	"bytes"
	"fmt"
	"log/slog"
	"slices"
)

// CreateReport generates a report document based on a given template and data.
// It parses the template file, processes any commands within the template
// using provided data, and outputs the final document as a byte slice.
//
// Parameters:
//   - templatePath: The file path to the template document.
//   - data: A pointer to ReportData containing data to be inserted into the template.
//
// Returns:
//   - A byte slice representing the generated document.
//   - An error if any occurs during template parsing, processing, or document generation.
func CreateReport(templatePath string, data *ReportData, options CreateReportOptions) ([]byte, error) {

	outBuffer := new(bytes.Buffer)
	zip, err := NewZipArchive(templatePath, outBuffer)
	if err != nil {
		return nil, err
	}

	// xml parse the document
	parseResult, err := ParseTemplate(zip)
	if err != nil {
		return nil, fmt.Errorf("ParseTemplate failed: %w", err)
	}

	if options.CmdDelimiter == nil {
		options.CmdDelimiter = &Delimiters{
			Open:  DEFAULT_CMD_DELIMITER,
			Close: DEFAULT_CMD_DELIMITER,
		}
	}
	if options.LiteralXmlDelimiter == "" {
		options.LiteralXmlDelimiter = DEFAULT_LITERAL_XML_DELIMITER
	}

	xmlOptions := XmlOptions{
		LiteralXmlDelimiter: options.LiteralXmlDelimiter,
	}

	preppedTemplate, err := PreprocessTemplate(parseResult.Root, *options.CmdDelimiter)
	if err != nil {
		return nil, fmt.Errorf("PreprocessTemplate failed: %w", err)
	}

	result, err := ProduceReport(data, preppedTemplate, NewContext(options, 73086257))
	//TODO ^ max id
	if err != nil {
		return nil, fmt.Errorf("ProduceReport failed: %w", err)
	}

	newXml := BuildXml(result.Report, xmlOptions, "")

	slog.Debug("Writing report...")
	zip.SetFile("word/document.xml", newXml)

	numImages := len(result.Images)
	numHtmls := len(result.Htmls)
	err = ProcessImages(result.Images, parseResult.MainDocument, parseResult.Zip)
	if err != nil {
		return nil, fmt.Errorf("ProcessImages failed: %w", err)
	}
	err = ProcessHtmls(result.Htmls, parseResult.MainDocument, parseResult.Zip)
	if err != nil {
		return nil, fmt.Errorf("ProcessHtmls failed: %w", err)
	}
	err = ProcessLinks(result.Links, parseResult.MainDocument, parseResult.Zip)
	if err != nil {
		return nil, fmt.Errorf("ProcessLinks failed: %w", err)
	}

	// Additionals headers and footers
	for extraPath, extraNode := range parseResult.Extras {
		prepped, err := PreprocessTemplate(extraNode, *options.CmdDelimiter)
		if err != nil {
			return nil, fmt.Errorf("PreprocessTemplate failed: %w", err)
		}
		r, err := ProduceReport(data, prepped, NewContext(options, 73086257))
		if err != nil {
			return nil, fmt.Errorf("ProduceReport failed: %w", err)
		}
		extraXml := BuildXml(r.Report, xmlOptions, "")
		slog.Debug(fmt.Sprintf("Writing %s...", extraPath))
		zip.SetFile(extraPath, extraXml)
	}

	if numHtmls > 0 || numImages > 0 {
		slog.Debug("Completing [Content_Types].xml...")

		contentTypes := parseResult.ContentTypes
		children := contentTypes.Children()
		ensureContentType := func(extension string, contentType string) {
			containsExtension := slices.ContainsFunc(children, func(n Node) bool {
				nonTextNode, isNonTextNode := n.(*NonTextNode)
				return isNonTextNode && nonTextNode.Attrs["Extension"] == extension
			})
			if containsExtension {
				return
			}
			AddChild(contentTypes, NewNonTextNode("Default", map[string]string{"Extension": extension, "ContentType": contentType}, nil))
		}
		if numImages > 0 {
			slog.Debug("Completing [Content_Types].xml for IMAGES...")
			ensureContentType("png", "image/png")
			ensureContentType("jpg", "image/jpeg")
			ensureContentType("jpeg", "image/jpeg")
			ensureContentType("gif", "image/gif")
			ensureContentType("bmp", "image/bmp")
			ensureContentType("svg", "image/svg+xml")
		}
		if numHtmls > 0 {
			slog.Debug("Completing [Content_Types].xml for HTML...")
			ensureContentType("html", "text/html")
		}
		finalContentTypesXml := BuildXml(parseResult.ContentTypes, xmlOptions, "")
		zip.SetFile(CONTENT_TYPES_PATH, finalContentTypesXml)
	}

	err = zip.Close()
	if err != nil {
		return nil, fmt.Errorf("Error closing zip : %w", err)
	}
	return outBuffer.Bytes(), nil
}
