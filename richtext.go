package main

import (
	stdhtml "html"
	"io"
	"strings"

	"github.com/xuri/excelize/v2"
	xhtml "golang.org/x/net/html"
)

func parseHTMLBoldRichText(value any) ([]excelize.RichTextRun, bool, error) {
	text, ok := value.(string)
	if !ok {
		return nil, false, nil
	}

	if !containsBoldHTML(text) {
		return nil, false, nil
	}

	tokenizer := xhtml.NewTokenizerFragment(strings.NewReader(text), "span")
	runs := make([]excelize.RichTextRun, 0, 4)
	boldDepth := 0

	for {
		switch tokenizer.Next() {
		case xhtml.ErrorToken:
			if err := tokenizer.Err(); err != nil && err != io.EOF {
				return nil, false, err
			}
			return compactRichTextRuns(runs), true, nil
		case xhtml.TextToken:
			tokenText := stdhtml.UnescapeString(string(tokenizer.Text()))
			if tokenText == "" {
				continue
			}

			run := excelize.RichTextRun{Text: tokenText}
			if boldDepth > 0 {
				run.Font = &excelize.Font{Bold: true}
			}
			runs = append(runs, run)
		case xhtml.StartTagToken, xhtml.EndTagToken, xhtml.SelfClosingTagToken:
			tagName, _ := tokenizer.TagName()
			switch strings.ToLower(string(tagName)) {
			case "b", "strong":
				if tokenizer.Token().Type == xhtml.EndTagToken {
					if boldDepth > 0 {
						boldDepth--
					}
				} else {
					boldDepth++
				}
			case "br":
				if tokenizer.Token().Type != xhtml.EndTagToken {
					runs = append(runs, excelize.RichTextRun{Text: "\n"})
				}
			}
		}
	}
}

func containsBoldHTML(text string) bool {
	lower := strings.ToLower(text)
	return strings.Contains(lower, "<b>") ||
		strings.Contains(lower, "</b>") ||
		strings.Contains(lower, "<strong>") ||
		strings.Contains(lower, "</strong>")
}

func compactRichTextRuns(runs []excelize.RichTextRun) []excelize.RichTextRun {
	if len(runs) < 2 {
		return runs
	}

	out := make([]excelize.RichTextRun, 0, len(runs))
	for _, run := range runs {
		if len(out) == 0 {
			out = append(out, run)
			continue
		}

		last := &out[len(out)-1]
		if sameRichTextFont(last.Font, run.Font) {
			last.Text += run.Text
			continue
		}

		out = append(out, run)
	}

	return out
}

func sameRichTextFont(a, b *excelize.Font) bool {
	if a == nil || b == nil {
		return a == nil && b == nil
	}

	return a.Bold == b.Bold &&
		a.Italic == b.Italic &&
		a.Underline == b.Underline &&
		a.Family == b.Family &&
		a.Size == b.Size &&
		a.Strike == b.Strike &&
		a.Color == b.Color &&
		a.ColorIndexed == b.ColorIndexed &&
		a.ColorTint == b.ColorTint &&
		a.VertAlign == b.VertAlign
}
