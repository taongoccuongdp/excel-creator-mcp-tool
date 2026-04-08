package main

import (
	"encoding/json"
	"fmt"

	"github.com/xuri/excelize/v2"
)

func resolveStyleID(
	f *excelize.File,
	styles map[string]StyleSpec,
	styleRef string,
	styleCache map[string]int,
) (int, error) {
	spec, ok := styles[styleRef]
	if !ok {
		return 0, fmt.Errorf("unknown style_ref: %s", styleRef)
	}

	keyBytes, err := json.Marshal(spec)
	if err != nil {
		return 0, fmt.Errorf("marshal style %s: %w", styleRef, err)
	}
	key := string(keyBytes)

	if id, ok := styleCache[key]; ok {
		return id, nil
	}

	style, err := toExcelizeStyle(spec)
	if err != nil {
		return 0, err
	}

	id, err := f.NewStyle(style)
	if err != nil {
		return 0, fmt.Errorf("create style %s: %w", styleRef, err)
	}

	styleCache[key] = id
	return id, nil
}
