package main

import (
	"bytes"
	"encoding/json"
	"fmt"
	"io"
	"mime/multipart"
	"net/http"
	"os"
	"path/filepath"
	"strings"
	"time"
)


func uploadExcelFile(localPath string, task string) (*UploadResponse, error) {
	file, err := os.Open(localPath)
	if err != nil {
		return nil, fmt.Errorf("open file for upload: %w", err)
	}
	defer file.Close()

	var body bytes.Buffer
	writer := multipart.NewWriter(&body)

	formFile, err := writer.CreateFormFile("file", filepath.Base(localPath))
	if err != nil {
		return nil, fmt.Errorf("create upload form file: %w", err)
	}
	if _, err := io.Copy(formFile, file); err != nil {
		return nil, fmt.Errorf("copy upload file: %w", err)
	}

	if err := writer.WriteField("date_in_url", "false"); err != nil {
		return nil, fmt.Errorf("write date_in_url field: %w", err)
	}
	if err := writer.WriteField("task", task); err != nil {
		return nil, fmt.Errorf("write task field: %w", err)
	}
	if err := writer.Close(); err != nil {
		return nil, fmt.Errorf("close multipart writer: %w", err)
	}

	upload_url := envOrDefault("EXCEL_UPLOAD_URL", "")
	if upload_url == "" {
		return nil, fmt.Errorf("Did not set url to upload excel file to `EXCEL_UPLOAD_URL`.")
	}

	req, err := http.NewRequest(http.MethodPost, upload_url, &body)
	if err != nil {
		return nil, fmt.Errorf("create upload request: %w", err)
	}
	req.Header.Set("Content-Type", writer.FormDataContentType())

	client := &http.Client{Timeout: 2 * time.Minute}
	resp, err := client.Do(req)
	if err != nil {
		return nil, fmt.Errorf("send upload request: %w", err)
	}
	defer resp.Body.Close()

	respBody, err := io.ReadAll(resp.Body)
	if err != nil {
		return nil, fmt.Errorf("read upload response: %w", err)
	}

	if resp.StatusCode < 200 || resp.StatusCode >= 300 {
		return nil, fmt.Errorf("upload request failed with status %d: %s", resp.StatusCode, strings.TrimSpace(string(respBody)))
	}

	var out UploadResponse
	if err := json.Unmarshal(respBody, &out); err != nil {
		return nil, fmt.Errorf("decode upload response: %w", err)
	}
	if out.Status != 1 || strings.TrimSpace(out.URL) == "" {
		return nil, fmt.Errorf("upload response is not successful: %s", strings.TrimSpace(string(respBody)))
	}

	return &out, nil
}
