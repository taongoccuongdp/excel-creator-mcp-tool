package main

import (
	"context"
	"errors"
	"fmt"
	"log"
	"net/http"
	"os"
	"path/filepath"
	"strings"

	"github.com/modelcontextprotocol/go-sdk/mcp"
)

func main() {
	const serviceName = "excel-creator"

	server := mcp.NewServer(&mcp.Implementation{
		Name:    serviceName,
		Version: "0.3.0",
	}, nil)

	tool := &mcp.Tool{
		Name:         "create_excel_file",
		Description:  "Create an Excel .xlsx file from flat parameters, upload it, and return the download URL.",
		InputSchema:  buildInputSchema(),
		OutputSchema: buildOutputSchema(),
	}
	mcp.AddTool(server, tool, createExcelFile)

	addr := envOrDefault("MCP_ADDR", ":8080")

	if os.Getenv("MCP_STDIO") == "1" {
		log.Println("running in stdio mode")
		if err := server.Run(context.Background(), &mcp.StdioTransport{}); err != nil {
			log.Fatalf("server failed: %v", err)
		}
		return
	}

	mcpHandler := mcp.NewStreamableHTTPHandler(func(*http.Request) *mcp.Server {
		return server
	}, &mcp.StreamableHTTPOptions{
		Stateless:    true,
		JSONResponse: true,
	})
	sseHandler := mcp.NewSSEHandler(func(*http.Request) *mcp.Server {
		return server
	}, nil)

	mux := http.NewServeMux()
	mux.Handle("/mcp", mcpHandler)
	mux.Handle("/sse", sseHandler)
	mux.HandleFunc("/healthz", healthHandler(serviceName))
	mux.HandleFunc("/", func(w http.ResponseWriter, r *http.Request) {
		if r.URL.Path != "/" {
			http.NotFound(w, r)
			return
		}
		if r.Method == http.MethodGet || r.Method == http.MethodHead {
			healthHandler(serviceName)(w, r)
			return
		}
		mcpHandler.ServeHTTP(w, r)
	})

	log.Printf("MCP server listening on %s", addr)
	if err := http.ListenAndServe(addr, mux); err != nil {
		log.Fatalf("server failed: %v", err)
	}
}

func resolvePathInsideBase(baseDir string, requested string, requiredExt string) (string, error) {
	if strings.TrimSpace(requested) == "" {
		return "", errors.New("path is required")
	}

	cleaned := filepath.Clean(requested)
	if requiredExt != "" && !strings.EqualFold(filepath.Ext(cleaned), requiredExt) {
		return "", fmt.Errorf("path must end with %s", requiredExt)
	}

	absBase, err := filepath.Abs(baseDir)
	if err != nil {
		return "", fmt.Errorf("resolve base dir: %w", err)
	}

	var absTarget string
	if filepath.IsAbs(cleaned) {
		absTarget = cleaned
	} else {
		absTarget = filepath.Join(absBase, cleaned)
	}

	absTarget, err = filepath.Abs(absTarget)
	if err != nil {
		return "", fmt.Errorf("resolve target path: %w", err)
	}

	rel, err := filepath.Rel(absBase, absTarget)
	if err != nil {
		return "", fmt.Errorf("check target path: %w", err)
	}
	if rel == ".." || strings.HasPrefix(rel, ".."+string(os.PathSeparator)) {
		return "", fmt.Errorf("path escapes allowed directory: %s", requested)
	}

	return absTarget, nil
}

func envOrDefault(key, fallback string) string {
	if v := strings.TrimSpace(os.Getenv(key)); v != "" {
		return v
	}
	return fallback
}

func healthHandler(service string) http.HandlerFunc {
	return func(w http.ResponseWriter, r *http.Request) {
		w.Header().Set("Content-Type", "application/json")
		w.WriteHeader(http.StatusOK)
		if r.Method != http.MethodHead {
			_, _ = fmt.Fprintf(w, `{"status":"ok","service":"%s"}`, service)
		}
	}
}

func ptr[T any](v T) *T {
	return &v
}
