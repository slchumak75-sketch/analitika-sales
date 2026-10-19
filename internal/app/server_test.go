package app

import (
	"bytes"
	"encoding/json"
	"math"
	"mime/multipart"
	"net/http"
	"net/http/httptest"
	"os"
	"path/filepath"
	"strings"
	"testing"
	"time"

	"github.com/slchumak75-sketch/analitika-sales/internal/analysis"
)

func TestRootAndHealthEndpoints(t *testing.T) {
	handler := NewHandler(nil)

	root := httptest.NewRecorder()
	handler.ServeHTTP(root, httptest.NewRequest(http.MethodGet, "/", nil))
	if root.Code != http.StatusOK {
		t.Fatalf("GET / status = %d, want %d", root.Code, http.StatusOK)
	}
	if !strings.Contains(root.Header().Get("Content-Type"), "text/html") || !strings.Contains(root.Body.String(), "Аналитика продаж") {
		t.Fatal("root endpoint did not return the embedded HTML interface")
	}
	if root.Header().Get("X-Content-Type-Options") != "nosniff" {
		t.Fatal("security headers were not added")
	}

	health := httptest.NewRecorder()
	handler.ServeHTTP(health, httptest.NewRequest(http.MethodGet, "/api/health", nil))
	if health.Code != http.StatusOK || !strings.Contains(health.Body.String(), `"status":"ok"`) {
		t.Fatalf("health response = %d %q", health.Code, health.Body.String())
	}
}

func TestShutdownEndpointInvokesCallback(t *testing.T) {
	called := make(chan struct{}, 1)
	handler := NewHandler(func() { called <- struct{}{} })
	recorder := httptest.NewRecorder()
	handler.ServeHTTP(recorder, httptest.NewRequest(http.MethodPost, "/api/shutdown", nil))
	if recorder.Code != http.StatusOK {
		t.Fatalf("shutdown response status = %d, want 200", recorder.Code)
	}
	select {
	case <-called:
	case <-time.After(time.Second):
		t.Fatal("shutdown callback was not invoked")
	}
}

func TestAnalyzeUploadRepositoryWorkbook(t *testing.T) {
	data, err := os.ReadFile(filepath.Join("..", "..", "ОТЧЕТ ПО ОРГАНИЗАЦИИ мир.xlsx"))
	if err != nil {
		t.Skipf("repository workbook is not available: %v", err)
	}

	var body bytes.Buffer
	writer := multipart.NewWriter(&body)
	part, err := writer.CreateFormFile("file", "ОТЧЕТ ПО ОРГАНИЗАЦИИ мир.xlsx")
	if err != nil {
		t.Fatalf("create form file: %v", err)
	}
	if _, err := part.Write(data); err != nil {
		t.Fatalf("write form file: %v", err)
	}
	if err := writer.Close(); err != nil {
		t.Fatalf("close multipart writer: %v", err)
	}

	request := httptest.NewRequest(http.MethodPost, "/api/analyze", &body)
	request.Header.Set("Content-Type", writer.FormDataContentType())
	recorder := httptest.NewRecorder()
	NewHandler(nil).ServeHTTP(recorder, request)
	if recorder.Code != http.StatusOK {
		t.Fatalf("analyze status = %d, body = %s", recorder.Code, recorder.Body.String())
	}

	var payload struct {
		FileName string          `json:"fileName"`
		Report   analysis.Report `json:"report"`
	}
	if err := json.Unmarshal(recorder.Body.Bytes(), &payload); err != nil {
		t.Fatalf("decode response: %v", err)
	}
	if payload.FileName == "" || payload.Report.Mode != analysis.ModeReport {
		t.Fatalf("unexpected response: file=%q mode=%q", payload.FileName, payload.Report.Mode)
	}
	if payload.Report.NetProfit == nil || math.Abs(*payload.Report.NetProfit-93362.63733333332) > 0.01 {
		t.Fatalf("unexpected net profit: %v", payload.Report.NetProfit)
	}
	if payload.Report.ProfitTotal == nil || math.Abs(*payload.Report.ProfitTotal-116703.29666666663) > 0.01 {
		t.Fatalf("unexpected profit before tax: %v", payload.Report.ProfitTotal)
	}
}
