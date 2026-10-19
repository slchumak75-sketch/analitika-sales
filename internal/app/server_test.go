package app

import (
	"net/http"
	"net/http/httptest"
	"strings"
	"testing"
	"time"
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
