package app

import (
	_ "embed"
	"encoding/json"
	"errors"
	"io"
	"net/http"
	"path/filepath"
	"strings"

	"github.com/slchumak75-sketch/analitika-sales/internal/analysis"
	"github.com/slchumak75-sketch/analitika-sales/internal/xlsx"
)

const (
	maxUploadSize  = 50 << 20
	maxRequestSize = maxUploadSize + (1 << 20) // Multipart boundaries and headers.
)

//go:embed web/index.html
var page []byte

type analysisResponse struct {
	FileName string          `json:"fileName"`
	Report   analysis.Report `json:"report"`
}

type errorResponse struct {
	Error string `json:"error"`
}

// NewHandler returns the local HTTP interface used by the desktop executable.
// The application listens on loopback and does not transmit workbook data.
func NewHandler(shutdown func()) http.Handler {
	mux := http.NewServeMux()
	mux.HandleFunc("/", func(w http.ResponseWriter, r *http.Request) {
		if r.Method != http.MethodGet || r.URL.Path != "/" {
			http.NotFound(w, r)
			return
		}
		w.Header().Set("Content-Type", "text/html; charset=utf-8")
		w.Header().Set("Cache-Control", "no-store")
		_, _ = w.Write(page)
	})
	mux.HandleFunc("/api/health", func(w http.ResponseWriter, r *http.Request) {
		if r.Method != http.MethodGet {
			writeError(w, http.StatusMethodNotAllowed, "Метод не поддерживается")
			return
		}
		w.Header().Set("Content-Type", "application/json; charset=utf-8")
		_ = json.NewEncoder(w).Encode(map[string]string{"status": "ok"})
	})
	mux.HandleFunc("/api/analyze", analyzeUpload)
	mux.HandleFunc("/api/shutdown", func(w http.ResponseWriter, r *http.Request) {
		if r.Method != http.MethodPost {
			writeError(w, http.StatusMethodNotAllowed, "Метод не поддерживается")
			return
		}
		w.Header().Set("Content-Type", "application/json; charset=utf-8")
		_ = json.NewEncoder(w).Encode(map[string]string{"status": "closing"})
		if shutdown != nil {
			go shutdown()
		}
	})

	return secureHeaders(mux)
}

func analyzeUpload(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodPost {
		writeError(w, http.StatusMethodNotAllowed, "Метод не поддерживается")
		return
	}
	r.Body = http.MaxBytesReader(w, r.Body, maxRequestSize)
	if err := r.ParseMultipartForm(maxRequestSize); err != nil {
		if errors.Is(err, http.ErrNotMultipart) {
			writeError(w, http.StatusBadRequest, "Отправьте файл через форму загрузки")
			return
		}
		var maxBytesError *http.MaxBytesError
		if errors.As(err, &maxBytesError) {
			writeError(w, http.StatusRequestEntityTooLarge, "Размер файла превышает лимит 50 МиБ")
			return
		}
		writeError(w, http.StatusBadRequest, "Не удалось прочитать загруженный файл")
		return
	}
	if r.MultipartForm != nil {
		defer r.MultipartForm.RemoveAll()
	}

	file, header, err := r.FormFile("file")
	if err != nil {
		writeError(w, http.StatusBadRequest, "Выберите Excel-файл для анализа")
		return
	}
	defer file.Close()

	fileName := filepath.Base(header.Filename)
	extension := strings.ToLower(filepath.Ext(fileName))
	if extension != ".xlsx" && extension != ".xlsm" {
		writeError(w, http.StatusBadRequest, "Поддерживаются книги .xlsx и .xlsm. Формат старых файлов .xls пока не поддерживается")
		return
	}
	data, err := io.ReadAll(io.LimitReader(file, maxUploadSize+1))
	if err != nil {
		writeError(w, http.StatusBadRequest, "Не удалось прочитать файл")
		return
	}
	if len(data) > maxUploadSize {
		writeError(w, http.StatusRequestEntityTooLarge, "Размер файла превышает лимит 50 МиБ")
		return
	}

	table, err := xlsx.ReadFirstTable(data)
	if err != nil {
		writeError(w, http.StatusUnprocessableEntity, err.Error())
		return
	}
	report, err := analysis.Analyze(table)
	if err != nil {
		writeError(w, http.StatusUnprocessableEntity, err.Error())
		return
	}

	w.Header().Set("Content-Type", "application/json; charset=utf-8")
	w.Header().Set("Cache-Control", "no-store")
	w.WriteHeader(http.StatusOK)
	if err := json.NewEncoder(w).Encode(analysisResponse{FileName: fileName, Report: report}); err != nil {
		return
	}
}

func secureHeaders(next http.Handler) http.Handler {
	return http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		w.Header().Set("X-Content-Type-Options", "nosniff")
		w.Header().Set("Referrer-Policy", "no-referrer")
		w.Header().Set("X-Frame-Options", "DENY")
		w.Header().Set("Content-Security-Policy", "default-src 'self'; connect-src 'self'; style-src 'self' 'unsafe-inline'; script-src 'self' 'unsafe-inline'; img-src 'self' data:; base-uri 'none'; frame-ancestors 'none'; form-action 'self'")
		next.ServeHTTP(w, r)
	})
}

func writeError(w http.ResponseWriter, status int, message string) {
	w.Header().Set("Content-Type", "application/json; charset=utf-8")
	w.Header().Set("Cache-Control", "no-store")
	w.WriteHeader(status)
	_ = json.NewEncoder(w).Encode(errorResponse{Error: message})
}
