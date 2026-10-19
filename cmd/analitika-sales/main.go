package main

import (
	"context"
	"flag"
	"fmt"
	"net"
	"net/http"
	"os"
	"os/exec"
	"runtime"
	"time"

	"github.com/slchumak75-sketch/analitika-sales/internal/app"
)

func main() {
	address := flag.String("addr", "127.0.0.1:0", "локальный адрес приложения")
	noBrowser := flag.Bool("no-browser", false, "не открывать браузер автоматически (для разработки)")
	flag.Parse()

	if err := run(*address, !*noBrowser); err != nil {
		fmt.Fprintln(os.Stderr, "Аналитика продаж:", err)
		os.Exit(1)
	}
}

func run(address string, openBrowser bool) error {
	listener, err := net.Listen("tcp", address)
	if err != nil {
		return fmt.Errorf("не удалось запустить локальный сервер: %w", err)
	}

	shutdownRequested := make(chan struct{}, 1)
	server := &http.Server{
		Handler:           app.NewHandler(func() { signalShutdown(shutdownRequested) }),
		ReadHeaderTimeout: 10 * time.Second,
		ReadTimeout:       3 * time.Minute,
		WriteTimeout:      3 * time.Minute,
		IdleTimeout:       60 * time.Second,
	}
	serveResult := make(chan error, 1)
	go func() {
		serveResult <- server.Serve(listener)
	}()

	url := "http://" + listener.Addr().String() + "/"
	if openBrowser {
		time.AfterFunc(250*time.Millisecond, func() {
			_ = launchBrowser(url)
		})
	} else {
		fmt.Printf("Откройте интерфейс: %s\n", url)
	}

	select {
	case err := <-serveResult:
		if err != nil && err != http.ErrServerClosed {
			return err
		}
		return nil
	case <-shutdownRequested:
		ctx, cancel := context.WithTimeout(context.Background(), 10*time.Second)
		defer cancel()
		if err := server.Shutdown(ctx); err != nil {
			_ = server.Close()
			return err
		}
		return nil
	}
}

func signalShutdown(channel chan<- struct{}) {
	select {
	case channel <- struct{}{}:
	default:
	}
}

func launchBrowser(url string) error {
	var command *exec.Cmd
	switch runtime.GOOS {
	case "windows":
		command = exec.Command("rundll32", "url.dll,FileProtocolHandler", url)
	case "darwin":
		command = exec.Command("open", url)
	default:
		command = exec.Command("xdg-open", url)
	}
	if err := command.Start(); err != nil {
		return err
	}
	go func() { _ = command.Wait() }()
	return nil
}
